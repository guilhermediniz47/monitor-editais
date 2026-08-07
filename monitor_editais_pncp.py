#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
monitor_editais_pncp.py  (v4.2 - resiliente a falhas do PNCP)
======================================================================
Monitor de editais do PNCP (UF=PR).

COMPORTAMENTO AO SER EXECUTADO (dois cliques no .exe, sem argumentos):
  - Le a WATCHLIST da planilha 'watchlist_editais.xlsx' (aba 'Editais'),
    na MESMA PASTA do executavel.
  - Consulta o PNCP puxando editais publicados/atualizados nos ULTIMOS
    7 DIAS ate HOJE (dia da execucao).
  - Detecta RETOMADA (suspenso -> ativo) e REPUBLICACAO (revogado/anulado
    que reaparece semelhante).
  - Envia alerta no TEAMS (INFRA / EDIFICACOES) se webhooks.env existir;
    caso contrario, apenas MOSTRA os alertas na tela (util em testes).
  - Mantem historico local 'editais_pncp.db' para comparar execucoes.

NOVIDADES v4.2:
  - Repeticao automatica (retry) com espera progressiva em caso de
    timeout / HTTP 500-503 / 429 do PNCP (o servidor oscila).
  - Timeout de leitura ampliado e pequena pausa entre paginas (educado).
  - Modalidades focadas em OBRAS (Concorrencia), sem Pregao Eletronico,
    que traz volume enorme de bens/servicos comuns e sobrecarrega.

Colunas da planilha (uma linha = um edital vigiado):
  Ente Licitante | Objeto | Valor Licitado | Gerencia | Observacoes

Conformidade com a API (Swagger oficial pncp.gov.br/api/consulta/v3/api-docs):
  Base https://pncp.gov.br/api/consulta ; GET /v1/contratacoes/publicacao
  e /v1/contratacoes/atualizacao ; datas yyyyMMdd ;
  codigoModalidadeContratacao (OBRIGATORIO) ; uf ; pagina (OBRIGATORIO) ;
  tamanhoPagina MAXIMO 50 nestes endpoints.
======================================================================
"""

import argparse
import datetime as dt
import os
import re
import sqlite3
import sys
import time
import unicodedata
from difflib import SequenceMatcher

try:
    import requests
except ImportError:
    sys.exit("Falta a biblioteca 'requests' (embutida no .exe).")

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Falta a biblioteca 'openpyxl' (embutida no .exe).")

try:
    from rapidfuzz import fuzz
    def obj_sim(a, b):
        return fuzz.token_set_ratio(a, b) / 100.0
except ImportError:
    def obj_sim(a, b):
        return SequenceMatcher(None, a, b).ratio()

try:
    import holidays
    _BR_HOLIDAYS = holidays.Brazil(subdiv="PR")
    def eh_dia_util(d):
        return d.weekday() < 5 and d not in _BR_HOLIDAYS
except ImportError:
    def eh_dia_util(d):
        return d.weekday() < 5


# ======================================================================
# LOCALIZACAO DOS ARQUIVOS (funciona como .py e como .exe do PyInstaller)
# ======================================================================
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _carrega_env():
    for nome in (".env", "webhooks.env", "config.env"):
        caminho = os.path.join(BASE_DIR, nome)
        if os.path.exists(caminho):
            try:
                with open(caminho, "r", encoding="utf-8") as fh:
                    for linha in fh:
                        linha = linha.strip()
                        if not linha or linha.startswith("#") or "=" not in linha:
                            continue
                        k, v = linha.split("=", 1)
                        os.environ.setdefault(k.strip(), v.strip())
            except OSError:
                pass
            break


_carrega_env()


# ======================================================================
# CONFIGURACAO
# ======================================================================
BASE = "https://pncp.gov.br/api/consulta"
UF_ALVO = "PR"
JANELA_DIAS = 7
TAM_PAGINA = 50          # MAXIMO permitido nos endpoints de contratacoes

# Resiliencia a instabilidade do PNCP:
TIMEOUT = 90             # segundos por requisicao (o servidor oscila)
TENTATIVAS = 4           # repeticoes por pagina antes de desistir
BACKOFF = 4              # espera base (s); cresce a cada tentativa
PAUSA_PAGINA = 0.4       # pausa entre paginas (educado com o servidor)

# Modalidades FOCADAS EM OBRAS/ENGENHARIA.
#   4 = Concorrencia - Eletronica   (modalidade tipica de obras)
#   5 = Concorrencia - Presencial
#   8 = Dispensa de Licitacao
#   9 = Inexigibilidade
# OBS.: o codigo 6 (Pregao Eletronico) foi REMOVIDO de proposito: traz um
# volume enorme de bens/servicos comuns (nao-obras) e sobrecarrega a coleta.
# Confirme os codigos na tabela "Modalidade de Contratacao" do manual do PNCP.
MODALIDADES_OBRAS = [4, 5, 8, 9]

XLSX = os.environ.get("WATCHLIST_XLSX",
                      os.path.join(BASE_DIR, "watchlist_editais.xlsx"))
ABA = os.environ.get("WATCHLIST_ABA", "Editais")
DB = os.environ.get("MONITOR_DB", os.path.join(BASE_DIR, "editais_pncp.db"))

WEBHOOKS = {
    "INFRA":       os.environ.get("TEAMS_WEBHOOK_INFRA", "COLE_A_URL_DO_WEBHOOK_INFRA"),
    "EDIFICACOES": os.environ.get("TEAMS_WEBHOOK_EDIF",  "COLE_A_URL_DO_WEBHOOK_EDIFICACOES"),
}
GERENTES = {
    "INFRA":       "Lincoln Santos de Andrade",
    "EDIFICACOES": "Alexandre Cardoso Dal Ross",
}

LIMIAR_OBJETO = 0.80
TOLERANCIA_VALOR = 0.15
SCORE_MINIMO = 0.70


# ======================================================================
# TEXTO / FINGERPRINT
# ======================================================================
STOPWORDS = {
    "de", "da", "do", "das", "dos", "e", "para", "com", "em", "no", "na",
    "a", "o", "as", "os", "um", "uma", "por", "que", "ao", "aos", "obra",
    "contratacao", "empresa", "especializada", "servicos", "execucao",
}


def normaliza(txt):
    if not txt:
        return ""
    txt = unicodedata.normalize("NFKD", str(txt))
    txt = "".join(c for c in txt if not unicodedata.combining(c)).lower()
    txt = re.sub(r"[^a-z0-9\s]", " ", txt)
    return " ".join(t for t in txt.split() if t not in STOPWORDS and len(t) > 2)


def so_digitos(s):
    return re.sub(r"\D", "", str(s or ""))


def norm_gerencia(g):
    g = normaliza(g).upper().replace(" ", "")
    if g.startswith("EDIF"):
        return "EDIFICACOES"
    if g.startswith("INFRA"):
        return "INFRA"
    return g


def casa_ente(cnpj_watch, nome_watch, cnpj_item, orgao_item):
    cw, ci = so_digitos(cnpj_watch), so_digitos(cnpj_item)
    if cw and ci and cw == ci:
        return 1.0
    if nome_watch and orgao_item:
        s = obj_sim(normaliza(nome_watch), normaliza(orgao_item))
        return 0.85 if s >= 0.80 else 0.0
    return 0.0


def score_match(watch, objeto_i, valor_i, cnpj_i, orgao_i):
    s_obj = obj_sim(normaliza(watch["objeto"]), normaliza(objeto_i))
    s_ente = casa_ente(watch["ente_cnpj"], watch["ente_nome"], cnpj_i, orgao_i)
    s_val = 0.0
    if watch["valor"] and valor_i:
        maior = max(watch["valor"], valor_i)
        dif = abs(watch["valor"] - valor_i) / maior if maior else 1.0
        s_val = 1.0 if dif <= TOLERANCIA_VALOR else max(0.0, 1 - dif)
    return 0.50 * s_obj + 0.30 * s_ente + 0.20 * s_val, s_obj, s_ente


# ======================================================================
# WATCHLIST A PARTIR DO EXCEL
# ======================================================================
def _to_float(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.-]", "", str(v)).replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def carrega_watchlist():
    if not os.path.exists(XLSX):
        print(f"[ERRO] Planilha nao encontrada: {XLSX}")
        print("       Coloque 'watchlist_editais.xlsx' na mesma pasta do programa.")
        return []
    wb = load_workbook(XLSX, data_only=True, read_only=True)
    if ABA not in wb.sheetnames:
        print(f"[ERRO] Aba '{ABA}' nao existe. Abas: {wb.sheetnames}")
        return []
    ws = wb[ABA]
    watch = []
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 3:
            continue
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        ente = str(row[0]).strip() if len(row) > 0 and row[0] is not None else ""
        objeto = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        valor = _to_float(row[2] if len(row) > 2 else None)
        gerencia = norm_gerencia(row[3] if len(row) > 3 else "")
        if not objeto or gerencia not in ("INFRA", "EDIFICACOES"):
            continue
        cnpj = ente if len(so_digitos(ente)) >= 11 else None
        watch.append({
            "id": i, "ente_cnpj": cnpj,
            "ente_nome": None if cnpj else ente,
            "objeto": objeto, "valor": valor, "gerencia": gerencia,
        })
    wb.close()
    return watch


def cmd_list(args):
    watch = carrega_watchlist()
    if not watch:
        return
    print(f"\n{'LIN':<5}{'GER':<12}{'ENTE':<26}{'VALOR':>16}   OBJETO")
    print("-" * 92)
    for w in watch:
        ente = w["ente_cnpj"] or (w["ente_nome"] or "")[:24]
        v = f"R$ {w['valor']:,.2f}" if w["valor"] else "-"
        print(f"{w['id']:<5}{w['gerencia']:<12}{ente:<26}{v:>16}   {w['objeto'][:38]}")
    print(f"\nTotal: {len(watch)} edital(is) vigiado(s).\n")


# ======================================================================
# BANCO DE ESTADO
# ======================================================================
def init_db():
    con = sqlite3.connect(DB)
    con.execute("""CREATE TABLE IF NOT EXISTS editais (
        id_pncp TEXT PRIMARY KEY, cnpj_orgao TEXT, orgao TEXT, municipio TEXT,
        objeto TEXT, valor REAL, situacao TEXT, situacao_ant TEXT,
        data_public TEXT, data_atualiz TEXT, visto_em TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS alertas (
        id INTEGER PRIMARY KEY AUTOINCREMENT, tipo TEXT, id_pncp TEXT,
        id_pncp_ref TEXT, gerencia TEXT, score REAL, detalhe TEXT, criado_em TEXT)""")
    con.execute("""CREATE UNIQUE INDEX IF NOT EXISTS ix_alerta_unico
        ON alertas(tipo, id_pncp, id_pncp_ref)""")
    con.commit()
    return con


def ja_alertado(con, tipo, id_pncp, ref):
    return con.execute(
        "SELECT 1 FROM alertas WHERE tipo=? AND id_pncp=? AND id_pncp_ref=?",
        (tipo, id_pncp, ref or "")).fetchone() is not None


# ======================================================================
# COLETA NO PNCP (com retry/backoff)
# ======================================================================
def _iso(d):
    return d.strftime("%Y%m%d")   # formato exigido pela API: yyyyMMdd


def coleta(endpoint, data_ini, data_fim, modalidade):
    """Itera as paginas do endpoint com repeticao automatica em falhas
    transitorias do PNCP (timeout / HTTP 5xx / 429)."""
    registros, pagina = [], 1
    while True:
        params = {"dataInicial": _iso(data_ini), "dataFinal": _iso(data_fim),
                  "codigoModalidadeContratacao": modalidade, "uf": UF_ALVO,
                  "pagina": pagina, "tamanhoPagina": TAM_PAGINA}
        r, sucesso = None, False
        for tent in range(1, TENTATIVAS + 1):
            try:
                r = requests.get(f"{BASE}{endpoint}", params=params, timeout=TIMEOUT)
            except requests.RequestException as e:
                print(f"  ! rede (tent {tent}/{TENTATIVAS}) {endpoint} pag {pagina}: {e}")
                time.sleep(BACKOFF * tent)
                continue
            if r.status_code == 204:      # sem conteudo -> fim do endpoint
                return registros
            if r.status_code in (429, 500, 502, 503, 504):
                print(f"  ! HTTP {r.status_code} (tent {tent}/{TENTATIVAS}) "
                      f"{endpoint} pag {pagina} - aguardando e tentando de novo")
                time.sleep(BACKOFF * tent)
                continue
            if r.status_code != 200:      # erro nao-transitorio -> encerra endpoint
                print(f"  ! HTTP {r.status_code} em {endpoint} pag {pagina}")
                return registros
            sucesso = True
            break

        if not sucesso:                   # esgotou as tentativas nesta pagina
            print(f"  ! desistindo da pag {pagina} de {endpoint} "
                  f"apos {TENTATIVAS} tentativas (PNCP instavel).")
            break

        payload = r.json()
        data = payload.get("data", payload.get("content", []))
        registros.extend(data)
        total_pag = payload.get("totalPaginas", payload.get("totalPages", 1))
        if pagina >= total_pag or not data:
            break
        pagina += 1
        time.sleep(PAUSA_PAGINA)
    return registros


def extrai(item):
    uni = item.get("unidadeOrgao", {}) or {}
    org = item.get("orgaoEntidade", {}) or {}
    return {
        "id_pncp": item.get("numeroControlePNCP") or item.get("id"),
        "cnpj_orgao": org.get("cnpj"), "orgao": org.get("razaoSocial"),
        "municipio": uni.get("municipioNome"), "uf": uni.get("ufSigla"),
        "objeto": item.get("objetoCompra") or item.get("objeto") or "",
        "valor": item.get("valorTotalEstimado") or 0.0,
        "situacao": (item.get("situacaoCompraNome")
                     or str(item.get("situacaoCompraId") or "")),
        "data_public": item.get("dataPublicacaoPncp"),
        "data_atualiz": item.get("dataAtualizacao")
                        or item.get("dataAtualizacaoGlobal"),
    }


# ======================================================================
# MOTOR DE DETECCAO
# ======================================================================
SIT_SUSPENSA  = ("suspens",)
SIT_ENCERRADA = ("revog", "anulad")
SIT_ATIVA     = ("divulg", "receb", "aberta", "publicad")


def _tem(sit, chaves):
    s = (sit or "").lower()
    return any(k in s for k in chaves)


def melhor_watch(watchlist, c):
    melhor, best, bobj = None, 0.0, 0.0
    for w in watchlist:
        sc, s_obj, _ = score_match(w, c["objeto"], c["valor"],
                                   c["cnpj_orgao"], c["orgao"])
        if sc > best:
            melhor, best, bobj = w, sc, s_obj
    if melhor and best >= SCORE_MINIMO and bobj >= LIMIAR_OBJETO:
        return melhor, best, bobj
    return None, 0.0, 0.0


def processa(con, registros, watchlist):
    cur = con.cursor()
    agora = dt.datetime.now().isoformat(timespec="seconds")
    alertas = []
    for item in registros:
        c = extrai(item)
        if not c["id_pncp"] or (c.get("uf") not in (UF_ALVO, None)):
            continue
        w, score, s_obj = melhor_watch(watchlist, c)
        prev = cur.execute("SELECT situacao FROM editais WHERE id_pncp=?",
                           (c["id_pncp"],)).fetchone()
        if not w and not prev:
            continue

        if w and prev and _tem(prev[0], SIT_SUSPENSA) and _tem(c["situacao"], SIT_ATIVA):
            if not ja_alertado(con, "RETOMADA", c["id_pncp"], ""):
                det = (f"{c['orgao']} ({c['municipio']}): "
                       f"'{prev[0]}' -> '{c['situacao']}' | R$ {c['valor']:,.2f}")
                alertas.append(dict(tipo="RETOMADA", w=w, c=c, ref="",
                                    score=score, det=det))

        if w and not prev:
            cands = cur.execute(
                "SELECT id_pncp,objeto,valor,cnpj_orgao,orgao FROM editais "
                "WHERE situacao LIKE '%revog%' OR situacao LIKE '%anulad%'").fetchall()
            ref = ""
            for (idp, obj, val, cnpj, org) in cands:
                sc, so, _ = score_match(
                    {"objeto": obj, "valor": val, "ente_cnpj": cnpj,
                     "ente_nome": org}, c["objeto"], c["valor"],
                    c["cnpj_orgao"], c["orgao"])
                if sc >= SCORE_MINIMO and so >= LIMIAR_OBJETO:
                    ref = idp; break
            if ref and not ja_alertado(con, "REPUBLICACAO", c["id_pncp"], ref):
                det = (f"NOVO {c['id_pncp']} ~ REVOGADO {ref} | "
                       f"{c['orgao']} ({c['municipio']}) | R$ {c['valor']:,.2f}")
                alertas.append(dict(tipo="REPUBLICACAO", w=w, c=c, ref=ref,
                                    score=score, det=det))

        cur.execute("""INSERT INTO editais
            (id_pncp,cnpj_orgao,orgao,municipio,objeto,valor,
             situacao,situacao_ant,data_public,data_atualiz,visto_em)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(id_pncp) DO UPDATE SET
             situacao_ant=editais.situacao, situacao=excluded.situacao,
             valor=excluded.valor, data_atualiz=excluded.data_atualiz,
             visto_em=excluded.visto_em""",
            (c["id_pncp"], c["cnpj_orgao"], c["orgao"], c["municipio"],
             c["objeto"], c["valor"], c["situacao"],
             prev[0] if prev else None,
             c["data_public"], c["data_atualiz"], agora))

    for a in alertas:
        cur.execute("""INSERT OR IGNORE INTO alertas
            (tipo,id_pncp,id_pncp_ref,gerencia,score,detalhe,criado_em)
            VALUES (?,?,?,?,?,?,?)""",
            (a["tipo"], a["c"]["id_pncp"], a["ref"], a["w"]["gerencia"],
             a["score"], a["det"], agora))
    con.commit()
    return alertas


# ======================================================================
# NOTIFICACAO NO TEAMS
# ======================================================================
def envia_teams(gerencia, tipo, c, det, score, ref=""):
    url = WEBHOOKS.get(gerencia, "")
    if not url or url.startswith("COLE_"):
        print(f"[AVISO] Webhook de {gerencia} nao configurado; alerta nao enviado.")
        return False
    cor = "Attention" if tipo == "REPUBLICACAO" else "Warning"
    titulo = ("Edital REPUBLICADO (apos revogacao)" if tipo == "REPUBLICACAO"
              else "Edital RETOMADO (apos suspensao)")
    link = f"https://pncp.gov.br/app/editais/{c.get('id_pncp','')}"
    facts = [
        {"title": "Evento", "value": tipo},
        {"title": "Ente", "value": f"{c.get('orgao','-')} ({c.get('municipio','-')})"},
        {"title": "Valor", "value": f"R$ {c.get('valor',0):,.2f}"},
        {"title": "Situacao", "value": c.get("situacao", "-")},
        {"title": "Score", "value": f"{score:.0%}"},
        {"title": "Gerencia", "value": f"{gerencia} - {GERENTES.get(gerencia,'')}"},
    ]
    if ref:
        facts.append({"title": "Edital anterior", "value": ref})
    card = {"type": "message", "attachments": [{
        "contentType": "application/vnd.microsoft.card.adaptive",
        "content": {"$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard", "version": "1.4",
            "body": [
                {"type": "TextBlock", "size": "Large", "weight": "Bolder",
                 "color": cor, "text": titulo, "wrap": True},
                {"type": "TextBlock", "text": c.get("objeto", "")[:280],
                 "wrap": True, "spacing": "Small"},
                {"type": "FactSet", "facts": facts}],
            "actions": [{"type": "Action.OpenUrl", "title": "Abrir no PNCP",
                         "url": link}]}}]}
    try:
        r = requests.post(url, json=card, timeout=30)
        ok = r.status_code in (200, 202)
        print(f"  -> Teams[{gerencia}] {'enviado' if ok else 'falhou'} (HTTP {r.status_code})")
        return ok
    except requests.RequestException as e:
        print(f"  ! erro Teams: {e}"); return False


def notifica(alertas):
    if not alertas:
        print(">> Nenhum evento novo (retomada/republicacao) para os editais vigiados.")
        return
    print(f"\n===== {len(alertas)} ALERTA(S) =====")
    for a in alertas:
        print(f"[{a['tipo']}] {a['det']}")
        envia_teams(a["w"]["gerencia"], a["tipo"], a["c"],
                    a["det"], a["score"], a["ref"])


# ======================================================================
# EXECUCAO PRINCIPAL
# ======================================================================
def executar_monitoramento():
    hoje = dt.date.today()
    marca = "" if eh_dia_util(hoje) else "  (atencao: hoje nao e dia util BR/PR)"
    print("=" * 62)
    print(" MONITOR DE EDITAIS - PNCP (Parana)")
    print(f" Execucao: {hoje:%d/%m/%Y}{marca}")
    print(f" Janela consultada: {(hoje - dt.timedelta(days=JANELA_DIAS)):%d/%m/%Y}"
          f" ate {hoje:%d/%m/%Y}")
    print("=" * 62)

    watchlist = carrega_watchlist()
    if not watchlist:
        print("Watchlist vazia ou planilha ausente. Nada a fazer.")
        return
    print(f"Editais vigiados na planilha: {len(watchlist)}")

    con = init_db()
    ini = hoje - dt.timedelta(days=JANELA_DIAS)
    todos = []
    for mod in MODALIDADES_OBRAS:
        print(f"  . coletando modalidade {mod} ...")
        todos += coleta("/v1/contratacoes/atualizacao", ini, hoje, mod)
        todos += coleta("/v1/contratacoes/publicacao",  ini, hoje, mod)
    print(f"Registros coletados do PNCP: {len(todos)}")

    alertas = processa(con, todos, watchlist)
    notifica(alertas)
    con.close()
    print("\nConcluido.")


def cmd_run(args):
    executar_monitoramento()


def cmd_testalert(args):
    g = norm_gerencia(args.gerencia or "INFRA")
    demo = {"id_pncp": "TESTE-0000", "orgao": "MUNICIPIO DE EXEMPLO",
            "municipio": "Curitiba", "valor": 1234567.89,
            "situacao": "Divulgada no PNCP",
            "objeto": "Teste de alerta do monitor de editais (PNCP)."}
    envia_teams(g, "RETOMADA", demo, "teste de conectividade", 0.99)


def _pausa_se_dois_cliques():
    if os.name == "nt":
        try:
            input("\nPressione ENTER para fechar...")
        except EOFError:
            pass


# ======================================================================
def main():
    if len(sys.argv) == 1:
        try:
            executar_monitoramento()
        except Exception as e:
            print(f"\n[ERRO] {e}")
        _pausa_se_dois_cliques()
        return

    p = argparse.ArgumentParser(description="Monitor de editais PNCP (PR).")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("list", help="listar a watchlist").set_defaults(func=cmd_list)
    rn = sub.add_parser("run", help="rodar varredura (ultimos 7 dias)")
    rn.set_defaults(func=cmd_run)
    t = sub.add_parser("testalert", help="testar webhook do Teams")
    t.add_argument("--gerencia", choices=["INFRA", "EDIFICACOES"])
    t.set_defaults(func=cmd_testalert)
    args = p.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
