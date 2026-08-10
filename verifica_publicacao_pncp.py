#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verifica_publicacao_pncp.py  (v2 - TODOS os 399 municipios do PR)
======================================================================
Verifica, para TODOS os municipios do Parana, se cada um JA PUBLICA
no PNCP (qualquer contratacao, qualquer modalidade) no periodo
01/01/2025 ate hoje.

Contexto legal: art. 176 da Lei 14.133/2021 concede aos municipios com
ate 20.000 habitantes prazo estendido para adesao ao PNCP; por isso
alguns ainda podem nao publicar.

COMO FUNCIONA
  1) Baixa da API do IBGE a lista OFICIAL de TODOS os municipios do PR
     (UF 41) -> nao e preciso digitar nada; sao os 399 municipios.
  2) Para cada municipio, consulta a API do PNCP
     (/v1/contratacoes/publicacao) no periodo 01/01/2025 ate hoje.
     -> AO ACHAR O PRIMEIRO registro (em qualquer modalidade), marca
        SIM e ja passa para o proximo municipio (agiliza a checagem).
     -> Se nenhuma modalidade retornar nada, marca NAO.
  3) Gera a planilha 'publicacao_pncp_2026.xlsx' (os que NAO publicam
     vem destacados em vermelho) e imprime a lista dos que NAO publicam.

Rode NA SUA MAQUINA (as APIs bloqueiam IPs de datacenter/nuvem) OU gere
o .exe pelo GitHub Actions (workflow incluido).
Dependencias:  requests, openpyxl
======================================================================
"""

import datetime as dt
import os
import re
import sys
import time
import unicodedata

try:
    import requests
except ImportError:
    sys.exit("Instale 'requests':  pip install requests")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    sys.exit("Instale 'openpyxl':  pip install openpyxl")

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------
BASE_PNCP = "https://pncp.gov.br/api/consulta"
IBGE_URL = "https://servicodados.ibge.gov.br/api/v1/localidades/estados/41/municipios"
DATA_INI = dt.date(2025, 1, 1)          # inicio do periodo de checagem
DATA_FIM = dt.date.today()              # ate hoje
ARQ_SAIDA = "publicacao_pncp_2026.xlsx"
TIMEOUT = 60
TENTATIVAS = 4
BACKOFF = 5
HEADERS = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}

# Modalidades ordenadas para achar QUALQUER publicacao rapido (as mais
# usadas primeiro). Codigos oficiais da tabela de dominio do PNCP.
#  6 Pregao Elet. | 8 Dispensa | 9 Inexigib. | 7 Pregao Pres.
#  4 Concorr.Elet.| 5 Concorr.Pres.|12 Credenciamento|13 Leilao Pres.
#  1 Leilao Elet. | 2 Dialogo | 3 Concurso |10 Manif.Interesse |11 Pre-qualif.
MODALIDADES = [6, 8, 9, 7, 4, 5, 12, 13, 1, 2, 3, 10, 11]


# ----------------------------------------------------------------------
def norm(s):
    if not s:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c)).lower()
    s = s.replace("'", " ").replace("`", " ")
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return " ".join(s.split())


def get_municipios_pr():
    """Lista oficial de TODOS os municipios do PR: [(nome, ibge), ...]."""
    r = requests.get(IBGE_URL, timeout=TIMEOUT, headers=HEADERS)
    r.raise_for_status()
    muns = [(m["nome"], str(m["id"])) for m in r.json()]
    muns.sort(key=lambda x: norm(x[0]))
    return muns


def _get(endpoint, params):
    for tent in range(1, TENTATIVAS + 1):
        try:
            r = requests.get(f"{BASE_PNCP}{endpoint}", params=params,
                             timeout=TIMEOUT, headers=HEADERS)
        except requests.RequestException:
            if tent > 1:
                print(f"     >> PNCP instavel, aguarde... ({tent}/{TENTATIVAS})")
            time.sleep(BACKOFF * tent)
            continue
        if r.status_code == 204:
            return {"__vazio__": True}
        if r.status_code in (429, 500, 502, 503, 504):
            if tent > 1:
                print(f"     >> PNCP instavel, aguarde... ({tent}/{TENTATIVAS})")
            time.sleep(BACKOFF * tent)
            continue
        if r.status_code != 200:
            return None
        try:
            return r.json()
        except ValueError:
            return None
    return None


def publica_no_periodo(ibge):
    """Retorna (True, modalidade) ASSIM QUE achar o 1o registro; senao
    (False, None). Para na primeira modalidade com resultado (agiliza)."""
    di, dfim = DATA_INI.strftime("%Y%m%d"), DATA_FIM.strftime("%Y%m%d")
    for mod in MODALIDADES:
        params = {"dataInicial": di, "dataFinal": dfim,
                  "codigoModalidadeContratacao": mod, "uf": "PR",
                  "codigoMunicipioIbge": ibge, "pagina": 1, "tamanhoPagina": 10}
        payload = _get("/v1/contratacoes/publicacao", params)
        if not payload or payload.get("__vazio__"):
            continue
        data = payload.get("data", payload.get("content", []))
        if data:                         # achou -> ja retorna (short-circuit)
            return True, mod
    return False, None


# ----------------------------------------------------------------------
def salva_excel(resultados):
    wb = Workbook()
    ws = wb.active
    ws.title = "Publicacao PNCP"
    headers = ["Municipio", "Codigo IBGE", "Publica no PNCP?",
               "Modalidade do 1o achado", "Periodo verificado"]
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
    verm = PatternFill("solid", fgColor="F8CBAD")   # destaca NAO publica
    periodo = f"{DATA_INI:%d/%m/%Y} a {DATA_FIM:%d/%m/%Y}"
    for (nome, ibge, publica, mod) in resultados:
        ws.append([nome, ibge, "SIM" if publica else "NAO",
                   mod if publica else "-", periodo])
        if not publica:
            for col in range(1, len(headers) + 1):
                ws.cell(row=ws.max_row, column=col).fill = verm
    for i, wid in enumerate([28, 13, 18, 24, 22], start=1):
        ws.column_dimensions[chr(64 + i)].width = wid
    ws.freeze_panes = "A2"
    caminho = os.path.join(BASE_DIR, ARQ_SAIDA)
    wb.save(caminho)
    return caminho


def main():
    print("=" * 66)
    print(" VERIFICADOR DE PUBLICACAO NO PNCP (todos os municipios do PR)")
    print(f" Periodo: {DATA_INI:%d/%m/%Y} ate {DATA_FIM:%d/%m/%Y}")
    print("=" * 66)

    print("Baixando a lista oficial de municipios do PR (API do IBGE)...")
    try:
        municipios = get_municipios_pr()
    except Exception as e:
        sys.exit(f"[ERRO] Nao consegui a lista do IBGE: {e}\n"
                 "Rode este script na sua maquina (rede sem bloqueio).")
    print(f"  OK: {len(municipios)} municipios do PR.\n")

    resultados, nao_publicam = [], []
    total = len(municipios)
    for i, (nome, ibge) in enumerate(municipios, start=1):
        publica, mod = publica_no_periodo(ibge)
        status = f"SIM (mod {mod})" if publica else "NAO PUBLICA"
        print(f"[{i:>3}/{total}] {nome:<30} {ibge}  -> {status}")
        resultados.append((nome, ibge, publica, mod))
        if not publica:
            nao_publicam.append((nome, ibge))
        time.sleep(0.15)                 # educado com o servidor

    caminho = salva_excel(resultados)

    print("\n" + "=" * 66)
    print(" RESULTADO")
    print("=" * 66)
    print(f"Total verificado: {total} municipios")
    print(f"Publicam: {total - len(nao_publicam)}   |   "
          f"NAO publicam: {len(nao_publicam)}\n")
    print(f"Municipios que AINDA NAO PUBLICAM no PNCP "
          f"({DATA_INI:%d/%m/%Y} a {DATA_FIM:%d/%m/%Y}):")
    if nao_publicam:
        for nome, ibge in nao_publicam:
            print(f"   - {nome} ({ibge})")
    else:
        print("   (nenhum - todos publicaram ao menos uma vez no periodo)")
    print(f"\nPlanilha salva em:\n   {caminho}")

    if os.name == "nt":
        try:
            input("\nPressione ENTER para fechar...")
        except EOFError:
            pass


if __name__ == "__main__":
    main()
