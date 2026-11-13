# -*- coding: utf-8 -*-
"""
Extração ZTPM028 / ZTMM204 + empilhamento na ZTMM028_25_26

Fluxo:
 1. Janela para informar Data Início / Data Fim (dd.mm.aaaa)
 2. ZTPM028 -> exporta ztpm028.xlsx  (força fechar o Excel aberto pelo SAP)
 3. 'Nº documento 2' do ztpm028 -> ZTMM204 -> exporta ztmm204.xlsx  (força fechar o Excel)
 4. ZTMM028_25_26.xlsx (aba Data), aberta uma única vez:
    - empilha ztpm028 A:T + data de hoje na coluna W
    - PROCV: base K (Nº documento 2) x ztmm204 B (Nº documento)
             -> H (Nº NF) em U (NF TRIANGULAR) e I (Séries) em V (SERE NF TRIANGULAR)

Requisitos: pywin32, openpyxl  |  SAP GUI com scripting habilitado e logado.
"""

import re
import sys
import time
import subprocess
import unicodedata
import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
from pathlib import Path

import win32com.client
import win32clipboard
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

# ============================== CONFIGURAÇÃO ==============================
PASTA_SAIDA = Path.home() / "OneDrive - Claro SA" / "Planilha de Atendimento" / "Repositorio" / "NF Triangulação"

ARQ_ZTPM028 = "ztpm028.xlsx"
ARQ_ZTMM204 = "ztmm204.xlsx"

EMPRESAS_ZTMM204 = ["001", "705", "881"]
COLUNA_DOC2 = "Nº documento 2"

GRID_ZTPM028 = "wnd[0]/usr/cntlCC_ALV/shellcont/shell"
GRID_ZTMM204 = "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell"

TIMEOUT_ARQUIVO = 120     # segundos esperando o arquivo exportado aparecer
ESPERA_EXCEL_SAP = 15     # segundos aguardando o SAP abrir o Excel após exportar
TIMEOUT_DESBLOQUEIO = 30  # segundos aguardando o arquivo ser liberado após matar o Excel

# ---- Empilhamento + PROCV (ZTMM028_25_26.xlsx) ----
ARQ_BASE = "ZTMM028_25_26.xlsx"
ABA_BASE = "Data"
N_COLS_EMPILHAR = 20              # ztpm028 A:T -> base A:T (por posição; cabeçalhos da base estão truncados)
COL_NF_DESTINO = "U"              # NF TRIANGULAR
COL_SERIE_DESTINO = "V"           # SERE NF TRIANGULAR
COL_DATA = "W"                    # data do empilhamento
TITULO_COL_DATA = "DATA EMPILHAMENTO"
COL_CHAVE_BASE = "K"              # Nº documento 2 (base)

# PROCV na ztmm204: B = Nº documento | H = Nº NF | I = Séries
COL_CHAVE_204 = "B"
COL_NF_204 = "H"
COL_SERIE_204 = "I"
# ==========================================================================


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


# ------------------------------- INPUT DATAS -------------------------------
def pedir_datas():
    resultado = {}
    hoje = datetime.now()

    root = tk.Tk()
    root.title("Extração ZTPM028 / ZTMM204")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    frm = tk.Frame(root, padx=16, pady=14)
    frm.pack()

    tk.Label(frm, text="Data início (dd.mm.aaaa):").grid(row=0, column=0, sticky="w", pady=4)
    e_ini = tk.Entry(frm, width=14, justify="center")
    e_ini.insert(0, (hoje - timedelta(days=1)).strftime("%d.%m.%Y"))
    e_ini.grid(row=0, column=1, padx=6)

    tk.Label(frm, text="Data fim (dd.mm.aaaa):").grid(row=1, column=0, sticky="w", pady=4)
    e_fim = tk.Entry(frm, width=14, justify="center")
    e_fim.insert(0, hoje.strftime("%d.%m.%Y"))
    e_fim.grid(row=1, column=1, padx=6)

    def ok(_=None):
        try:
            d_ini = datetime.strptime(e_ini.get().strip(), "%d.%m.%Y")
            d_fim = datetime.strptime(e_fim.get().strip(), "%d.%m.%Y")
        except ValueError:
            messagebox.showerror("Data inválida", "Use o formato dd.mm.aaaa", parent=root)
            return
        if d_fim < d_ini:
            messagebox.showerror("Data inválida", "Data fim menor que data início.", parent=root)
            return
        resultado["ini"] = d_ini.strftime("%d.%m.%Y")
        resultado["fim"] = d_fim.strftime("%d.%m.%Y")
        root.destroy()

    btns = tk.Frame(frm, pady=8)
    btns.grid(row=2, column=0, columnspan=2)
    tk.Button(btns, text="Executar", width=12, command=ok).pack(side="left", padx=4)
    tk.Button(btns, text="Cancelar", width=12, command=root.destroy).pack(side="left", padx=4)

    root.bind("<Return>", ok)
    e_ini.focus_set()
    e_ini.select_range(0, "end")
    root.mainloop()

    if not resultado:
        return None, None
    return resultado["ini"], resultado["fim"]


# --------------------------------- SAP ------------------------------------
def conectar_sap():
    try:
        sapgui = win32com.client.GetObject("SAPGUI")
        app = sapgui.GetScriptingEngine
        conn = app.Children(0)
        return conn.Children(0)
    except Exception as e:
        raise RuntimeError("SAP GUI não encontrado. Abra e faça login no SAP antes de rodar.") from e


def existe(session, id_):
    try:
        session.findById(id_)
        return True
    except Exception:
        return False


def press_primeiro(session, ids):
    for id_ in ids:
        if existe(session, id_):
            session.findById(id_).press()
            return id_
    raise RuntimeError(f"Nenhum dos botões encontrados: {ids}")


def abrir_transacao(session, tcode):
    session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{tcode}"
    session.findById("wnd[0]").sendVKey(0)


def msg_status(session):
    try:
        return session.findById("wnd[0]/sbar").text
    except Exception:
        return ""


# ------------------------------ EXCEL / ARQUIVOS ---------------------------
def excel_rodando():
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
            capture_output=True, text=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        ).stdout
        return "excel.exe" in out.lower()
    except Exception:
        return False


def matar_excel():
    """Força o fechamento de TODOS os processos do Excel (sem salvar)."""
    if not excel_rodando():
        return False
    subprocess.run(
        ["taskkill", "/F", "/IM", "EXCEL.EXE", "/T"],
        capture_output=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    log("Excel fechado à força.")
    time.sleep(1)
    return True


def arquivo_livre(caminho: Path):
    """True se o arquivo não está bloqueado por outro processo."""
    if not caminho.exists():
        return True
    try:
        with open(caminho, "r+b"):
            return True
    except (PermissionError, OSError):
        return False


def esperar_liberar(caminho: Path, timeout=TIMEOUT_DESBLOQUEIO):
    fim = time.time() + timeout
    while time.time() < fim:
        if arquivo_livre(caminho):
            return
        matar_excel()
        time.sleep(1)
    raise RuntimeError(f"Arquivo continua bloqueado: {caminho}\nFeche o Excel/OneDrive e rode de novo.")


def apagar_anterior(caminho: Path):
    if caminho.exists():
        esperar_liberar(caminho)
        caminho.unlink()


def esperar_arquivo(caminho: Path, timeout=TIMEOUT_ARQUIVO):
    fim = time.time() + timeout
    ultimo = -1
    while time.time() < fim:
        if caminho.exists():
            tam = caminho.stat().st_size
            if tam > 0 and tam == ultimo:
                return
            ultimo = tam
        time.sleep(1)
    raise RuntimeError(f"Arquivo não foi gerado no tempo limite: {caminho}")


def fechar_excel_do_sap(caminho: Path):
    """O SAP abre o Excel alguns segundos depois de salvar: espera ele aparecer e mata."""
    fim = time.time() + ESPERA_EXCEL_SAP
    while time.time() < fim:
        if excel_rodando():
            time.sleep(1)  # deixa terminar de abrir
            matar_excel()
            break
        time.sleep(0.5)
    esperar_liberar(caminho)


def exportar_alv(session, grid_id, pasta: Path, nome: str):
    destino = pasta / nome
    matar_excel()
    apagar_anterior(destino)

    grid = session.findById(grid_id)
    grid.contextMenu()
    grid.selectContextMenuItem("&XXL")

    if existe(session, "wnd[1]") and not existe(session, "wnd[1]/usr/ctxtDY_PATH"):
        press_primeiro(session, ["wnd[1]/tbar[0]/btn[20]", "wnd[1]/tbar[0]/btn[0]"])

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(pasta)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome
    press_primeiro(session, ["wnd[1]/tbar[0]/btn[11]", "wnd[1]/tbar[0]/btn[0]"])

    esperar_arquivo(destino)
    fechar_excel_do_sap(destino)
    log(f"Exportado: {destino}")
    return destino


def _norm(txt):
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", txt.lower())


# ------------------------------ ZTPM028 -----------------------------------
def rodar_ztpm028(session, dt_ini, dt_fim):
    log("ZTPM028: abrindo transação...")
    abrir_transacao(session, "ztpm028")

    session.findById("wnd[0]/usr/ctxtS_DOCN_1-LOW").text = "1"
    session.findById("wnd[0]/usr/ctxtS_DOCN_1-HIGH").text = "9999999999"
    session.findById("wnd[0]/usr/ctxtS_PSTDAT-LOW").text = dt_ini
    session.findById("wnd[0]/usr/ctxtS_PSTDAT-HIGH").text = dt_fim
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    if not existe(session, GRID_ZTPM028):
        raise RuntimeError(f"ZTPM028 sem resultado. SAP: {msg_status(session) or '(sem mensagem)'}")

    return exportar_alv(session, GRID_ZTPM028, PASTA_SAIDA, ARQ_ZTPM028)


def ler_documentos2(caminho: Path):
    wb = load_workbook(caminho, read_only=True, data_only=True)
    linhas = list(wb.worksheets[0].iter_rows(values_only=True))
    wb.close()
    cab = [str(h or "").strip() for h in linhas[0]] if linhas else []
    alvo = _norm(COLUNA_DOC2)
    i = next((j for j, h in enumerate(cab) if _norm(h) == alvo), None)
    if i is None:
        raise RuntimeError(f"Coluna '{COLUNA_DOC2}' não encontrada. Colunas: {cab}")
    col = cab[i]

    vals = []
    for r in linhas[1:]:
        v = r[i] if i < len(r) else None
        if v is None:
            continue
        v = str(v).strip()
        if v.endswith(".0"):
            v = v[:-2]
        if v and v.strip("0") and v not in vals:
            vals.append(v)

    log(f"'{col}': {len(vals)} documento(s) único(s).")
    return vals


def copiar_clipboard(valores):
    texto = "\r\n".join(valores)
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(texto, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


# ------------------------------ ZTMM204 -----------------------------------
def rodar_ztmm204(session, dt_ini, dt_fim, documentos):
    log("ZTMM204: abrindo transação...")
    abrir_transacao(session, "ztmm204")

    session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = EMPRESAS_ZTMM204[0]
    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = dt_ini
    session.findById("wnd[0]/usr/ctxtS_DATA-HIGH").text = dt_fim

    session.findById("wnd[0]/usr/btn%_S_BUKRS_%_APP_%-VALU_PUSH").press()
    base = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{}]"
    for i, emp in enumerate(EMPRESAS_ZTMM204[1:], start=1):
        session.findById(base.format(i)).text = emp
    session.findById("wnd[1]").sendVKey(8)

    copiar_clipboard(documentos)
    session.findById("wnd[0]/usr/btn%_S_DOCNUM_%_APP_%-VALU_PUSH").press()
    if existe(session, "wnd[1]/tbar[0]/btn[16]"):
        session.findById("wnd[1]/tbar[0]/btn[16]").press()   # apagar seleção anterior
    session.findById("wnd[1]/tbar[0]/btn[24]").press()       # colar da área de transferência
    session.findById("wnd[1]/tbar[0]/btn[8]").press()        # copiar

    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    if not existe(session, GRID_ZTMM204):
        raise RuntimeError(f"ZTMM204 sem resultado. SAP: {msg_status(session) or '(sem mensagem)'}")

    return exportar_alv(session, GRID_ZTMM204, PASTA_SAIDA, ARQ_ZTMM204)


# ------------------------- EMPILHAR + PROCV --------------------------------
def _chave(v):
    """Normaliza Nº documento: tira .0, espaços e zeros à esquerda."""
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.lstrip("0") if s.isdigit() else s


def _ler_linhas(caminho: Path):
    esperar_liberar(caminho)
    wb = load_workbook(caminho, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    linhas = [r for r in ws.iter_rows(min_row=2, values_only=True)
              if any(v not in (None, "") for v in r)]
    wb.close()
    return linhas


def _ultima_linha(ws, col=1):
    r = ws.max_row
    while r > 1 and ws.cell(r, col).value in (None, ""):
        r -= 1
    return r


def montar_lookup_204(arq_204: Path):
    """{Nº documento: (Nº NF, Séries)} da ztmm204 - primeira ocorrência, igual PROCV."""
    i_k = column_index_from_string(COL_CHAVE_204) - 1
    i_nf = column_index_from_string(COL_NF_204) - 1
    i_se = column_index_from_string(COL_SERIE_204) - 1
    lookup = {}
    for r in _ler_linhas(arq_204):
        k = _chave(r[i_k])
        if k and k not in lookup:
            lookup[k] = (r[i_nf], r[i_se])
    log(f"ztmm204: {len(lookup)} documento(s) para o PROCV.")
    return lookup


def empilhar_e_procv(arq_028: Path, lookup_204: dict):
    """Abre a base UMA vez: empilha o ztpm028 (A:T) com a data de hoje e faz o PROCV da ztmm204."""
    arq_base = PASTA_SAIDA / ARQ_BASE
    if not arq_base.exists():
        raise RuntimeError(f"Planilha base não encontrada: {arq_base}")
    esperar_liberar(arq_base)

    dados = _ler_linhas(arq_028)

    log(f"Abrindo {ARQ_BASE} (pode levar alguns segundos)...")
    wb = load_workbook(arq_base)
    ws = wb[ABA_BASE]

    c_data = column_index_from_string(COL_DATA)
    if ws.cell(1, c_data).value in (None, ""):
        ws.cell(1, c_data).value = TITULO_COL_DATA

    # ---------- 1) Empilhar ----------
    hoje = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    inicio = _ultima_linha(ws) + 1
    linha = inicio
    for reg in dados:
        for c in range(N_COLS_EMPILHAR):
            ws.cell(linha, c + 1).value = reg[c] if c < len(reg) else None
        cel = ws.cell(linha, c_data)
        cel.value = hoje
        cel.number_format = "dd/mm/yyyy"
        linha += 1
    ultima = linha - 1
    log(f"Empilhadas {len(dados)} linha(s) ({inicio} a {ultima}) com data {hoje:%d/%m/%Y}.")

    # ---------- 2) PROCV: base K -> ztmm204 B -> H/I em U/V ----------
    c_k = column_index_from_string(COL_CHAVE_BASE)
    c_nf = column_index_from_string(COL_NF_DESTINO)
    c_se = column_index_from_string(COL_SERIE_DESTINO)
    achados = novos_sem = 0
    for r in range(2, ultima + 1):
        k = _chave(ws.cell(r, c_k).value)
        if k in lookup_204:
            nf, serie = lookup_204[k]
            ws.cell(r, c_nf).value = nf
            ws.cell(r, c_se).value = serie
            achados += 1
        elif r >= inicio:
            novos_sem += 1
    log(f"PROCV: {achados} linha(s) com NF/Série preenchidas.")
    if novos_sem:
        log(f"⚠ {novos_sem} linha(s) empilhada(s) hoje sem NF na ztmm204.")

    ws.auto_filter.ref = f"A1:{COL_DATA}{ultima}"

    esperar_liberar(arq_base)
    try:
        wb.save(arq_base)
    except PermissionError:
        raise RuntimeError(f"Não consegui salvar {ARQ_BASE} (está aberta?). Feche e rode de novo.")
    log(f"Salvo: {arq_base}")


# -------------------------------- MAIN ------------------------------------
def main():
    dt_ini, dt_fim = pedir_datas()
    if not dt_ini:
        log("Cancelado pelo usuário.")
        return

    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    log(f"Período: {dt_ini} a {dt_fim}")
    log(f"Pasta de saída: {PASTA_SAIDA}")

    session = conectar_sap()

    arq_028 = rodar_ztpm028(session, dt_ini, dt_fim)

    documentos = ler_documentos2(arq_028)
    lookup_204 = {}
    if documentos:
        arq_204 = rodar_ztmm204(session, dt_ini, dt_fim, documentos)
        lookup_204 = montar_lookup_204(arq_204)
    else:
        log("Nenhum 'Nº documento 2' encontrado. ZTMM204 não será executada.")

    matar_excel()  # garantia final antes de mexer na base
    empilhar_e_procv(arq_028, lookup_204)

    abrir_transacao(session, "")
    log("Concluído ✅")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"ERRO: {e}")
        try:
            messagebox.showerror("Erro", str(e))
        except Exception:
            pass
        sys.exit(1)
