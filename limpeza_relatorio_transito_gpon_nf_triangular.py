import pandas as pd
import numpy as np
import os
import time
from datetime import datetime
from pathlib import Path
import re
from pathlib import Path
import win32com.client as win32

inicio = time.time()

PASTA_RELATORIO_TRANSITO = r"C:\Users\F252727\OneDrive - Claro SA\Planilha de Atendimento\Arquivos Relatorio de Transito" #ok
PASTA_CONSOLIDADO_RESID_2026 = r"C:\Users\F252727\OneDrive - Claro SA\USER-Status Reservas Projetos - Documentos" #ok
PASTA_RESERVAS_11SP = r"C:\Users\F252727\OneDrive - Claro SA\USER-Status Reservas Projetos - Documentos" #ok
PASTA_RESERVAS_RESID_2025 = r"C:\Users\F252727\OneDrive - Claro SA\Planilha de Atendimento\Arquivos Relatorio de Transito" #ok
PASTA_PROJETOS_PEP = r"C:\Users\F252727\OneDrive - Claro SA\Planilha de Atendimento\Arquivos Relatorio de Transito" #ok
PASTA_RESERVAS_SIAKI = r"C:\Users\F252727\OneDrive - Claro SA\Planilha de Atendimento\Arquivos Relatorio de Transito" #ok
PASTA_PLAN_ATENDIMENTO_ID = r"C:\Users\F252727\OneDrive - Claro SA\USER-_Logística Claro Brasil - Planejamento - STATUS SHAREPOINT" #ok
PASTA_CPE = r"C:\Users\F252727\OneDrive - Claro SA\USER-Status Reservas Projetos - Documentos"
PASTA_SOBRESSALENTES =r"C:\Users\F252727\OneDrive - Claro SA\USER-Status Reservas Projetos - Documentos"
LOGINS = r"C:\Users\F252727\OneDrive - Claro SA\Planilha de Atendimento\Arquivos Relatorio de Transito"
PASTA_RESERVAS_ELIMINAR = r"C:\Users\F252727\OneDrive - Claro SA\USER-Status Reservas Projetos - Documentos"
PASTA_CLIENTE_RETIRA = r"C:\Users\F252727\OneDrive - Claro SA\USER-Status Reservas Projetos - Documentos"
PASTA_PARCEIRAS = r"C:\Users\F252727\OneDrive - Claro SA\Planilha de Atendimento\Arquivos Relatorio de Transito"
PASTA_CARGA_FRIA = r"C:\Users\F252727\OneDrive - Claro SA\Planilha de Atendimento\Arquivos Relatorio de Transito"

def arquivo_excel_mais_recente(pasta, prefixo):
    arquivos = [
        os.path.join(pasta, f)
        for f in os.listdir(pasta)
        if f.startswith(prefixo) and f.endswith((".xlsx", ".csv", ".xlsb", ".xlsm",".xls"))
    ]

    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo Excel encontrado com esse prefixo")

    return max(arquivos, key=os.path.getmtime)


def buscar_projeto_geral(pep):
    if pd.isna(pep): #verifica se a coluna pep é vazia e traz null
        return None

    match = pp.loc[ #verifica se na coluna "inicio" está na coluna PEP_DESTINO e tras os valoes da coluna pep_geral.  
        pp["inicio"].apply(lambda x: str(pep).startswith(str(x))),
        "projeto_geral"
    ]
    return match.iloc[0] if not match.empty else None  #retorna somente a primeiro valor encontrado do pep_geral


def buscar_atendimento(reserva):
    if pd.isna(reserva):
        return pd.Series([None])

    reserva = str(reserva)

    mask = (
        lm["reserva"].str.contains(reserva, na=False) |
        lm["status.1"].str.contains(reserva, na=False) |
        lm["tran"].str.contains(reserva, na=False)
    )

    resultado = lm.loc[mask]

    if resultado.empty:
        return pd.Series([None])

    return pd.Series([resultado["projeto_lm"].iloc[0]])

def formatacao_excel(workbook, worksheet, df):

    worksheet.hide_gridlines(2) # esconde as linhas de grade
    formato_borda = workbook.add_format({'border': 4}) # padrao da borda em todas as células
    rows, cols = df.shape
    worksheet.conditional_format(0, 0, rows, cols - 1, {'type': 'no_errors', 'format': formato_borda}) # aplica a borda em todas as células
    
    for i, col in enumerate(df.columns): #ajustar a largura das colunas 
        column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, column_len)  

    formato_cabecalho = workbook.add_format({
        'bg_color': 'black',
        'font_color': 'white',
        'bold': True
    })#formato para o cabeçalho

    for col_num, valor in enumerate(df.columns.values):
        worksheet.write(0, col_num, valor, formato_cabecalho)

    worksheet.set_zoom(90) #ajustar o zoom para 90%

# limpa os código SAP, retira os zeros na frente
def limpar_codigo(valor):
    if valor is None:
        return None
    valor = str(valor).strip()
    if valor.isdigit():
        return valor.lstrip("0") or "0"
    return valor

def separar_identificador(valor):
    valor = str(valor).strip()

    # remove .0 do final
    valor = re.sub(r"\.0$", "", valor)

    if valor.upper().startswith("OS"):
        numero = re.sub(r"\D", "", valor)
        return pd.Series(["OS", numero, numero])

    elif valor.upper().startswith("ID"):
        numero = re.sub(r"\D", "", valor)
        return pd.Series(["ID", numero, numero])

    elif valor.isdigit():
        return pd.Series(["PIPEFY", valor, valor])

    return pd.Series([None, None, None])


# Bases de localizacao: procurar junto ao script e na pasta do relatorio.
# Para outra pasta, informe-a nesta variavel.
PASTA_BASES_LOCALIZACAO = r"C:\Users\F252727\OneDrive - Claro SA\Planilha de Atendimento\Arquivos Relatorio de Transito"

_UFS_TRANSITO = set("AC AL AP AM BA CE DF ES GO MA MT MS MG PA PB PR PE PI RJ RN RS RO RR SC SP SE TO".split())

def _texto_localizacao(serie):
    return (serie.astype("string").str.strip().str.upper()
            .str.replace(r"\.0$", "", regex=True)
            .replace({"": pd.NA, "NAN": pd.NA, "NONE": pd.NA, "NULL": pd.NA}))

def _base_localizacao(nome):
    pastas = [Path(__file__).resolve().parent, Path(PASTA_RELATORIO_TRANSITO)]
    if PASTA_BASES_LOCALIZACAO:
        pastas.insert(0, Path(PASTA_BASES_LOCALIZACAO))
    for pasta in pastas:
        caminho = pasta / nome
        if caminho.is_file():
            base = pd.read_excel(caminho, engine="openpyxl", dtype="string")
            base.columns = base.columns.astype(str).str.strip().str.lower()
            print("Base de localizacao:", caminho)
            return base
    raise FileNotFoundError(f"Base {nome!r} nao encontrada. Coloque-a junto ao script "
                            "ou configure PASTA_BASES_LOCALIZACAO.")

def _coluna_localizacao(base, alternativas, nome):
    for coluna in alternativas:
        if coluna in base.columns:
            return coluna
    raise ValueError(f"{nome}: coluna obrigatoria ausente: {alternativas}. "
                     f"Colunas encontradas: {list(base.columns)}")

def _mapa_localizacao(base, chave, valor, nome):
    dados = pd.DataFrame({"chave": _texto_localizacao(base[chave]),
                          "valor": _texto_localizacao(base[valor])}).dropna().drop_duplicates()
    conflitos = dados.groupby("chave")["valor"].nunique()
    ambiguos = conflitos[conflitos.gt(1)].index
    if len(ambiguos):
        print(f"AVISO {nome}: chaves com mais de um resultado nao serao escolhidas "
              f"automaticamente: {', '.join(map(str, ambiguos))}")
    return dados.loc[~dados["chave"].isin(ambiguos)].drop_duplicates("chave").set_index("chave")["valor"]

def _uf_do_endereco_localizacao(endereco):
    if pd.isna(endereco):
        return pd.NA
    texto = str(endereco).upper().strip()
    # Somente UF no final do endereco, com CEP opcional.
    padrao = (r"(?:[,;/\-]|\s)\s*(" + "|".join(sorted(_UFS_TRANSITO))
              + r")\s*(?:[,;/\-]?\s*(?:CEP\s*[:\-]?\s*)?\d{5}-?\d{3})?\s*$")
    resultado = re.search(padrao, texto)
    return resultado.group(1) if resultado else pd.NA

def atualizar_localizacao_transito(rt):
    """Executar imediatamente apos a leitura, antes dos filtros e demais regras.

    As bases sao somente consultadas. Centros sao preenchidos no relatorio.
    Resultados ausentes/ambiguos ficam vazios e sao registrados no log.
    """
    rt = rt.copy()
    obrigatorias = {"cod_material_sap", "endereco", "cd_destino", "cd_origem"}
    faltantes = obrigatorias.difference(rt.columns)
    if faltantes:
        raise ValueError(f"Relatorio sem colunas obrigatorias: {sorted(faltantes)}")
    deposito = _base_localizacao("adhoc_deposito 1.xlsx")
    massivo = _base_localizacao("centros massivo.xlsx")
    centro = _base_localizacao("adhoc_centro 1.xlsx")
    regional = _base_localizacao("adhoc_regional 1.xlsx")

    # O anexo utiliza COD_CENTRO; aceitar tambem COD_TRANSITO.
    dep_chave = _coluna_localizacao(deposito, ["cod_transito", "cod_centro"], "deposito")
    dep_uf = _coluna_localizacao(deposito, ["sgl_uf", "uf"], "deposito")
    reg_valor = _coluna_localizacao(regional, ["regional", "centro oeste"], "regional")
    mapa_uf = _mapa_localizacao(deposito, dep_chave, dep_uf, "deposito")
    mapa_descricao = _mapa_localizacao(centro, "cod_centro", "dsc_centro_distribuicao", "centro")
    mapa_regional = _mapa_localizacao(regional, "uf", reg_valor, "regional")
    # Prioridades definidas: SP -> N190 e RJ -> N290.
    # Aplicadas apenas ao preenchimento de centros vazios.
    massivo_preferencial = massivo.loc[
        ~_texto_localizacao(massivo["uf"]).isin(["SP", "RJ"])
        | (_texto_localizacao(massivo["uf"]).eq("SP") & _texto_localizacao(massivo["centros"]).eq("N190"))
        | (_texto_localizacao(massivo["uf"]).eq("RJ") & _texto_localizacao(massivo["centros"]).eq("N290"))
    ].copy()
    mapa_massivo_uf = _mapa_localizacao(massivo_preferencial, "uf", "centros", "massivo UF -> CENTROS")
    mapa_massivo_cd = _mapa_localizacao(massivo, "centros", "cd", "massivo CENTROS -> CD")
    mapa_massivo_desc_uf = _mapa_localizacao(massivo_preferencial, "uf", "cd", "massivo UF -> CD")

    sap = _texto_localizacao(rt["cod_material_sap"]).str.lstrip("0")
    especial = sap.eq("41001597").fillna(False)
    destino = _texto_localizacao(rt["cd_destino"])
    origem = _texto_localizacao(rt["cd_origem"])

    # 1. UF_DESTINO: endereco apenas para 41001597; deposito para os demais.
    uf_destino = destino.map(mapa_uf).astype("string")
    uf_destino.loc[especial] = rt.loc[especial, "endereco"].map(_uf_do_endereco_localizacao)
    rt["uf_destino"] = uf_destino.where(uf_destino.isin(_UFS_TRANSITO), pd.NA)

    # 2. CD_DESTINO: preencher somente vazios, se TODOS forem do material especial.
    vazios = destino.isna()
    if vazios.any():
        if especial.loc[vazios].all():
            novos = rt.loc[vazios, "uf_destino"].map(mapa_massivo_uf)
            rt["cd_destino"] = rt["cd_destino"].astype("object")
            rt.loc[vazios, "cd_destino"] = novos
            print(f"Localizacao: {int(novos.notna().sum())} CD_DESTINO preenchido(s).")
        else:
            print(f"AVISO: existem {int((vazios & ~especial).sum())} CD_DESTINO vazio(s) "
                  "de outros materiais. Preenchimento massivo nao executado.")
    destino = _texto_localizacao(rt["cd_destino"])

    # 3. DESCRICAO_CD_DESTINO: massivo para o especial, centro para os demais.
    descricao = destino.map(mapa_descricao).astype("string")
    descricao_especial = destino.loc[especial].map(mapa_massivo_cd)
    # Sem centro: descricao por UF somente quando o nome do CD for inequivoco.
    sem_centro = destino.loc[especial].isna()
    descricao_especial.loc[sem_centro] = rt.loc[descricao_especial.index[sem_centro], "uf_destino"].map(mapa_massivo_desc_uf)
    descricao.loc[especial] = descricao_especial
    rt["descricao_cd_destino"] = descricao

    # 4. UF_ORIGEM: sempre pelo centro original na base de depositos.
    rt["uf_origem"] = origem.map(mapa_uf).astype("string")
    rt["uf_origem"] = rt["uf_origem"].where(rt["uf_origem"].isin(_UFS_TRANSITO), pd.NA)
    # 5. REGIONAL: sempre pela UF_DESTINO ja corrigida.
    rt["regional"] = rt["uf_destino"].map(mapa_regional)
    for coluna in ["uf_destino", "cd_destino", "descricao_cd_destino", "uf_origem", "regional"]:
        quantidade = int(_texto_localizacao(rt[coluna]).isna().sum())
        if quantidade:
            print(f"AVISO localizacao: {coluna.upper()} com {quantidade} linha(s) sem resultado.")
    return rt


# NF triangular: usa a mesma pasta configurada para as bases de localizacao.
PASTA_ZTMM028 = PASTA_BASES_LOCALIZACAO or PASTA_RELATORIO_TRANSITO
PREFIXO_ZTMM028 = "ZTMM028_"

def _normalizar_numero_ztmm(serie):
    return (serie.astype("string").str.strip()
            .str.replace(r"\.0$", "", regex=True)
            .replace({"": pd.NA, "NAN": pd.NA, "NONE": pd.NA,
                      "NULL": pd.NA, "ERRO": pd.NA, "#N/A": pd.NA,
                      "#VALUE!": pd.NA, "#REF!": pd.NA}))

def inserir_nf_triangular(df, arquivo_ztmm):
    """Insere NF TRIANGULAR e SERE NF TRIANGULAR no final, pela RESERVA."""
    ztmm = pd.read_excel(
        arquivo_ztmm,
        sheet_name="Data",
        usecols=["Nº reserva", "NF TRIANGULAR", "SERE NF TRIANGULAR"],
        dtype="string",
        engine="openpyxl"
    )
    ztmm.columns = ["reserva", "nf", "serie"]
    for coluna in ztmm.columns:
        ztmm[coluna] = _normalizar_numero_ztmm(ztmm[coluna])

    reserva_valida = ztmm["reserva"].str.fullmatch(r"\d+", na=False)
    nf_valida = (ztmm["nf"].str.fullmatch(r"\d+", na=False)
                 & ztmm["nf"].str.lstrip("0").ne(""))
    serie_valida = (ztmm["serie"].str.fullmatch(r"\d+", na=False)
                    & ztmm["serie"].str.lstrip("0").ne(""))

    ztmm = ztmm.loc[
        reserva_valida & nf_valida,
        ["reserva", "nf", "serie"]
    ].copy()
    ztmm.loc[~serie_valida.loc[ztmm.index], "serie"] = pd.NA
    ztmm = ztmm.drop_duplicates(["reserva", "nf", "serie"])

    def consolidar(grupo):
        # Se uma reserva tiver mais de uma NF, mantem todas sem duplicar o relatorio.
        nfs, series = [], []
        for nf, serie in grupo[["nf", "serie"]].itertuples(index=False, name=None):
            nfs.append(nf)
            series.append("" if pd.isna(serie) else serie)
        return pd.Series({
            "NF TRIANGULAR": "; ".join(nfs),
            "SERE NF TRIANGULAR": "; ".join(series).strip("; ")
        })

    registros = []
    for reserva, grupo in ztmm.groupby("reserva", sort=False):
        valores = consolidar(grupo)
        registros.append((reserva, valores["NF TRIANGULAR"], valores["SERE NF TRIANGULAR"]))
    mapa = pd.DataFrame(registros, columns=["reserva", "NF TRIANGULAR", "SERE NF TRIANGULAR"]).set_index("reserva")

    reserva_relatorio = _normalizar_numero_ztmm(df["RESERVA"])
    df["NF TRIANGULAR"] = reserva_relatorio.map(mapa["NF TRIANGULAR"])
    df["SERE NF TRIANGULAR"] = reserva_relatorio.map(mapa["SERE NF TRIANGULAR"])
    df["SERE NF TRIANGULAR"] = df["SERE NF TRIANGULAR"].replace("", pd.NA)

    print(
        f"NF triangular: {int(df['NF TRIANGULAR'].notna().sum())} linha(s) preenchida(s)."
    )
    return df

#pega a data atual
data_hoje = datetime.now().strftime("%Y-%m-%d")

#------------------------ conecta as planilhas ------------------------#
planilha_de_transito = arquivo_excel_mais_recente(
    pasta=PASTA_RELATORIO_TRANSITO,
    prefixo="relatorio_transito_empresarial_")
print("Arquivo utilizado:", planilha_de_transito)
rt = pd.read_excel(planilha_de_transito)
rt.columns = (rt.columns.str.lower())
rt = atualizar_localizacao_transito(rt)
rt = rt[
    (rt["dt_criacao_reserva"] >= "2025-01-01") &
    (rt["dt_criacao_reserva"] <  "2027-01-01")
]
rt["cod_material_sap"] = rt["cod_material_sap"].apply(limpar_codigo)
rt["num_item"] = rt["num_item"].apply(limpar_codigo)
rt["local_de_descarga"] = (rt["local_de_descarga"].where(rt["local_de_descarga"].notna(), None).astype(str).str.strip())
rt["chave"] = (rt["reserva"].fillna("").astype(str)+ rt["cod_material_sap"].fillna("").astype(str)+ rt["num_item"].fillna("").astype(str)) #convertendo as colunas em int e evitando de trazer "nan" quando for concatenar
planilha_de_transito = set(rt["local_de_descarga"].dropna())


planilha_consolidado_2026 = arquivo_excel_mais_recente(
    pasta=PASTA_CONSOLIDADO_RESID_2026,
    prefixo="Reservas Consolidadas 2026")
print("Arquivo utilizado:", planilha_consolidado_2026)
r2026 = pd.read_excel(planilha_consolidado_2026,sheet_name="BASE",header=0)
r2026.columns = (r2026.columns.str.lower().str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8').str.replace(r'[^\w\s]', '', regex=True).str.replace(' ', '_'))
r2026["item"] = r2026["item"].astype("string").str.strip()
r2026["chave"] = r2026["reserva"].fillna("").astype(str) + r2026["codigo_sap"].fillna("").astype(str) + r2026["item"].fillna("").astype(str)
r2026["chave"] = r2026["chave"].astype(str).str.replace(".0", "", regex=False)
#print(r2026.columns.tolist())
#print(r2026["chave"][:10])

planilha_consolidado_11sp = arquivo_excel_mais_recente(
    pasta=PASTA_RESERVAS_11SP,
    prefixo="reservas 2025-2026 - 11SP")
print("Arquivo utilizado:", planilha_consolidado_11sp)
r11sp = pd.read_excel(planilha_consolidado_11sp)
r11sp.columns = ( r11sp.columns.str.lower())

planilha_resid_2025 = arquivo_excel_mais_recente(
    pasta=PASTA_RESERVAS_RESID_2025,
    prefixo="Consolidado reservas 2025 - projetos")
print("Arquivo utilizado:", planilha_resid_2025)
r2025 = pd.read_excel(planilha_resid_2025,header=1)
r2025.columns = ( r2025.columns.str.lower())

planilha_projeto_pep = arquivo_excel_mais_recente(
    pasta=PASTA_PROJETOS_PEP,
    prefixo="projetos bi - app")
print("Arquivo utilizado:", planilha_projeto_pep)
pp = pd.read_excel(planilha_projeto_pep,header=1, usecols=lambda col: col != 0)
pp.columns = ( pp.columns.str.lower() )
pp.columns = pp.columns.str.replace(" ", "_")
#print(pp.columns.tolist())

planilha_siaki = arquivo_excel_mais_recente(
    pasta=PASTA_RESERVAS_SIAKI,
    prefixo="reserva_epo")
print("Arquivo utilizado:", planilha_siaki)
extensao = os.path.splitext(planilha_siaki)[1].lower()
if extensao == ".xlsx":
    siaki = pd.read_excel(planilha_siaki)
elif extensao == ".csv":
    siaki = pd.read_csv(planilha_siaki, sep=";", encoding="utf-8")
siaki["flag_siaki"] = "SIM"
siaki.columns = ( siaki.columns.str.lower())
COLUNAS_TEXTO = ["reserva"]
for col in COLUNAS_TEXTO:
        siaki[col] = (
            siaki[col].where(siaki[col].notna(), None).astype(str).str.strip().replace("", None))

planinha_atendimento = arquivo_excel_mais_recente(
    pasta=PASTA_PLAN_ATENDIMENTO_ID,
    prefixo="ATENDIMENTO ID")
print("Arquivo utilizado:", planinha_atendimento)
lm = pd.read_excel(planinha_atendimento,sheet_name="ATENDIMENTO ID(s)",header=5)
lm = lm.rename(columns={"N°RESERVA": "reserva"}) #altera o nome da coluna Nreserva para reserva
lm.columns = ( lm.columns.str.lower())
cols_texto = ["reserva", "status.1", "tran", "a01"]
lm[cols_texto] = lm[cols_texto].fillna("").astype(str)
lm["ids"] = (lm["ids"].where(lm["ids"].notna(), None).astype(str).str.strip())
lm["ids"] = "ID" + lm["ids"]

projetos = {  #insere os projetos de acordo com a esteira
        " A0":"NODE B",
        "A08":"NODE B",
        "A01":"LISTA DE MATERIAIS",
        "M01":"LISTA DE MATERIAIS",
        "A10":"NODE B",
        "A06":"ENGENHARIA/PLANTA",
        "A05":"ENGENHARIA/PLANTA",
        "A02":"LISTA DE MATERIAIS",
        "A07":"TELMEX",
        "A09":"NODE B",
        "A04":"PJE",
        "A08":"NODE B",
        "A03":"ENGENHARIA/PLANTA",
        "A08":"NODE B",
        "A11":"GPON Estruturado",
}
lm["a01"] = lm["a01"].map(projetos).fillna(lm["a01"])


planilha_cpe = arquivo_excel_mais_recente(
    pasta=PASTA_CPE,
    prefixo="RESERVAS CPE")
print("Arquivo utilizado:", planilha_cpe)
cpe = pd.read_excel(planilha_cpe, sheet_name="2026", converters={"Data Reserva": str})
cpe.columns = (cpe.columns.str.lower())
cpe = cpe.rename(columns={"no. reserva": "reserva"})
cpe["projeto"] = "CPE"
cpe["os cpe"] = (cpe["os cpe"].where(cpe["os cpe"].notna(), None).astype(str).str.strip())
cpe["os cpe"] = "OS" + cpe["os cpe"]

planilha_sobressalentes = arquivo_excel_mais_recente(
    pasta=PASTA_SOBRESSALENTES,
    prefixo="SOBRESSALENTES _")
print("Arquivo utilizado:", planilha_sobressalentes)
sobn = pd.read_excel(planilha_sobressalentes)
sobn.columns = ( sobn.columns.str.lower())
sobn = sobn.rename(columns={"status projeto": "projeto"})
sobn["reserva"] = (sobn["reserva"].where(sobn["reserva"].notna(), None).astype(str).str.strip())


planilha_LOGIN = arquivo_excel_mais_recente(
    pasta=LOGINS,
    prefixo="logins_p")
print("Arquivo utilizado:", planilha_LOGIN)
lg = pd.read_excel(planilha_LOGIN)
lg.columns = ( lg.columns.str.lower())
lg["usuario"] = (lg["usuario"].where(lg["usuario"].notna(), None).astype(str).str.strip())


planilha_reservas_a_eliminar = arquivo_excel_mais_recente(
    pasta=PASTA_RESERVAS_ELIMINAR,
    prefixo="reservas_para_")
print("Arquivo utilizado:", planilha_reservas_a_eliminar)
el = pd.read_excel(planilha_reservas_a_eliminar)
el.columns = (el.columns.str.lower())
el["reserva"] = (el["reserva"].where(el["reserva"].notna(), None).astype(str).str.strip())
reservas_eliminar = set(el["reserva"].dropna())

planilha_cliente_retira = arquivo_excel_mais_recente(
    pasta=PASTA_CLIENTE_RETIRA,
    prefixo="reservas_cliente")
print("Arquivo utilizado:", planilha_cliente_retira)
cr = pd.read_excel(planilha_cliente_retira)
cr.columns = (cr.columns.str.lower())
cr["reserva"] = (cr["reserva"].where(cr["reserva"].notna(), None).astype(str).str.strip())
reservas_cliente_retira = set(cr["reserva"].dropna())

planilha_parceiras= arquivo_excel_mais_recente(
    pasta=PASTA_PARCEIRAS,
    prefixo="parceiros")
print("Arquivo utilizado:", planilha_parceiras)
ld = pd.read_excel(planilha_parceiras)
ld.columns = (ld.columns.str.lower())
ld["loc. desc."] = (ld["loc. desc."].where(ld["loc. desc."].notna(), None).astype(str).str.strip())
locais_ld = set(ld["loc. desc."].dropna())

planilha_carga_fria = arquivo_excel_mais_recente(
    pasta=PASTA_CARGA_FRIA,
    prefixo="carga_fria")
print("Arquivo utilizado:", planilha_carga_fria)
cf = pd.read_excel(planilha_carga_fria, header=1)
cf.columns = (cf.columns.str.lower())
cf["reserva"] = (cf["reserva"].where(cf["reserva"].notna(), None).astype(str).str.strip())
reservas_cliente_retira = set(cf["reserva"].dropna())
#------------------------ fim conecta as planilhas ------------------------#

#insere o arquivo em uma variavel
arquivo_xlsx = Path(rf"C:\Users\F252727\OneDrive - Claro SA\Planilha de Atendimento\BI\BI_Relatorio_Transito\Dados_Transitos\Fatos\2026\novembro\relatorio_transito{data_hoje}.xlsx")

#normaliza todas as colunas RESERVAS das planilhas.
dfs = [rt, r2026, r11sp, r2025, cpe, sobn,lm,cf]
for df in dfs:
    df["reserva"] = df["reserva"].astype("string").str.strip()

#cria a coluna PEP GERAL e chama a função puxa os projetos da planilha pep geral
rt["projeto_geral"] = rt["pep_destino"].apply(buscar_projeto_geral)
#rt["projeto_geral"] = rt.pop("projeto_geral")

#--- alterar nome da coluna a01 para projeto_atendimento 
lm = lm.rename(columns={"a01": "projeto_atendimento"})

# busca atendimento de ID

padrao_reserva = re.compile(r"\b\d{6,}\b")

registros = []

for _, row in lm.iterrows():
    projeto = row["projeto_atendimento"]
    ids_val = row["ids"]

    for col in ["reserva", "status.1", "tran"]:
        encontrados = padrao_reserva.findall(row[col])

        for reserva in encontrados:
            registros.append((reserva, projeto, ids_val))


mapa_reserva_projeto = (
    pd.DataFrame(registros, columns=["reserva", "projeto_atendimento","ids"])
      .drop_duplicates(subset="reserva", keep="first")
)

mapa_reserva_projeto["reserva"] = (
    mapa_reserva_projeto["reserva"].astype("string").str.strip()
)

# buscar reservas de CPE-------------------

registros_cpe = []

for _, row in cpe.iterrows():
    os_cpe_val = row["os cpe"]
    projeto_cpe = row["projeto"]  # geralmente "CPE"

    encontrados = padrao_reserva.findall(row["reserva"])

    for reserva in encontrados:
        registros_cpe.append((reserva, projeto_cpe, os_cpe_val))

mapa_cpe = (
    pd.DataFrame(
        registros_cpe,
        columns=["reserva", "projeto_cpe", "os cpe"]
    )
    .drop_duplicates(subset="reserva", keep="first")
)

mapa_cpe["reserva"] = mapa_cpe["reserva"].astype("string").str.strip()

#-----------------------  DISTINCT ON ----------------------- #
#inserindo um equivalente ao distinct on, para puxar apenas o primeiro valor que aparecer de projeto de cada reserva
r2026 = r2026.drop_duplicates(subset="chave", keep="first")
r11sp = r11sp.drop_duplicates(subset="reserva", keep="first")
r2025 = r2025.drop_duplicates(subset="reserva", keep="first")
cpe   = cpe.drop_duplicates(subset="reserva", keep="first")
sobn  = sobn.drop_duplicates(subset="reserva", keep="first")
lm    = lm.drop_duplicates(subset="reserva", keep="first")
lg    = lg.drop_duplicates(subset="usuario", keep="first")
siaki = siaki.drop_duplicates(subset="reserva", keep="first")
cf    = cf.drop_duplicates(subset="reserva", keep="first")
#-----------------------  fim distinct on ----------------------- #

#-----------------------  renomeando as colunas "PROJETOS" ----------------------- #
r2026 = r2026.rename(columns={"projeto": "projeto_2026"})
r11sp = r11sp.rename(columns={"projeto": "projeto_11sp"})
r2025 = r2025.rename(columns={"projeto": "projeto_2025"})
#lm = lm.rename(columns={"a01": "projeto_atendimento"})
cpe   = cpe.rename(columns={"projeto": "projeto_cpe"})
sobn  = sobn.rename(columns={"projeto": "projeto_sobn"})
cf    = cf.rename(columns={"projeto": "projeto_cf"})
#-------------- aplicando as merger ------------------#
#equivalente ao left join para unir os dados
df = rt.merge(r2026[["chave", "projeto_2026", "cidade_projeto", "npipefy"]], on="chave", how="left")
df = df.merge(r11sp[["reserva", "projeto_11sp"]], on="reserva", how="left")
df = df.merge(r2025[["reserva", "projeto_2025"]], on="reserva", how="left")
df = df.merge(mapa_reserva_projeto,on="reserva",how="left")
df = df.merge(mapa_cpe,on="reserva",how="left")
#df = df.merge(cpe[["reserva", "projeto_cpe","os cpe"]], on="reserva", how="left")
df = df.merge(sobn[["reserva", "projeto_sobn"]], on="reserva", how="left")
df = df.merge(lg[["usuario", "nome_completo", "area"]], left_on="criado_pelo_usuario", right_on="usuario", how="left")
df = df.merge(siaki[["reserva", "flag_siaki"]],on="reserva",how="left")
df = df.merge(cf[["reserva", "projeto_cf"]], on="reserva", how="left")
#-------------- fim das merger ------------------#

# SAP 41001597: manter somente linhas da area PLANEJAMENTO

sap = (df["cod_material_sap"].astype("string").str.strip().str.replace(r"\.0$", "", regex=True).str.lstrip("0"))
area = df["area"].astype("string").str.strip().str.upper()
df = df.loc[sap.ne("41001597").fillna(True)| area.eq("PLANEJAMENTO").fillna(False)].copy()


#------- aplica busca no atendimento_i
#df[["projeto_lm"]] = df["reserva"].apply(buscar_atendimento)

#----------- unificando os projetos em uma unico coluna---------------#
df["projeto_unificado"] = (
    df["projeto_2026"]
    .combine_first(df["projeto_11sp"])
    .combine_first(df["projeto_2025"])
    #.combine_first(df["projeto_lm"])
    .combine_first(df["projeto_atendimento"])
    .combine_first(df["projeto_cpe"])
    .combine_first(df["projeto_sobn"])
    .combine_first(df["projeto_cf"])     
)

df["area"] = df["area"].fillna("OUTROS") #criando a coluna area
df["nome_completo"] = df["nome_completo"].fillna("") #criando a coluna nome usuario



#------------- case when ----------------#
#abaixo de 2026 se for de planejamento projeto unificado vira projeto geral
df["ANO_CRIACAO"] = pd.to_datetime(
    pd.to_datetime(df["dt_criacao_reserva"], errors="coerce")
).dt.year

mask_fallback = (
    (df["ANO_CRIACAO"] < 2026) |
    (
        (df["ANO_CRIACAO"] >= 2026) &
        (df["area"] != "PLANEJAMENTO")
    )
)

# aplica fallback

df.loc[
    df["projeto_unificado"].isna() & mask_fallback,
    "projeto_unificado"
] = df.loc[
    df["projeto_unificado"].isna() & mask_fallback,
    "projeto_geral"
]


df["area"] = df["area"].fillna("OUTROS") #criando a coluna area
df["nome_completo"] = df["nome_completo"].fillna("") #criando a coluna nome usuario

#------- transformando as colunas em datas e str para case when da coluna SIAK -------#
df["dt_criacao_reserva"] = pd.to_datetime(
    df["dt_criacao_reserva"], errors="coerce"
)

df["pep_destino"] = df["pep_destino"].astype("string")

#print(lm.columns.tolist())

#----- aplicando case when da coluna siak ----#
df["SIAKI"] = "NAO" 

df.loc[
    (df["dt_criacao_reserva"] <= "2025-08-31") |
    (df["pep_destino"].str.endswith("YY", na=False)) |
    (df["flag_siaki"] == "SIM"),
    "SIAKI"
] = "SIM"

# cria ID OS
#df["id_os_pipefy"] = df["os cpe"].combine_first(df["ids"])

# cria ID OS
df["id_os_pipefy"] = (
    df["os cpe"]
    .combine_first(df["ids"])
    .combine_first(df["npipefy"])
)

df["cidade_projeto"] = (df["cidade_projeto"].combine_first(df["cidade"]))

mask_projeto_vazio = (
    df["projeto_unificado"].isna() | (df["projeto_unificado"].astype(str).str.strip() == ""))

mask_area = df["area"] != "PLANEJAMENTO"

df.loc[mask_projeto_vazio & mask_area, "projeto_unificado"] = df.loc[mask_projeto_vazio & mask_area, "projeto_geral"]

#--- mover colunas

posicoes = {
        "id_os_pipefy": 0,
        "uf_origem": 1,
        "cd_origem": 2,
        "projeto_unificado": 3,
        "reserva": 4,
        "num_item": 5,
        "local_de_descarga": 6,
        "dt_criacao_reserva": 7,
        "cd_destino": 8,
        "cod_material_sap": 9,
        "descricao": 10,
        "qtd": 11,
        "cidade_projeto": 12,
        "endereco": 13,
        "uf_destino": 14,
        "nº_nf-e": 15,
        "serie": 16,
        "data_emissao_nf": 17,
        "protocolo": 18,
        "romaneio": 19,
        "transportadora": 20,
        "expedicao": 21,
        "dsc_ultima_ocorrencia": 22,
        "previsao_de_entrega": 23,
        "dt_limite_entrega": 24,
        "nova_data_planejada": 25,
        "dt_entrega_realizada": 26,
        "situacao": 27,
        "situacao_datas_operacao": 28,
        "criado_pelo_usuario": 29,
        "nome_completo": 30,
        "area": 31,
        "SIAKI": 32,
        "pep_destino": 33
}

cols = df.columns.tolist()

# remove as colunas 
for col in posicoes:
    if col in cols:
        cols.remove(col)

# insere a coluna nas posições 
for col, pos in sorted(posicoes.items(), key=lambda x: x[1]):
    cols.insert(pos, col)

df = df[cols]

# Eliminar colunas auxiliares 

df = df.drop('projeto_2026', axis=1)
df = df.drop('projeto_11sp', axis=1)
#df = df.drop('projeto_lm', axis=1)
df = df.drop('projeto_cpe', axis=1)
df = df.drop('projeto_sobn', axis=1)
df = df.drop('usuario', axis=1)
df = df.drop('ANO_CRIACAO', axis=1)
df = df.drop('projeto_2025', axis=1)
df = df.drop('flag_siaki', axis=1)
df = df.drop('os cpe', axis=1)
df = df.drop('ids', axis=1)
df = df.drop('projeto_atendimento', axis=1)
df = df.drop('projeto_cf', axis=1)

#remover as reservas que estão na planilha de eliminar reserva
#df = df[~df["reserva"].isin(reservas_eliminar)]

#coloca as reservas da planilha cliente retira o status cliente retira na base principal
df.loc[df["reserva"].isin(reservas_cliente_retira),"transportadora"] = "CLIENTE RETIRA"


#caso tenha alguma reserva em duplicidade
el = el.drop_duplicates(subset="reserva", keep="first")

#ENCONTRA A RESERVA NA PLANILHA DE ELIMINAR RESERVA E COLOCA O STATUS ENTREGA NA BASE PRINCIPAL
df.loc[
    df["reserva"].isin(el["reserva"]),
    "situacao"
] = "ENTREGA"

# quando o projeto unificado estiver fazio e a reserva estiver na planilha de eliminar reserva, o projeto unificado vai receber o valor do projeto geral
mask_reserva_eliminar = df["reserva"].isin(reservas_eliminar)

mask_projeto_vazio = (
    df["projeto_unificado"].isna() |
    (df["projeto_unificado"].astype(str).str.strip() == "")
)

df.loc[
    mask_reserva_eliminar & mask_projeto_vazio,
    "projeto_unificado"
] = df.loc[
    mask_reserva_eliminar & mask_projeto_vazio,
    "projeto_geral"
]

#pagar colunas
df = df.drop('id registro', axis=1)
df = df.drop('npipefy', axis=1)
df = df.drop('chave', axis=1)
df = df.drop('cod_empresa', axis=1)
df = df.drop('descricao_empresa', axis=1)
#df = df.drop('regional', axis=1) #---
# DESCRICAO_CD_DESTINO mantida: preenchida na etapa inicial de localizacao.
#df = df.drop('cidade', axis=1) #----
df = df.drop('projeto_geral', axis=1)
df = df.drop('descricao_cd_origem', axis=1)
df = df.drop('lead_time', axis=1)


df["projeto_unificado"] = (df["projeto_unificado"].str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("utf-8").str.replace(r"\s+", " ", regex=True).str.upper() )#padroniza a coluna projeto unificado
df["cidade_projeto"] = (df["cidade_projeto"].str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("utf-8").str.replace(r"\s+", " ", regex=True).str.upper()) #padroniza a coluna projeto geral
df = df.fillna("") #retira o NaN
# Prioridade final: todo material 41001597 pertence ao projeto GPON Resid.
# Aplicar depois da padronizacao para preservar exatamente o texto solicitado.
mask_gpon_resid = (
    df["cod_material_sap"].astype("string").str.strip()
    .str.replace(r"\.0$", "", regex=True).str.lstrip("0")
    .eq("41001597").fillna(False)
)
df.loc[mask_gpon_resid, "projeto_unificado"] = "GPON Resid"

df.columns = (df.columns.str.upper())#transformar colunas em maiusculo
df.columns = df.columns.str.replace("_", " ").str.upper()#remove _ e deixa os espaços
planilha_ztmm028 = arquivo_excel_mais_recente(PASTA_ZTMM028, PREFIXO_ZTMM028)
print("Arquivo utilizado para NF triangular:", planilha_ztmm028)
df = inserir_nf_triangular(df, planilha_ztmm028)

df = df.sort_values(by=["DT CRIACAO RESERVA", "RESERVA"],ascending=[False, True]) #ordena por data de criaçao da reserva  

# -------- salvar arquivo de reserva e projeto --------
df_resumo = df[["RESERVA", "PROJETO UNIFICADO"]].copy()
#rome duplicadas
df_resumo = df_resumo.drop_duplicates(subset="RESERVA")

#seleciona a pasta
arquivo_resumo = Path(
    f"C:\\Users\\F252727\\OneDrive - Claro SA\\Planilha de Atendimento\\reserva_e_projeto.xlsx"
)
df_resumo.to_excel(arquivo_resumo, index=False)

#df_resumo.to_csv(arquivo_resumo.with_suffix(".csv"), index=False, sep=";") --- caso seja nescessario arquivo .csv

print("Arquivo Reserva e Projeto salvo!")
#----- fim salvar arquivo de reserva e projeto ---------


""" # --- abre Excel ---
excel = win32.Dispatch("Excel.Application")
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(str(arquivo_xlsx_temp)) """

#------ trata as formatacões


#coloca na coluna SITUACAO DATAS OPERACAO "atrasados" tudo que for expedido numa data maior que dez dia. 
mask_atraso = (
    (
        pd.Timestamp.today() -
        pd.to_datetime(df["DATA EMISSAO NF"], errors="coerce")
    ).dt.days > 10
) & (
    df["SITUACAO DATAS OPERACAO"].isna() |
    (df["SITUACAO DATAS OPERACAO"].astype(str).str.strip() == "") |
    (df["SITUACAO DATAS OPERACAO"] == "AGUARDANDO COLETA/ATUALIZAÇÃO SISTÊMICA")
) & (
    df["TRANSPORTADORA"] != "CLIENTE RETIRA"
) & (
    df["DT LIMITE ENTREGA"].isna() | (df["DT LIMITE ENTREGA"].astype(str).str.strip() == "")
)

df.loc[mask_atraso, "SITUACAO DATAS OPERACAO"] = "ATRASADO"

# coloca atraso para as reservas que não foram faturadas entro de 10 dias
mask_atraso = (
    (
        pd.Timestamp.today() -
        pd.to_datetime(df["DT CRIACAO RESERVA"], errors="coerce")
    ).dt.days > 10
) & (
    df["DATA EMISSAO NF"].isna() |
    (df["DATA EMISSAO NF"].astype(str).str.strip() == ""))

df.loc[mask_atraso, "SITUACAO DATAS OPERACAO"] = "ATRASADO"


#tratar as datas 
Datas=["DT CRIACAO RESERVA", "DATA EMISSAO NF", "EXPEDICAO", "PREVISAO DE ENTREGA","DT LIMITE ENTREGA","NOVA DATA PLANEJADA","DT ENTREGA REALIZADA"]

for col in Datas:
    df[col] = pd.to_datetime(df[col],errors="coerce").dt.strftime("%d/%m/%Y")
#--fim


# Status que será reclassificado
mask_aguardando = (
    df["SITUACAO"].astype(str).str.strip()
    == "AGUARDANDO COLETA/ATUALIZAÇÃO SISTÊMICA"
)

# Verifica se DATA EMISSAO NF está vazia
mask_sem_nota = (
    df["DATA EMISSAO NF"].isna() |
    (df["DATA EMISSAO NF"].astype(str).str.strip() == "")
)

# Sem data da nota fiscal: EM SEPARACAO
df.loc[
    mask_aguardando & mask_sem_nota,
    "SITUACAO"
] = "EM SEPARACAO"

# Com data da nota fiscal: EM EXPEDICAO
df.loc[
    mask_aguardando & ~mask_sem_nota,
    "SITUACAO"
] = "EM EXPEDICAO"


df_fora_siaki = df[
    (df["SIAKI"] == "NAO") &
    (df["LOCAL DE DESCARGA"].isin(locais_ld))
]


df_sem_projeto = df[
    (df["PROJETO UNIFICADO"] == "") &
    (df["AREA"]== "PLANEJAMENTO")]

# Agrupa por nome e conta reservas distintas
df_resumo_sem_projeto = (
    df_sem_projeto
    .assign(
        **{
            "NOME COMPLETO": df_sem_projeto["NOME COMPLETO"]
            .replace("", "USUARIO NAO IDENTIFICADO")
            .fillna("USUARIO NAO IDENTIFICADO")
        }
    )
    .groupby("NOME COMPLETO", as_index=False)
    .agg(
        **{
            "QTD RESERVAS DISTINTAS": (
                "RESERVA",
                "nunique"
            )
        }
    )
    .sort_values(
        by="QTD RESERVAS DISTINTAS",
        ascending=False
    )
)


df_ATRASO = df[
    (
        (df["SITUACAO"] == "TRANSITO") |
        (df["SITUACAO"] == "AGUARDANDO COLETA/ATUALIZAÇÃO SISTÊMICA")
        ) &
        (df["SITUACAO DATAS OPERACAO"] == "ATRASADO") &
        (df["AREA"] == "PLANEJAMENTO")]
#df_ATRASO["DIAS SEM EXPEDICAO"] = (pd.Timestamp.today() - pd.to_datetime(df_ATRASO["DATA EMISSAO NF"],errors="coerce")).dt.days


# Aplicar na coluna "ID OS PIPEFY" a função separar_identificador e criar as colunas "TIPO_IDENTIFICADOR", "NUMERO_IDENTIFICADOR" e "CHAVE_CONSULTA" 
df[["TIPO_IDENTIFICADOR", "NUMERO_IDENTIFICADOR", "CHAVE_CONSULTA"]] = (
    df["ID OS PIPEFY"].apply(separar_identificador)
)

df["ID OS PIPEFY"] = (df["ID OS PIPEFY"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip())

#arquivo_identificador = Path(
#    f"C:\\Users\\F252727\\OneDrive - Claro SA\\Planilha de Atendimento\\base_relatorio_transito.xlsx"
#)
#df.to_excel(arquivo_identificador, index=False)

#print("Base para Copilot Salvo!")

# Manter as duas colunas de triangulacao no final, inclusive apos os identificadores.
colunas_triangular = ["NF TRIANGULAR", "SERE NF TRIANGULAR"]
df = df[[c for c in df.columns if c not in colunas_triangular] + colunas_triangular]

#criar um novo arquivo xlsx
writer = pd.ExcelWriter(arquivo_xlsx, engine='xlsxwriter')

#aba principal
df.to_excel(writer, sheet_name='RELATORIO', index=False)

if not df_fora_siaki.empty:
    df_fora_siaki.to_excel(writer, sheet_name="FORA SIAKI", index=False)

if not df_sem_projeto.empty:
    df_sem_projeto.to_excel(writer, sheet_name="SEM PROJETO", index=False)

if not df_resumo_sem_projeto.empty:
    df_resumo_sem_projeto.to_excel(writer,sheet_name="RESUMO SEM PROJETO",index=False)    
        
if not df_ATRASO.empty:
    df_ATRASO.to_excel(writer, sheet_name="ATRASADOS", index=False)

workbook  = writer.book
    
if "RELATORIO" in writer.sheets:
    formatacao_excel(workbook,writer.sheets["RELATORIO"],df)
if "FORA SIAKI" in writer.sheets:
    formatacao_excel(workbook,writer.sheets["FORA SIAKI"],df_fora_siaki)
if "SEM PROJETO" in writer.sheets:
    formatacao_excel(workbook,writer.sheets["SEM PROJETO"],df_sem_projeto)
if "ATRASADOS" in writer.sheets:
    formatacao_excel(workbook,writer.sheets["ATRASADOS"],df_ATRASO)    
if "RESUMO SEM PROJETO" in writer.sheets:
    formatacao_excel(workbook,writer.sheets["RESUMO SEM PROJETO"],df_resumo_sem_projeto)

writer.close()

arquivo_xlsx = Path(arquivo_xlsx)#abrir xlsx para salvar em xlsb
arquivo_xlsb = Path(rf"C:\Users\F252727\OneDrive - Claro SA\Planilha de Atendimento\Arquivos Relatorio de Transito\relatorio_transito_{data_hoje}.xlsb" )
# Se o arquivo já existir, remove antes de gerar o novo
excel = win32.Dispatch("Excel.Application") #abrir excel para salvar em xlsb
try: #verificar se o excel está aberto, caso contrário, continuar sem mostrar o excel
    excel.Visible = False #manter o excel oculto durante a manipulação
except Exception:
    pass
excel.DisplayAlerts = False #desativar alertas para evitar pop-ups

#Valida se existem as abas antes de acessar para alterar as datas 
def get_ws(wb, nome_aba):
    for ws in wb.Worksheets:
        if ws.Name == nome_aba:
            return ws
    return None

wb = excel.Workbooks.Open(str(arquivo_xlsx)) #abrir o arquivo xlsx
ws = get_ws(wb, "RELATORIO")
wsiaki = get_ws(wb, "FORA SIAKI")
wssemprojeto = get_ws(wb, "SEM PROJETO")
wsatraso = get_ws(wb, "ATRASADOS")
#--- fim da validação 
    
# colunas substituir (ex: datas)
colunas = ["DT CRIACAO RESERVA", "DATA EMISSAO NF","EXPEDICAO","PREVISAO DE ENTREGA","DT LIMITE ENTREGA","NOVA DATA PLANEJADA","DT ENTREGA REALIZADA"]
#substituir o - por / nas colunas datas de todas as abas
for col_nome in colunas:
    col_idx = ws.Rows(1).Find(col_nome).Column
    rng = ws.Range(ws.Cells(2, col_idx), ws.Cells(ws.UsedRange.Rows.Count, col_idx))
    rng.Replace("-", "/")
for col_nome in colunas:
    col_idx = wsiaki.Rows(1).Find(col_nome).Column
    rng = wsiaki.Range(wsiaki.Cells(2, col_idx), wsiaki.Cells(wsiaki.UsedRange.Rows.Count, col_idx))
    rng.Replace("-", "/")
for col_nome in colunas:
    col_idx = wssemprojeto.Rows(1).Find(col_nome).Column
    rng = wssemprojeto.Range(wssemprojeto.Cells(2, col_idx), wssemprojeto.Cells(wssemprojeto.UsedRange.Rows.Count, col_idx))
    rng.Replace("-", "/")
for col_nome in colunas:
    col_idx = wsatraso.Rows(1).Find(col_nome).Column
    rng = wsatraso.Range(wsatraso.Cells(2, col_idx), wsatraso.Cells(wsatraso.UsedRange.Rows.Count, col_idx))
    rng.Replace("-", "/")


#----- salva em xlxb
wb.SaveAs(str(arquivo_xlsb), FileFormat=50)  # 50 = XLSB
wb.Close(False)
excel.Quit()
#arquivo_xlsx.unlink(missing_ok=True)


#apaga XLSX temporário
""" arquivo_xlsx_temp.unlink(missing_ok=True) """


duracao = time.time() - inicio
minutos = int(duracao // 60) 
segundos = int(duracao % 60)
print(f"Processamento: {minutos:.2f} minutos")


#279022 - python
#278470 - banco

