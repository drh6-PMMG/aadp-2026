"""
Script de Análise AADP 2026 — Versão Final
Baseado no modelo Geral.xlsx atualizado (01/07/2026).

CORREÇÕES APLICADAS:
- Encoding cp1252 (não UTF-8)
- normaliza() para match robusto dos conceitos (sem acentos)
- Status: "Aberta", "Encerrada", "Parcialmente Encerrada", "Homologação"
- Coluna "Certificação Homologador" (col 13 do modelo)
- Coluna "Situação Comissão" (não "Tipo de Comissão")
- Aba Avaliadores Pendentes: AV1 + AV2 combinados + seção Homologador
"""
import csv, sys, io, os, zipfile, unicodedata
from datetime import datetime
from collections import defaultdict, Counter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

try:
    import xlsxwriter
    print("xlsxwriter OK")
except ImportError as e:
    print(f"ERRO: {e}"); sys.exit(1)

# ──────────────────────────── CONFIGURAÇÕES ───────────────────────────────────
BASE_DIR        = r"c:\Users\guilh\Downloads\analise AADP 2026"
AVALIACOES_FILE = os.path.join(BASE_DIR, "avaliacoes.csv")
SIGEF_FILE      = os.path.join(BASE_DIR, "SIGEF.csv")
OUT_DIR         = os.path.join(BASE_DIR, "Resultado_AADP_2026")
os.makedirs(OUT_DIR, exist_ok=True)

SITUACOES_ALVO = {
    "ATIV. DIRECAO GERAL", "ATIV. FIM DESTACADO", "ATIV. FIM NA SEDE",
    "ATIV. MEIO", "ATIVIDADE MEIO", "DISP MED DEFINITIVA",
    "GESTANTE/LAC/ADOTANT", "QUADRO ESPECIALISTA",
}

# ─────────────────────── LÓGICA DE PARIDADE ────────────────────────────────
# Resolução 5458/2025 — chaves SEM acentos para match robusto
CONCEITO_FAIXA = {
    "nivel superior de desempenho":       (9.00, 10.00),
    "nivel alto de desempenho":           (7.00,  8.99),
    "nivel intermediario de desempenho":  (6.00,  6.99),
    "nivel baixo de desempenho":          (3.00,  5.99),
    "nivel inferior de desempenho":       (0.00,  2.99),
}

def normaliza(texto: str) -> str:
    """Remove acentos e coloca em minúsculas — match robusto independente de encoding."""
    t = unicodedata.normalize("NFD", texto.lower())
    return "".join(c for c in t if unicodedata.category(c) != "Mn")

def is_empty(v) -> bool:
    return not v or str(v).strip() in ("", "-")

def concordam(conceito: str, nota_str: str):
    """True=concordam, False=divergem, None=dados insuficientes."""
    if is_empty(conceito) or is_empty(nota_str):
        return None
    try:
        nota = float(str(nota_str).replace(",", "."))
    except ValueError:
        return None
    faixa = CONCEITO_FAIXA.get(normaliza(conceito.strip()))
    if faixa is None:
        return None
    return faixa[0] <= nota <= faixa[1]

def calc_cert_hom(j: str, l: str) -> str:
    """
    Certificação Homologador:
    SIM  = há divergência entre Conceito AV1 (J) e Nota AV2 (L)
    NÃO  = não há divergência
    -    = avaliação incompleta (J ou L ausentes)
    """
    if is_empty(j) or is_empty(l):
        return "-"
    c = concordam(j, l)
    if c is True:  return "NÃO"
    if c is False: return "SIM"
    return "-"

def calc_status(j: str, l: str, n: str) -> str:
    """
    Encerrada            = J+L concordam OU (J+L divergem + N presente)
    Homologação          = J+L presentes + divergência + N ausente
    Parcialmente Encerrada = J presente + L ausente
    Aberta               = J ausente
    """
    has_j = not is_empty(j)
    has_l = not is_empty(l)
    has_n = not is_empty(n)

    if not has_j:
        return "Aberta"
    if has_j and not has_l:
        return "Parcialmente Encerrada"
    # has_j e has_l
    c = concordam(j, l)
    if c is True:
        return "Encerrada"
    elif c is False:
        return "Encerrada" if has_n else "Homologação"
    else:
        return "Encerrada"

# ─────────────────────── LEITURA SIGEF ────────────────────────────────────
print("Carregando SIGEF.csv ...")
sigef_unidade: dict[str, str] = {}
with open(SIGEF_FILE, encoding="cp1252", errors="replace") as f:
    reader = csv.reader(f, delimiter=";")
    next(reader)
    for row in reader:
        if len(row) > 9:
            nrpm = row[0].strip().lstrip("0") or "0"
            sigef_unidade[nrpm] = row[9].strip()
print(f"  {len(sigef_unidade):,} registros no SIGEF.")

# ─────────────────────── LEITURA AVALIAÇÕES ───────────────────────────────
print("Carregando avaliacoes.csv ...")
records: list[dict] = []

with open(AVALIACOES_FILE, encoding="cp1252", errors="replace") as f:
    reader = csv.reader(f, delimiter=";")
    next(reader)
    for row in reader:
        while len(row) < 42:
            row.append("")
        sit = row[7].strip()
        if sit not in SITUACOES_ALVO:
            continue

        nrpm_av  = row[0].strip()
        local_av = row[5].strip()
        j        = row[9].strip()
        l        = row[11].strip()
        n        = row[13].strip()

        # Situação Comissão via SIGEF
        nrpm_norm  = nrpm_av.lstrip("0") or "0"
        sigef_unid = sigef_unidade.get(nrpm_norm, "")
        sit_comissao = (
            "Comissão Atual"
            if local_av.upper().strip() == sigef_unid.upper().strip()
            else "Nota Provisória"
        )

        cert_hom = calc_cert_hom(j, l)
        status   = calc_status(j, l, n)

        records.append({
            # Avaliado
            "nrPM_av": nrpm_av,       "nome_av":  row[1].strip(),
            "posto_av": row[2].strip(),"rpm_av":   row[3].strip(),
            "unid_princ_av": row[4].strip(), "local_av": local_av,
            "quadro_av": row[6].strip(), "sit_av": sit,
            # Avaliação
            "data_av1": row[8].strip(),  "conceito": j,
            "data_av2": row[10].strip(), "nota_geral": l,
            "cert_hom": cert_hom,
            "data_hom": row[12].strip(), "nota_hom": n,
            # Competências
            "comp1": row[14].strip(), "conc_comp1": row[15].strip(), "nota_comp1": row[16].strip(),
            "comp2": row[17].strip(), "conc_comp2": row[18].strip(), "nota_comp2": row[19].strip(),
            "comp3": row[20].strip(), "conc_comp3": row[21].strip(), "nota_comp3": row[22].strip(),
            "comp4": row[23].strip(), "conc_comp4": row[24].strip(), "nota_comp4": row[25].strip(),
            # Avaliador 1
            "nrpm_av1": row[26].strip(), "nome_av1":  row[27].strip(),
            "posto_av1": row[28].strip(), "rpm_av1":  row[29].strip(),
            "unid_princ_av1": row[30].strip(), "local_av1": row[31].strip(),
            "quadro_av1": row[32].strip(), "sit_av1":  row[33].strip(),
            # Avaliador 2
            "nrpm_av2": row[34].strip(), "nome_av2":  row[35].strip(),
            "posto_av2": row[36].strip(), "rpm_av2":  row[37].strip(),
            "unid_princ_av2": row[38].strip(), "local_av2": row[39].strip(),
            "quadro_av2": row[40].strip(), "sit_av2":  row[41].strip(),
            # Calculados
            "sit_comissao": sit_comissao,
            "status":       status,
        })

print(f"  {len(records):,} registros filtrados.")
sc = Counter(r["status"] for r in records)
cc = Counter(r["cert_hom"] for r in records)
print(f"  Status:       {dict(sc)}")
print(f"  Cert. Hom.:   {dict(cc)}")

# ─────────────────────── HEADERS ──────────────────────────────────────────
# 45 colunas — exatamente como no modelo Geral.xlsx
HEADERS_GERAL = [
    "nrPM (Avaliado)", "Nome Completo (Avaliado)", "Posto/Graduação (Avaliado)",
    "Unidade RPM (Avaliado)", "Unidade Principal (Avaliado)", "Local/Unidade (Avaliado)",
    "Quadro Atual (Avaliado)", "Situação Funcional Atual (Avaliado)",
    "Data Avaliação 1", "Conceito Geral",
    "Data Avaliação 2", "NotaGeral",
    "Certificação Homologador",   # col 13 — nova
    "Data Homologação", "Nota Homologação",
    "Competência 1", "Conceito (Competência 1)", "Nota (Competência 1)",
    "Competência 2", "Conceito (Competência 2)", "Nota (Competência 2)",
    "Competência 3", "Conceito (Competência 3)", "Nota (Competência 3)",
    "Competência 4", "Conceito (Competência 4)", "Nota (Competência 4)",
    "nrPM (Avaliador1)", "Nome Completo (Avaliador1)", "Posto/Graduação (Avaliador1)",
    "Unidade RPM Atual (Avaliador1)", "Unidade Principal Atual (Avaliador1)",
    "Local/Unidade Atual (Avaliador1)", "Quadro Atual (Avaliador1)",
    "Situação Funcional Atual (Avaliador1)",
    "nrPM (Avaliador2)", "Nome Completo (Avaliador2)", "Posto/Graduação (Avaliador2)",
    "Unidade RPM Atual (Avaliador2)", "Unidade Principal Atual (Avaliador2)",
    "Local/Unidade Atual (Avaliador2)", "Quadro Atual (Avaliador2)",
    "Situação Funcional Atual (Avaliador2)",
    "Situação Comissão",   # col 44
    "Status Avaliação",    # col 45
]

COL_CERT   = HEADERS_GERAL.index("Certificação Homologador")    # 12 (0-based)
COL_SIT    = HEADERS_GERAL.index("Situação Comissão")           # 43
COL_STATUS = HEADERS_GERAL.index("Status Avaliação")            # 44

def rec_to_row(r: dict) -> list:
    return [
        r["nrPM_av"],   r["nome_av"],  r["posto_av"],
        r["rpm_av"],    r["unid_princ_av"], r["local_av"],
        r["quadro_av"], r["sit_av"],
        r["data_av1"],  r["conceito"],
        r["data_av2"],  r["nota_geral"],
        r["cert_hom"],
        r["data_hom"],  r["nota_hom"],
        r["comp1"], r["conc_comp1"], r["nota_comp1"],
        r["comp2"], r["conc_comp2"], r["nota_comp2"],
        r["comp3"], r["conc_comp3"], r["nota_comp3"],
        r["comp4"], r["conc_comp4"], r["nota_comp4"],
        r["nrpm_av1"], r["nome_av1"], r["posto_av1"],
        r["rpm_av1"],  r["unid_princ_av1"], r["local_av1"],
        r["quadro_av1"], r["sit_av1"],
        r["nrpm_av2"], r["nome_av2"], r["posto_av2"],
        r["rpm_av2"],  r["unid_princ_av2"], r["local_av2"],
        r["quadro_av2"], r["sit_av2"],
        r["sit_comissao"],
        r["status"],
    ]

# ─────────────────────── CORES ────────────────────────────────────────────
STATUS_BG = {
    "Encerrada":             "#70AD47",
    "Homologação":           "#FFD966",
    "Parcialmente Encerrada":"#FF8C00",
    "Aberta":                "#FF0000",
}
STATUS_FG = {
    "Encerrada":             "#000000",
    "Homologação":           "#000000",
    "Parcialmente Encerrada":"#000000",
    "Aberta":                "#FFFFFF",
}
SIT_BG = {"Comissão Atual": "#4472C4", "Nota Provisória": "#FFFF00"}
SIT_FG = {"Comissão Atual": "#FFFFFF", "Nota Provisória": "#000000"}

COL_WIDTHS = [
    10, 35, 14, 16, 28, 38, 10, 24,
    14, 32, 14, 10,
    22,           # Certificação Homologador
    14, 12,
    20, 30, 10, 20, 30, 10, 30, 30, 10, 20, 30, 10,
    10, 35, 14, 16, 28, 38, 10, 24,
    10, 35, 14, 16, 28, 38, 10, 24,
    20, 24,       # Situação Comissão, Status Avaliação
]

# ─────────────────────── FORMATOS ─────────────────────────────────────────
def make_formats(wb):
    f = {}
    base = {"font_name": "Calibri", "font_size": 10, "border": 1, "valign": "vcenter"}

    f["hdr"]     = wb.add_format({**base, "bold": True, "bg_color": "#1F3864",
                                   "font_color": "#FFFFFF", "align": "center", "text_wrap": True})
    f["hdr_av"]  = wb.add_format({**base, "bold": True, "bg_color": "#2E4057",
                                   "font_color": "#FFFFFF", "align": "center", "text_wrap": True})
    f["hdr_sub"] = wb.add_format({**base, "bold": True, "bg_color": "#395474",
                                   "font_color": "#FFFFFF", "align": "center", "text_wrap": True})
    f["hdr_hom"] = wb.add_format({**base, "bold": True, "bg_color": "#7B3F00",
                                   "font_color": "#FFFFFF", "align": "center", "text_wrap": True})

    f["title"]   = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 13,
                                   "bg_color": "#1F3864", "font_color": "#FFFFFF",
                                   "align": "center", "valign": "vcenter"})
    f["title_av"]= wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 13,
                                   "bg_color": "#2E4057", "font_color": "#FFFFFF",
                                   "align": "center", "valign": "vcenter"})

    f["data"]    = wb.add_format({**base, "align": "left"})
    f["center"]  = wb.add_format({**base, "align": "center"})
    f["num"]     = wb.add_format({**base, "align": "center", "num_format": "#,##0"})
    f["bold_c"]  = wb.add_format({**base, "bold": True, "align": "center"})

    for st in STATUS_BG:
        f[f"st_{st}"] = wb.add_format({**base, "bold": True,
            "bg_color": STATUS_BG[st], "font_color": STATUS_FG[st], "align": "center"})
    for sc in SIT_BG:
        f[f"sc_{sc}"] = wb.add_format({**base, "bold": True,
            "bg_color": SIT_BG[sc], "font_color": SIT_FG[sc], "align": "center"})

    f["an_hdr"]  = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 10,
        "bg_color": "#2F5496", "font_color": "#FFFFFF", "align": "center", "border": 1})
    f["an_tot"]  = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 10,
        "bg_color": "#D9E1F2", "align": "center", "border": 1})
    f["an_dat"]  = wb.add_format({"font_name": "Calibri", "font_size": 10,
        "align": "center", "border": 1})
    f["an_ts"]   = wb.add_format({"italic": True, "font_name": "Calibri", "font_size": 10,
        "font_color": "#FFFFFF", "bg_color": "#1F3864", "align": "right", "valign": "vcenter"})

    f["rs_title"]= wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 13,
        "bg_color": "#1F3864", "font_color": "#FFFFFF", "align": "center", "valign": "vcenter"})
    f["rs_lbl"]  = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 11})
    f["rs_val"]  = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 11,
        "align": "center", "border": 1})
    f["rs_hdr"]  = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 10,
        "bg_color": "#2F5496", "font_color": "#FFFFFF", "align": "center", "border": 1})
    f["rs_ca"]   = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 10,
        "bg_color": "#4472C4", "font_color": "#FFFFFF", "border": 1})
    f["rs_np"]   = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 10,
        "bg_color": "#FFFF00", "font_color": "#000000", "border": 1})
    f["rs_num"]  = wb.add_format({"font_name": "Calibri", "font_size": 10,
        "align": "center", "border": 1})
    for st in STATUS_BG:
        f[f"rs_{st}"] = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 10,
            "bg_color": STATUS_BG[st], "font_color": STATUS_FG[st], "border": 1})

    return f

# ─────────────────────── ABA GERAL ────────────────────────────────────────
def write_geral(wb, ws, recs, fmt, suffix=""):
    row = 0
    if suffix:
        ws.merge_range(0, 0, 0, len(HEADERS_GERAL)-1,
                       f"ANÁLISE DE AVALIAÇÕES — {suffix}", fmt["title"])
        ws.set_row(0, 22); row = 1
    for c, h in enumerate(HEADERS_GERAL):
        ws.write(row, c, h, fmt["hdr"])
    ws.set_row(row, 32)
    ws.freeze_panes(row+1, 0)
    row += 1
    for r in recs:
        data = rec_to_row(r)
        for c, val in enumerate(data):
            if c == COL_STATUS:
                ws.write(row, c, val, fmt.get(f"st_{val}", fmt["center"]))
            elif c == COL_SIT:
                ws.write(row, c, val, fmt.get(f"sc_{val}", fmt["center"]))
            elif c == COL_CERT:
                ws.write(row, c, val, fmt["center"])
            else:
                ws.write(row, c, val, fmt["data"])
        row += 1
    for c, w in enumerate(COL_WIDTHS):
        if c < len(HEADERS_GERAL):
            ws.set_column(c, c, w)

# ─────────────────────── ABA AVALIAÇÕES PENDENTES ─────────────────────────
STATUS_PEND = {"Homologação", "Parcialmente Encerrada", "Aberta"}

def write_av_pend(wb, ws, recs, fmt, suffix=""):
    pend = [r for r in recs if r["status"] in STATUS_PEND]
    write_geral(wb, ws, pend, fmt,
                suffix="AVALIAÇÕES PENDENTES" + (f" — {suffix}" if suffix else ""))
    return pend

# ─────────────────────── ABA AVALIADORES PENDENTES ────────────────────────
#
# Estrutura do modelo: tabela única consolidada (AV1 + AV2 combinados)
# Colunas (14 no modelo):
#   Identificação (5 cols) | AV1 Em Aberto (CA, NP, Sub) | AV2 (CA_ab, CA_pe, NP_ab, NP_pe, Sub) | TOTAL
# + Seção separada para HOMOLOGADOR (avaliações em "Homologação")
#

HDRS_AVPEND = [
    "Nº PM", "Nome Completo", "Posto/Graduação", "Unidade RPM", "Unidade Principal",
    # AV1
    "CA\nEm Aberto", "NP\nEm Aberto", "Subtotal\nAv1",
    # AV2
    "CA\nEm Aberto", "CA\nParc.Enc.", "NP\nEm Aberto", "NP\nParc.Enc.", "Subtotal\nAv2",
    "TOTAL",
]
HDRS_HOM = [
    "Nº PM (Avaliado)", "Nome (Avaliado)", "Posto/Graduação", "Unidade RPM",
    "Unidade Principal", "Local/Unidade",
    "Conceito AV1 (J)", "Nota AV2 (L)",
    "Situação Comissão",
]

def build_avpend(recs):
    data = {}
    def get(nrpm, nome, posto, rpm, unid):
        if nrpm not in data:
            data[nrpm] = {"nome": nome, "posto": posto, "rpm": rpm, "unid": unid,
                          "av1_ca": 0, "av1_np": 0,
                          "av2_ca_ab": 0, "av2_ca_pe": 0,
                          "av2_np_ab": 0, "av2_np_pe": 0}
        return data[nrpm]
    for r in recs:
        st = r["status"]; sc = r["sit_comissao"]
        is_ca = sc == "Comissão Atual"
        # AV1: somente "Aberta"
        if st == "Aberta" and r["nrpm_av1"]:
            d = get(r["nrpm_av1"], r["nome_av1"], r["posto_av1"], r["rpm_av1"], r["unid_princ_av1"])
            if is_ca: d["av1_ca"] += 1
            else:     d["av1_np"] += 1
        # AV2: "Aberta" ou "Parcialmente Encerrada"
        if st in ("Aberta", "Parcialmente Encerrada") and r["nrpm_av2"]:
            d = get(r["nrpm_av2"], r["nome_av2"], r["posto_av2"], r["rpm_av2"], r["unid_princ_av2"])
            if st == "Aberta":
                if is_ca: d["av2_ca_ab"] += 1
                else:     d["av2_np_ab"] += 1
            else:
                if is_ca: d["av2_ca_pe"] += 1
                else:     d["av2_np_pe"] += 1
    return data

def write_avaliadores_pend(wb, ws, recs, fmt, suffix=""):
    avpend = build_avpend(recs)
    hom_list = [r for r in recs if r["status"] == "Homologação"]
    row = 0

    # ── Título consolidado ────────────────────────────────────────────────
    title = ("VISÃO CONSOLIDADA — AVALIADORES PENDENTES (Funções Av1 e Av2 combinadas)"
             + (f" | {suffix}" if suffix else ""))
    ws.merge_range(row, 0, row, len(HDRS_AVPEND)-1, title, fmt["title_av"])
    ws.set_row(row, 22); row += 1

    # Sub-cabeçalho de grupos
    ws.merge_range(row, 0, row, 4, "IDENTIFICAÇÃO", fmt["hdr_av"])
    ws.merge_range(row, 5, row, 7, "AVALIADOR 1 — Em Aberto", fmt["hdr_av"])
    ws.merge_range(row, 8, row, 12, "AVALIADOR 2 — Pendências (Em Aberto + Parc. Encerrada)", fmt["hdr_av"])
    ws.write(row, 13, "TOTAL", fmt["hdr_av"])
    ws.set_row(row, 22); row += 1

    # Cabeçalhos
    for c, h in enumerate(HDRS_AVPEND):
        ws.write(row, c, h, fmt["hdr_sub"])
    ws.set_row(row, 36)
    ws.freeze_panes(row+1, 0)
    row += 1

    # Dados ordenados por TOTAL desc
    items = sorted(avpend.items(), key=lambda x: -(
        x[1]["av1_ca"] + x[1]["av1_np"] +
        x[1]["av2_ca_ab"] + x[1]["av2_ca_pe"] +
        x[1]["av2_np_ab"] + x[1]["av2_np_pe"]
    ))
    for nrpm, d in items:
        s1 = d["av1_ca"] + d["av1_np"]
        s2 = d["av2_ca_ab"] + d["av2_ca_pe"] + d["av2_np_ab"] + d["av2_np_pe"]
        tot = s1 + s2
        if tot == 0: continue
        rdata = [nrpm, d["nome"], d["posto"], d["rpm"], d["unid"],
                 d["av1_ca"], d["av1_np"], s1,
                 d["av2_ca_ab"], d["av2_ca_pe"], d["av2_np_ab"], d["av2_np_pe"], s2,
                 tot]
        for c, v in enumerate(rdata):
            ws.write(row, c, v, fmt["data"] if c < 5 else fmt["num"])
        row += 1

    if not items:
        ws.merge_range(row, 0, row, len(HDRS_AVPEND)-1,
                       "Nenhum avaliador com pendências.", fmt["center"]); row += 1

    # ── SEÇÃO: HOMOLOGADOR ────────────────────────────────────────────────
    row += 2
    title_hom = ("HOMOLOGADOR — Avaliações com Divergência Aguardando Nota de Homologação (Col N)"
                 + (f" | {suffix}" if suffix else ""))
    ws.merge_range(row, 0, row, len(HDRS_HOM)-1, title_hom, fmt["hdr_hom"])
    ws.set_row(row, 22); row += 1

    info = ("Conceito AV1 (J) e Nota AV2 (L) divergem conforme Resolução 5458/2025. "
            "O Homologador não está identificado no CSV — listadas as avaliações pendentes de homologação.")
    ws.merge_range(row, 0, row, len(HDRS_HOM)-1, info, fmt["data"])
    ws.set_row(row, 28); row += 1

    for c, h in enumerate(HDRS_HOM):
        ws.write(row, c, h, fmt["hdr_hom"])
    ws.set_row(row, 28); row += 1

    if hom_list:
        for r in sorted(hom_list, key=lambda x: (x["rpm_av"], x["nome_av"])):
            rdata = [r["nrPM_av"], r["nome_av"], r["posto_av"], r["rpm_av"],
                     r["unid_princ_av"], r["local_av"],
                     r["conceito"], r["nota_geral"], r["sit_comissao"]]
            for c, v in enumerate(rdata):
                if c == 8:
                    ws.write(row, c, v, fmt.get(f"sc_{v}", fmt["center"]))
                else:
                    ws.write(row, c, v, fmt["data"])
            row += 1
    else:
        ws.merge_range(row, 0, row, len(HDRS_HOM)-1,
                       "Nenhuma avaliação em situação de Homologação pendente.", fmt["center"])

    # Larguras
    ws.set_column(0, 0, 12);  ws.set_column(1, 1, 38)
    ws.set_column(2, 2, 14);  ws.set_column(3, 3, 16)
    ws.set_column(4, 4, 28);  ws.set_column(5, 13, 14)

# ─────────────────────── ABA ANÁLISE ──────────────────────────────────────
STATUSES = ["Aberta", "Parcialmente Encerrada", "Homologação", "Encerrada"]

def write_analise(wb, ws, recs, fmt, suffix=""):
    counts = defaultdict(lambda: defaultdict(int))
    for r in recs:
        counts[r["status"]][r["sit_comissao"]] += 1

    titulo   = "TABELA DINÂMICA — ANÁLISE DE AVALIAÇÕES" + (f" — {suffix}" if suffix else "")
    data_hora = datetime.now().strftime("Dados consolidados em %d/%m/%Y às %Hh")

    ws.merge_range(0, 0, 0, 6, titulo,    fmt["title"])
    ws.merge_range(0, 7, 0, 11, data_hora, fmt["an_ts"])
    ws.set_row(0, 22)

    ws.write(1, 0, "Status Avaliação", fmt["an_hdr"])
    ws.write(1, 1, "Comissão Atual",   fmt["an_hdr"])
    ws.write(1, 2, "Nota Provisória",  fmt["an_hdr"])
    ws.write(1, 3, "Total",            fmt["an_hdr"])
    ws.write(1, 4, "Status",           fmt["an_hdr"])
    ws.write(1, 5, "Qtd",              fmt["an_hdr"])

    tots = {"CA": 0, "NP": 0, "T": 0}
    for i, st in enumerate(STATUSES):
        r = 2 + i
        ca  = counts[st]["Comissão Atual"]
        np_ = counts[st]["Nota Provisória"]
        tot = ca + np_
        tots["CA"] += ca; tots["NP"] += np_; tots["T"] += tot
        sf = fmt.get(f"st_{st}", fmt["an_hdr"])
        ws.write(r, 0, st, sf);  ws.write(r, 1, ca,  fmt["an_dat"])
        ws.write(r, 2, np_, fmt["an_dat"]); ws.write(r, 3, tot, fmt["an_dat"])
        ws.write(r, 4, st, sf);  ws.write(r, 5, tot, fmt["an_dat"])

    tr = 2 + len(STATUSES)
    ws.write(tr, 0, "TOTAL",     fmt["an_tot"])
    ws.write(tr, 1, tots["CA"],  fmt["an_tot"])
    ws.write(tr, 2, tots["NP"],  fmt["an_tot"])
    ws.write(tr, 3, tots["T"],   fmt["an_tot"])

    ws.set_column(0, 0, 26); ws.set_column(1, 3, 18)
    ws.set_column(4, 4, 26); ws.set_column(5, 5, 12)

    # Gráfico de pizza
    chart = wb.add_chart({"type": "pie"})
    n = len(STATUSES)
    chart.add_series({
        "name":       "Status",
        "categories": ["Análise", 2, 4, 2+n-1, 4],
        "values":     ["Análise", 2, 5, 2+n-1, 5],
        "data_labels": {"percentage": True, "category": True, "separator": "\n",
                        "font": {"size": 9}},
        "points": [
            {"fill": {"color": STATUS_BG["Aberta"]}},
            {"fill": {"color": STATUS_BG["Parcialmente Encerrada"]}},
            {"fill": {"color": STATUS_BG["Homologação"]}},
            {"fill": {"color": STATUS_BG["Encerrada"]}},
        ],
    })
    chart.set_title({"name": f"Status — {suffix}" if suffix else "Status das Avaliações"})
    chart.set_style(10); chart.set_size({"width": 480, "height": 320})
    ws.insert_chart("H2", chart)

# ─────────────────────── ABA RESUMO ───────────────────────────────────────
def write_resumo(wb, ws, recs, fmt, suffix=""):
    total = len(recs)
    ca_t = sum(1 for r in recs if r["sit_comissao"] == "Comissão Atual")
    np_t = sum(1 for r in recs if r["sit_comissao"] == "Nota Provisória")

    def cnt(st, sc=None):
        return sum(1 for r in recs if r["status"]==st and (sc is None or r["sit_comissao"]==sc))

    titulo = "RESUMO — ANÁLISE DE AVALIAÇÕES" + (f" — {suffix}" if suffix else " (AADP 2026)")
    ws.merge_range(0, 0, 0, 3, titulo, fmt["rs_title"])
    ws.set_row(0, 22)

    ws.write(1, 0, "TOTAL DE REGISTROS", fmt["rs_lbl"])
    ws.write(1, 2, total, fmt["rs_val"])

    ws.write(3, 0, "SITUAÇÃO DA COMISSÃO", fmt["rs_hdr"])
    ws.write(3, 1, "Total", fmt["rs_hdr"])
    ws.write(4, 0, "🟢 Comissão Atual",  fmt["rs_ca"])
    ws.write(4, 1, ca_t,                  fmt["rs_num"])
    ws.write(5, 0, "🔴 Nota Provisória", fmt["rs_np"])
    ws.write(5, 1, np_t,                  fmt["rs_num"])

    ws.write(7, 0, "STATUS DA AVALIAÇÃO",  fmt["rs_hdr"])
    ws.write(7, 1, "Total",                fmt["rs_hdr"])
    ws.write(7, 2, "Comissão Atual",       fmt["rs_hdr"])
    ws.write(7, 3, "Nota Provisória",      fmt["rs_hdr"])

    for i, st in enumerate(["Encerrada","Homologação","Parcialmente Encerrada","Aberta"]):
        r = 8 + i
        ws.write(r, 0, st,                    fmt.get(f"rs_{st}", fmt["rs_lbl"]))
        ws.write(r, 1, cnt(st),               fmt["rs_num"])
        ws.write(r, 2, cnt(st,"Comissão Atual"), fmt["rs_num"])
        ws.write(r, 3, cnt(st,"Nota Provisória"),fmt["rs_num"])

    ws.set_column(0, 0, 28); ws.set_column(1, 1, 12)
    ws.set_column(2, 2, 18); ws.set_column(3, 3, 18)

# ─────────────────────── CRIAÇÃO DO WORKBOOK ──────────────────────────────
def create_workbook(recs: list, filepath: str, suffix: str = ""):
    wb  = xlsxwriter.Workbook(filepath, {"constant_memory": False})
    fmt = make_formats(wb)

    write_geral(wb,            wb.add_worksheet("Geral"),                  recs, fmt, suffix)
    write_av_pend(wb,          wb.add_worksheet("Avaliações Pendentes"),   recs, fmt, suffix)
    write_avaliadores_pend(wb, wb.add_worksheet("Avaliadores Pendentes"),  recs, fmt, suffix)
    write_analise(wb,          wb.add_worksheet("Análise"),                recs, fmt, suffix)
    write_resumo(wb,           wb.add_worksheet("Resumo"),                 recs, fmt, suffix)

    wb.close()

# ─────────────────────── MAIN (com argparse) ──────────────────────────────
import argparse

parser = argparse.ArgumentParser(description="Gera relatórios AADP 2026 em Excel/ZIP.")
parser.add_argument("--mode",   choices=["all","geral","units"], default="all",
                    help="all=tudo, geral=somente Geral, units=unidades específicas")
parser.add_argument("--units",  default="",
                    help="Lista de RPMs separadas por vírgula (usado com --mode units)")
parser.add_argument("--output", default=None,
                    help="Caminho do arquivo ZIP de saída (padrão: Resultado_AADP_2026.zip)")
args = parser.parse_args()

zip_out = args.output or os.path.join(BASE_DIR, "Resultado_AADP_2026.zip")
units_req = [u.strip() for u in args.units.split(",") if u.strip()] if args.units else []

# Montar mapa RPM → lista de registros
rpm_map: dict[str, list] = defaultdict(list)
for r in records:
    rpm_map[r["rpm_av"].strip()].append(r)

arquivos_zip = []

# ── Geral ─────────────────────────────────────────────────────────────────
if args.mode in ("all", "geral"):
    print("\n=== Gerando Planilha GERAL ===")
    geral_path = os.path.join(OUT_DIR, "Analise_Avaliacoes_Geral.xlsx")
    create_workbook(records, geral_path, suffix="AADP 2026")
    print(f"  OK → {geral_path}")
    arquivos_zip.append(geral_path)

# ── Por RPM ───────────────────────────────────────────────────────────────
if args.mode in ("all", "units"):
    target_rpms = units_req if units_req else sorted(rpm_map.keys())
    print(f"\n=== Gerando Planilhas por Unidade RPM ({len(target_rpms)} unidades) ===")
    for rpm in target_rpms:
        recs_rpm = rpm_map.get(rpm, [])
        if not recs_rpm:
            print(f"  AVISO: nenhum registro para '{rpm}'")
            continue
        safe  = rpm.replace("/","_").replace("\\","_").replace(" ","_").replace(":","_")
        fname = f"Analise_Avaliacoes_{safe}.xlsx"
        fp    = os.path.join(OUT_DIR, fname)
        print(f"  {fname}  ({len(recs_rpm):,} reg.)")
        create_workbook(recs_rpm, fp, suffix=rpm)
        arquivos_zip.append(fp)
    print(f"\n  {len(arquivos_zip) - (1 if args.mode=='all' else 0)} planilhas por RPM geradas.")

# ── Compactar ZIP ─────────────────────────────────────────────────────────
print("\n=== Compactando ZIP ===")
with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as zf:
    for fp in arquivos_zip:
        zf.write(fp, arcname=os.path.basename(fp))

print(f"  ZIP → {zip_out}")
print(f"\n=== CONCLUÍDO ===")
print(f"  Total registros : {len(records):,}")
print(f"  Aberta          : {sc['Aberta']:,}")
print(f"  Parc. Encerrada : {sc['Parcialmente Encerrada']:,}")
print(f"  Homologação     : {sc['Homologação']:,}")
print(f"  Encerrada       : {sc['Encerrada']:,}")
print(f"  Planilhas       : {len(arquivos_zip)}")
print(f"  Arquivo final   : {zip_out}")

