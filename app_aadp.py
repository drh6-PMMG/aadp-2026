"""
AADP 2026 — Dashboard de Análise de Avaliações
Versão 3.0 — Google Drive + Streamlit Cloud + Geração em memória
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io, os, re, json, subprocess, unicodedata, csv, tempfile, zipfile
from pathlib import Path
from collections import defaultdict
from datetime import datetime

# gdown: download do Google Drive (opcional — só necessário no modo Drive)
try:
    import gdown
    GDOWN_OK = True
except ImportError:
    GDOWN_OK = False

pd.set_option("styler.render.max_elements", 5_000_000)

st.set_page_config(
    page_title="AADP 2026 — Análise de Avaliações",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────── CSS ──────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}

[data-testid="stSidebar"]{
  background:linear-gradient(180deg,#000000 0%,#1a1a1a 100%);
  border-right:1px solid #bca374;
}
[data-testid="stSidebar"] *{color:#e5dccb!important;}
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3{color:#bca374!important;}
.main{background:#f5f3ef;}

.main-title{
  background:linear-gradient(135deg,#000000 0%,#1f1f1f 100%);
  color:#bca374;padding:24px 28px;border-radius:12px;margin-bottom:20px;
  display:flex;flex-direction:column;align-items:center;text-align:center;gap:16px;
  box-shadow:0 4px 20px rgba(0,0,0,.15);
  border-bottom:3px solid #bca374;
}
.main-title h1{margin:0;font-size:2.1rem;font-weight:800;color:#bca374;letter-spacing:0.02em;}
.main-title p{margin:0;font-size:.9rem;opacity:.9;color:#e5dccb;margin-top:6px;}

.kpi-card{background:#fff;border-radius:12px;padding:16px 20px;
  box-shadow:0 2px 12px rgba(0,0,0,.08);border-left:5px solid;
  transition:transform .2s,box-shadow .2s;text-align:center;}
.kpi-card:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,0,0,.12);}
.kpi-card .label{font-size:.7rem;font-weight:700;text-transform:uppercase;
  letter-spacing:.06em;color:#555555!important;margin-bottom:6px;}
.kpi-card .value{font-size:1.9rem;font-weight:800;line-height:1;}
.kpi-card .sub{font-size:.72rem;color:#777777!important;margin-top:4px;}
.kpi-total    {border-color:#bca374;} .kpi-total    .value{color:#8c6e42!important;}
.kpi-ca       {border-color:#bca374;} .kpi-ca       .value{color:#8c6e42!important;}
.kpi-np       {border-color:#8c6e42;} .kpi-np       .value{color:#8c6e42!important;}
.kpi-aberta   {border-color:#FF4444;} .kpi-aberta   .value{color:#cc2222!important;}
.kpi-parc     {border-color:#FF8C00;} .kpi-parc     .value{color:#cc7000!important;}
.kpi-hom      {border-color:#FFD966;} .kpi-hom      .value{color:#b89500!important;}
.kpi-enc      {border-color:#70AD47;} .kpi-enc      .value{color:#4a8a28!important;}

.stTabs [data-baseweb="tab-list"]{background:#fff;border-radius:10px;padding:6px;
  box-shadow:0 2px 8px rgba(0,0,0,.07);gap:4px;}
.stTabs [data-baseweb="tab"]{border-radius:8px;font-weight:600;padding:8px 20px;font-size:.85rem;}
.stTabs [aria-selected="true"]{background:#bca374!important;color:#000!important;}

.section-hdr{background:#bca374;color:#000;padding:10px 16px;border-radius:8px;
  font-weight:600;font-size:.9rem;margin:16px 0 8px 0;}
.section-hdr-hom{background:#8c6e42;color:#fff;padding:10px 16px;border-radius:8px;
  font-weight:600;font-size:.9rem;margin:24px 0 8px 0;}
.info-box{background:#f9f6f0;border:1px solid #e5dccb;border-radius:8px;
  padding:12px 16px;font-size:.85rem;color:#8c6e42;margin-bottom:12px;}
.warn-box{background:#fff8e1;border:1px solid #ffe082;border-radius:8px;
  padding:12px 16px;font-size:.85rem;color:#7a5c00;margin-bottom:12px;}
div[data-testid="metric-container"]{background:#fff;border-radius:10px;
  padding:12px;box-shadow:0 2px 8px rgba(0,0,0,.07);}
.stButton button{background:linear-gradient(135deg,#000000,#242424);
  color:#bca374;border:1px solid #bca374;font-weight:600;border-radius:8px;
  transition:all .2s;box-shadow:0 2px 8px rgba(188,163,116,.2);}
.stButton button:hover{transform:translateY(-1px);box-shadow:0 4px 16px rgba(188,163,116,.4);
  background:#bca374;color:#000;}

/* Botões do Sidebar - Efeito Cristal Líquido Secundário (Inativo) */
div[data-testid="stSidebar"] button[data-testid="baseButton-secondary"] {
  background: linear-gradient(135deg, #242424 0%, #000000 100%) !important;
  color: #bca374 !important;
  border: 1px solid #8c6e42 !important;
  border-radius: 8px !important;
  box-shadow: inset 0 1px 2px rgba(255,255,255,0.1), 0 2px 4px rgba(0,0,0,0.5) !important;
  font-weight: 600 !important;
  transition: all 0.3s ease !important;
}
div[data-testid="stSidebar"] button[data-testid="baseButton-secondary"]:hover {
  background: linear-gradient(135deg, #bca374 0%, #1a1a1a 100%) !important;
  color: #ffffff !important;
  border: 1px solid #bca374 !important;
  box-shadow: inset 0 1px 2px rgba(255,255,255,0.2), 0 4px 12px rgba(188,163,116,0.4) !important;
  transform: translateY(-1px) !important;
}

/* Botões do Sidebar - Efeito Cristal Líquido Primário (Ativo - Caqui para Preto) */
div[data-testid="stSidebar"] button[data-testid="baseButton-primary"] {
  background: linear-gradient(135deg, #bca374 0%, #000000 100%) !important;
  color: #ffffff !important;
  border: 1px solid #ffffff !important;
  border-radius: 8px !important;
  box-shadow: inset 0 1px 3px rgba(255,255,255,0.3), 0 4px 15px rgba(188,163,116,0.5) !important;
  font-weight: 700 !important;
  transition: all 0.3s ease !important;
  text-shadow: 0 1px 2px rgba(0,0,0,0.8) !important;
}
div[data-testid="stSidebar"] button[data-testid="baseButton-primary"]:hover {
  background: linear-gradient(135deg, #ffffff 0%, #bca374 100%) !important;
  color: #000000 !important;
  border: 1px solid #ffffff !important;
  box-shadow: inset 0 1px 3px rgba(255,255,255,0.4), 0 6px 20px rgba(188,163,116,0.7) !important;
  transform: translateY(-2px) !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────── CONSTANTES ───────────────────────────────────────────
THIS_DIR  = Path(__file__).parent
DADOS_DIR = THIS_DIR / "dados"
DADOS_DIR.mkdir(exist_ok=True)
CONFIG_FILE = THIS_DIR / "config_aadp.json"

SITUACOES_ALVO = {
    "ATIV. DIRECAO GERAL","ATIV. FIM DESTACADO","ATIV. FIM NA SEDE",
    "ATIV. MEIO","ATIVIDADE MEIO","DISP MED DEFINITIVA",
    "GESTANTE/LAC/ADOTANT","QUADRO ESPECIALISTA",
}
CONCEITO_FAIXA = {
    "nivel superior de desempenho":       (9.00, 10.00),
    "nivel alto de desempenho":           (7.00,  8.99),
    "nivel intermediario de desempenho":  (6.00,  6.99),
    "nivel baixo de desempenho":          (3.00,  5.99),
    "nivel inferior de desempenho":       (0.00,  2.99),
}
STATUS_COLORS = {
    "Encerrada":             "#70AD47",
    "Homologação":           "#FFD966",
    "Parcialmente Encerrada":"#FF8C00",
    "Aberta":                "#FF4444",
}
SIT_COLORS = {"Comissão Atual":"#4472C4","Nota Provisória":"#FFC000"}
# Ordem para empilhamento: Encerradas embaixo, pendentes em cima
STACK_ORDER  = ["Encerrada","Aberta","Parcialmente Encerrada","Homologação"]

# ─────────────────────── CONFIGURAÇÃO ─────────────────────────────────────────
def load_config():
    if CONFIG_FILE.exists():
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    return {"db_path": str(DADOS_DIR)}

def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

cfg = load_config()

# ─────────────────────── LÓGICA DADOS ─────────────────────────────────────────
def normaliza(t):
    s = unicodedata.normalize("NFD", t.lower())
    return "".join(c for c in s if unicodedata.category(c) != "Mn")

def is_empty(v):
    return not v or str(v).strip() in ("", "-")

def concordam(j, l):
    if is_empty(j) or is_empty(l): return None
    try: nota = float(str(l).replace(",","."))
    except: return None
    faixa = CONCEITO_FAIXA.get(normaliza(j.strip()))
    if faixa is None: return None
    return faixa[0] <= nota <= faixa[1]

def calc_cert(j, l):
    if is_empty(j) or is_empty(l): return "-"
    c = concordam(j, l)
    return "NÃO" if c is True else ("SIM" if c is False else "-")

def calc_status(j, l, n):
    if is_empty(j): return "Aberta"
    if is_empty(l): return "Parcialmente Encerrada"
    c = concordam(j, l)
    if c is True: return "Encerrada"
    if c is False: return "Encerrada" if not is_empty(n) else "Homologação"
    return "Encerrada"

def rpm_sort_key(name):
    m = re.match(r'^(\d+)\s+RPM', str(name))
    if m: return (0, int(m.group(1)), "")
    return (1, 0, str(name))

def _baixar_drive(file_id: str, destino: str):
    """Baixa um arquivo do Google Drive para destino local.
    Compatível com todas as versões do gdown (com e sem parâmetro fuzzy).
    """
    if not GDOWN_OK:
        raise ImportError("Biblioteca 'gdown' não instalada. Execute: pip install gdown")
    import inspect
    url = f"https://drive.google.com/uc?id={file_id}&export=download"
    # gdown >= 4.6 suporta fuzzy; versões mais antigas não suportam
    sig = inspect.signature(gdown.download)
    if "fuzzy" in sig.parameters:
        gdown.download(url, destino, quiet=True, fuzzy=True)
    else:
        gdown.download(url, destino, quiet=True)
    if not os.path.exists(destino) or os.path.getsize(destino) == 0:
        raise FileNotFoundError(f"Falha ao baixar arquivo do Drive (ID: {file_id})")


def _parse_csv(av_f: str, si_f: str) -> pd.DataFrame:
    """Processa os dois CSVs e retorna o DataFrame final."""
    sigef = {}
    with open(si_f, encoding="cp1252", errors="replace") as f:
        for row in csv.reader(f, delimiter=";"):
            if len(row) > 9:
                sigef[row[0].strip().lstrip("0") or "0"] = row[9].strip()
    rows = []
    with open(av_f, encoding="cp1252", errors="replace") as f:
        reader = csv.reader(f, delimiter=";")
        next(reader)
        for row in reader:
            while len(row) < 50: row.append("")  # CSV tem 50 colunas (colégio até Homologador)
            sit = row[7].strip()
            if sit not in SITUACOES_ALVO: continue
            nrpm = row[0].strip(); local = row[5].strip()
            j = row[9].strip(); l = row[11].strip(); n = row[13].strip()
            sc = "Comissão Atual" if local.upper().strip() == sigef.get(nrpm.lstrip("0") or "0","").upper().strip() else "Nota Provisória"
            rows.append({
                "nrPM (Avaliado)":          nrpm,
                "Nome (Avaliado)":           row[1].strip(),
                "Posto/Grad. (Avaliado)":    row[2].strip(),
                "Unidade RPM (Avaliado)":    row[3].strip(),
                "Unidade Principal (Avaliado)": row[4].strip(),
                "Local/Unidade (Avaliado)":  local,
                "Quadro Atual (Avaliado)":   row[6].strip(),
                "Situação Funcional":        sit,
                "Data AV1":                  row[8].strip(),
                "Conceito Geral":            j,
                "Data AV2":                  row[10].strip(),
                "Nota Geral":                l,
                "Certificação Homologador":  calc_cert(j, l),
                "Data HOM":                  row[12].strip(),
                "Nota Homologação":          n,
                "Competência 1":             row[14].strip(),
                "Conceito Comp.1":           row[15].strip(), "Nota Comp.1": row[16].strip(),
                "Competência 2":             row[17].strip(),
                "Conceito Comp.2":           row[18].strip(), "Nota Comp.2": row[19].strip(),
                "Competência 3":             row[20].strip(),
                "Conceito Comp.3":           row[21].strip(), "Nota Comp.3": row[22].strip(),
                "Competência 4":             row[23].strip(),
                "Conceito Comp.4":           row[24].strip(), "Nota Comp.4": row[25].strip(),
                # Avaliador 1
                "nrPM (Av1)":    row[26].strip(), "Nome (Av1)":  row[27].strip(),
                "Posto (Av1)":   row[28].strip(), "RPM (Av1)":   row[29].strip(),
                "Unid. Principal (Av1)": row[30].strip(), "Local (Av1)": row[31].strip(),
                "Quadro (Av1)":  row[32].strip(), "Situação (Av1)": row[33].strip(),
                # Avaliador 2
                "nrPM (Av2)":    row[34].strip(), "Nome (Av2)":  row[35].strip(),
                "Posto (Av2)":   row[36].strip(), "RPM (Av2)":   row[37].strip(),
                "Unid. Principal (Av2)": row[38].strip(), "Local (Av2)": row[39].strip(),
                "Quadro (Av2)":  row[40].strip(), "Situação (Av2)": row[41].strip(),
                # Homologador (colunas 42–49)
                "nrPM (Hom)":    row[42].strip(), "Nome (Hom)":  row[43].strip(),
                "Posto (Hom)":   row[44].strip(), "RPM (Hom)":   row[45].strip(),
                "Unid. Principal (Hom)": row[46].strip(), "Local (Hom)": row[47].strip(),
                "Quadro (Hom)":  row[48].strip(), "Situação (Hom)": row[49].strip(),
                "Situação Comissão": sc,
                "Status Avaliação":  calc_status(j, l, n),
            })
    return pd.DataFrame(rows)

@st.cache_data(show_spinner="⏳ Carregando e processando dados...")
def load_data(db_path: str, drive_av_id: str = "", drive_si_id: str = ""):
    """Carrega dados de pasta local ou Google Drive."""
    if drive_av_id and drive_si_id:
        # ── Modo Google Drive ──────────────────────────────────────────────
        tmp = tempfile.mkdtemp(prefix="aadp_")
        av_f = os.path.join(tmp, "avaliacoes.csv")
        si_f = os.path.join(tmp, "SIGEF.csv")
        _baixar_drive(drive_av_id, av_f)
        _baixar_drive(drive_si_id, si_f)
    else:
        # ── Modo pasta local ───────────────────────────────────────────────
        av_f = os.path.join(db_path, "avaliacoes.csv")
        si_f = os.path.join(db_path, "SIGEF.csv")
        if not os.path.exists(av_f): raise FileNotFoundError(f"Não encontrado: {av_f}")
        if not os.path.exists(si_f): raise FileNotFoundError(f"Não encontrado: {si_f}")
    return _parse_csv(av_f, si_f)

def apply_filters(df, rpm_f, unid_f, sc_f, st_f, cert_f):
    if rpm_f:  df = df[df["Unidade RPM (Avaliado)"].isin(rpm_f)]
    if unid_f: df = df[df["Unidade Principal (Avaliado)"].isin(unid_f)]
    if sc_f:   df = df[df["Situação Comissão"].isin(sc_f)]
    if st_f:   df = df[df["Status Avaliação"].isin(st_f)]
    if cert_f: df = df[df["Certificação Homologador"].isin(cert_f)]
    return df

def fmt_num(n): return f"{n:,}".replace(",",".")

MAX_STYLE = 4_000_000
def safe_df(styled_or_df, height=520):
    if hasattr(styled_or_df, "data"):
        if styled_or_df.data.size <= MAX_STYLE:
            st.dataframe(styled_or_df, use_container_width=True, height=height)
        else:
            st.info("ℹ️ Tabela grande — exibida sem cores para melhor desempenho.")
            st.dataframe(styled_or_df.data, use_container_width=True, height=height)
    else:
        st.dataframe(styled_or_df, use_container_width=True, height=height)

def color_status(val):
    m = {"Encerrada":"background-color:#e8f5e1;color:#2d6a0f;font-weight:600",
         "Homologação":"background-color:#fff8db;color:#7a5c00;font-weight:600",
         "Parcialmente Encerrada":"background-color:#fff0db;color:#7a3d00;font-weight:600",
         "Aberta":"background-color:#fde8e8;color:#8b0000;font-weight:600"}
    return m.get(val, "")

def color_sit(val):
    if val == "Comissão Atual":   return "background-color:#dce8f5;color:#1a3a6a;font-weight:600"
    if val == "Nota Provisória":  return "background-color:#fff9e6;color:#7a5c00;font-weight:600"
    return ""

# ─────────────────────── SIDEBAR ──────────────────────────────────────────────
with st.sidebar:
    st.image("logo_drh.png", use_container_width=True)
    st.markdown("### AADP 2026")
    st.markdown("**Sistema de Análise de Avaliações**")
    st.markdown("---")

    # 🧭 BOTÕES DE NAVEGAÇÃO DA PÁGINA CENTRAL
    st.markdown("#### 🧭 Páginas")
    pages = [
        ("📊 Análise Gráfica", "Análise Gráfica"),
        ("📋 Dados Gerais", "Dados Gerais"),
        ("⏳ Avaliações Pendentes", "Avaliações Pendentes"),
        ("👥 Avaliadores Pendentes", "Avaliadores Pendentes"),
        ("📥 Gerar Relatório", "Gerar Relatório"),
        ("📄 Relatório Word", "Relatório Word"),
    ]
    if "active_page" not in st.session_state:
        st.session_state.active_page = "Análise Gráfica"
        
    for label, page_name in pages:
        is_active = st.session_state.active_page == page_name
        btn_type = "primary" if is_active else "secondary"
        if st.button(label, key=f"nav_{page_name}", use_container_width=True, type=btn_type):
            st.session_state.active_page = page_name
            st.rerun()
            
    st.markdown("---")
    
    # ⚙️ SEÇÕES COLAPSÁVEIS DA BARRA LATERAL
    if "show_fonte" not in st.session_state:
        st.session_state.show_fonte = False
    if "show_filtros" not in st.session_state:
        st.session_state.show_filtros = True

    # Botão de Fonte de Dados (não muda a página central, apenas recolhe/expande no sidebar)
    btn_fonte_label = "📂 Ocultar Fonte dos Dados" if st.session_state.show_fonte else "📂 Mostrar Fonte dos Dados"
    btn_fonte_type = "primary" if st.session_state.show_fonte else "secondary"
    if st.button(btn_fonte_label, use_container_width=True, key="btn_toggle_fonte", type=btn_fonte_type):
        st.session_state.show_fonte = not st.session_state.show_fonte
        st.rerun()
        
    # Inicializa variáveis para não dar NameError
    drive_av_id = drive_si_id = drive_geral_id = ""
    db_path = ""
    fonte = cfg.get("fonte_dados", "📁 Pasta local / Servidor")
    reload = False
    
    if st.session_state.show_fonte:
        st.markdown("#### 🗄️ Configurações da Fonte")
        fonte = st.radio("Origem:", ["📁 Pasta local / Servidor", "☁️ Google Drive"],
                         horizontal=True,
                         index=0 if "local" in cfg.get("fonte_dados", "local") else 1,
                         key="fonte_dados_radio",
                         help="Local: use pasta 'dados/'. Drive: informe os IDs dos arquivos.")
        if fonte != cfg.get("fonte_dados"):
            cfg["fonte_dados"] = fonte
            save_config(cfg)
            
        if "Drive" in fonte:
            st.markdown("<small>📌 Compartilhe os arquivos no Drive como <b>Qualquer pessoa com o link</b></small>",
                        unsafe_allow_html=True)
            drive_av_id = st.text_input("🔑 ID do avaliacoes.csv no Drive:",
                                         value=cfg.get("drive_av_id",""),
                                         key="drive_av_id_input",
                                         placeholder="Ex: 1BxiMVs...")
            drive_si_id = st.text_input("🔑 ID do SIGEF.csv no Drive:",
                                         value=cfg.get("drive_si_id",""),
                                         key="drive_si_id_input",
                                         placeholder="Ex: 1BxiMVs...")
            drive_geral_id = st.text_input("🔑 ID do Geral.xlsx no Drive:",
                                         value=cfg.get("drive_geral_id",""),
                                         key="drive_geral_id_input",
                                         placeholder="Ex: 1BxiMVs...")
            st.markdown("<small>ℹ️ O ID está na URL do arquivo compartilhado</small>", unsafe_allow_html=True)
            if st.button("💾 Salvar IDs", use_container_width=True, key="btn_save_ids"):
                cfg["drive_av_id"] = drive_av_id
                cfg["drive_si_id"] = drive_si_id
                cfg["drive_geral_id"] = drive_geral_id
                save_config(cfg); st.success("IDs salvos!")
        else:
            st.markdown(f"<small>📂 Pasta padrão: <code>dados/</code></small>", unsafe_allow_html=True)
            db_path = st.text_input("Caminho da pasta CSV:", value=cfg.get("db_path",""),
                                     key="db_path_input",
                                     placeholder=str(DADOS_DIR))
            if st.button("💾 Salvar Caminho", use_container_width=True, key="btn_save_caminho"):
                cfg["db_path"] = db_path; save_config(cfg); st.success("Caminho salvo!")
                
        reload = st.button("🔄 Recarregar Dados", use_container_width=True, type="primary", key="btn_reload")
        
    st.markdown("---")
    
    # Botão de Filtros de Visualização
    btn_filtros_label = "🔍 Ocultar Filtros" if st.session_state.show_filtros else "🔍 Mostrar Filtros"
    btn_filtros_type = "primary" if st.session_state.show_filtros else "secondary"
    if st.button(btn_filtros_label, use_container_width=True, key="btn_toggle_filtros", type=btn_filtros_type):
        st.session_state.show_filtros = not st.session_state.show_filtros
        st.rerun()

# ─────────────────────── CARREGAR DADOS ───────────────────────────────────────
try:
    if reload: st.cache_data.clear()
    df_full = load_data(
        db_path   = db_path or cfg.get("db_path", str(DADOS_DIR)),
        drive_av_id = drive_av_id or cfg.get("drive_av_id", ""),
        drive_si_id = drive_si_id or cfg.get("drive_si_id", ""),
    )
    data_ok = True; ts = datetime.now().strftime("%d/%m/%Y %H:%M")
except Exception as e:
    data_ok = False; err_msg = str(e)

# ─────────────────────── FILTROS ──────────────────────────────────────────────
rpm_filter = unid_filter = sit_com_filter = status_filter = cert_filter = []
if data_ok:
    all_rpm   = sorted(df_full["Unidade RPM (Avaliado)"].dropna().unique(), key=rpm_sort_key)
    all_status= ["Aberta","Parcialmente Encerrada","Homologação","Encerrada"]
    all_sit   = ["Comissão Atual","Nota Provisória"]
    all_cert  = ["SIM","NÃO","-"]
    with st.sidebar:
        if st.session_state.show_filtros:
            st.markdown("#### 🔍 Filtros de Visualização")
            rpm_filter = st.multiselect("🏢 Unidade RPM", all_rpm, placeholder="Todas")
            df_tmp = df_full[df_full["Unidade RPM (Avaliado)"].isin(rpm_filter)] if rpm_filter else df_full
            all_unid = sorted(df_tmp["Unidade Principal (Avaliado)"].dropna().unique())
            unid_filter    = st.multiselect("🏛️ Subunidade", all_unid, placeholder="Todas")
            st.markdown("")
            sit_com_filter = st.multiselect("🔵 Situação Comissão", all_sit, placeholder="Todas")
            status_filter  = st.multiselect("📊 Status",            all_status, placeholder="Todos")
            cert_filter    = st.multiselect("✅ Cert. Homologador",  all_cert,   placeholder="Todos")
            st.markdown("---")
            st.markdown(f"<small>🕐 Carregado: {ts}</small>", unsafe_allow_html=True)
            st.markdown(f"<small>📊 {fmt_num(len(df_full))} registros</small>", unsafe_allow_html=True)
    df = apply_filters(df_full, rpm_filter, unid_filter, sit_com_filter, status_filter, cert_filter)
else:
    df = pd.DataFrame()

# ─────────────────────── CABEÇALHO ────────────────────────────────────────────
logo_base64 = ""
if os.path.exists("logo_drh.png"):
    import base64
    with open("logo_drh.png", "rb") as f:
        logo_base64 = base64.b64encode(f.read()).decode("utf-8")

logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="width: 100%; max-width: 480px; height: auto; max-height: 120px; object-fit: contain; border-radius: 8px; align-self: center;" />' if logo_base64 else ""

st.markdown(f"""
<div class="main-title">
  {logo_html}
  <div>
    <h1>AADP 2026 — Análise de Avaliações</h1>
    <p>Polícia Militar de Minas Gerais · Resolução 5458/2025 · Painel de Controle</p>
  </div>
</div>""", unsafe_allow_html=True)

if not data_ok:
    st.error(f"❌ {err_msg}")
    st.markdown(f"""<div class="info-box">
    👈 Configure a pasta dos CSVs na barra lateral.<br>
    📂 Pasta padrão criada: <code>{DADOS_DIR}</code><br>
    Coloque os arquivos <code>avaliacoes.csv</code> e <code>SIGEF.csv</code> nessa pasta.
    </div>""", unsafe_allow_html=True)
    st.stop()

# ─────────────────────── KPI CARDS ────────────────────────────────────────────
n_total  = len(df)
n_enc    = (df["Status Avaliação"]=="Encerrada").sum()
n_hom    = (df["Status Avaliação"]=="Homologação").sum()
n_parc   = (df["Status Avaliação"]=="Parcialmente Encerrada").sum()
n_aberta = (df["Status Avaliação"]=="Aberta").sum()
n_ca     = (df["Situação Comissão"]=="Comissão Atual").sum()
n_np     = (df["Situação Comissão"]=="Nota Provisória").sum()
n_sim    = (df["Certificação Homologador"]=="SIM").sum()

filtro_ativo = bool(rpm_filter or unid_filter or sit_com_filter or status_filter or cert_filter)
ft = f"({fmt_num(n_total)} de {fmt_num(len(df_full))} com filtro)" if filtro_ativo else f"(total: {fmt_num(n_total)})"
st.markdown(f'<div class="info-box">📌 Exibindo {fmt_num(n_total)} avaliações {ft}</div>', unsafe_allow_html=True)

col_block1, col_block2 = st.columns([1, 1.25], gap="large")

with col_block1:
    # Bloco 1: Total em cima, comissão/provisória embaixo
    st.markdown('<div class="kpi-card kpi-total">'
                '<div class="label">TOTAL AVALIAÇÕES</div>'
                f'<div class="value">{fmt_num(n_total)}</div>'
                '<div class="sub">avaliações</div>'
                '</div>', unsafe_allow_html=True)
    st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)
    
    cb1_1, cb1_2 = st.columns(2)
    with cb1_1:
        st.markdown('<div class="kpi-card kpi-ca">'
                    '<div class="label">COMISSÃO ATUAL</div>'
                    f'<div class="value">{fmt_num(n_ca)}</div>'
                    f'<div class="sub">{n_ca/max(n_total,1)*100:.1f}%</div>'
                    '</div>', unsafe_allow_html=True)
    with cb1_2:
        st.markdown('<div class="kpi-card kpi-np">'
                    '<div class="label">NOTA PROVISÓRIA</div>'
                    f'<div class="value">{fmt_num(n_np)}</div>'
                    f'<div class="sub">Hom.SIM: {fmt_num(n_sim)}</div>'
                    '</div>', unsafe_allow_html=True)

with col_block2:
    # Bloco 2: Abertas, Parc, Homologação em cima, Encerradas embaixo
    cb2_1, cb2_2, cb2_3 = st.columns(3)
    with cb2_1:
        st.markdown('<div class="kpi-card kpi-aberta">'
                    '<div class="label">ABERTAS</div>'
                    f'<div class="value">{fmt_num(n_aberta)}</div>'
                    '<div class="sub">AV1 pendente</div>'
                    '</div>', unsafe_allow_html=True)
    with cb2_2:
        st.markdown('<div class="kpi-card kpi-parc">'
                    '<div class="label">PARC. ENCERRADA</div>'
                    f'<div class="value">{fmt_num(n_parc)}</div>'
                    '<div class="sub">AV2 pendente</div>'
                    '</div>', unsafe_allow_html=True)
    with cb2_3:
        st.markdown('<div class="kpi-card kpi-hom">'
                    '<div class="label">HOMOLOGAÇÃO</div>'
                    f'<div class="value">{fmt_num(n_hom)}</div>'
                    '<div class="sub">HOM pendente</div>'
                    '</div>', unsafe_allow_html=True)
                    
    st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)
    
    st.markdown('<div class="kpi-card kpi-enc">'
                '<div class="label">ENCERRADAS</div>'
                f'<div class="value">{fmt_num(n_enc)}</div>'
                f'<div class="sub">{n_enc/max(n_total,1)*100:.1f}%</div>'
                '</div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────── SELEÇÃO DE ABAS VIA SESSION STATE ───────────────────
if "active_page" not in st.session_state:
    st.session_state.active_page = "Análise Gráfica"

active_page = st.session_state.active_page

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — ANÁLISE GRÁFICA
# ══════════════════════════════════════════════════════════════════════════════
if active_page == "Análise Gráfica":
    st.markdown("### 📊 Análise Gráfica das Avaliações")

    # ── LINHA 1: Pizza de Status (destaque, full width) ──────────────────────
    st.markdown("---")
    ordered_labels = ["Aberta","Parcialmente Encerrada","Homologação","Encerrada"]
    sd = df.groupby("Status Avaliação").size()
    vals_pizza  = [int(sd.get(s, 0)) for s in ordered_labels]
    cols_pizza  = [STATUS_COLORS[s] for s in ordered_labels]

    fig_status = go.Figure()

    # Sombra para efeito 3D (circle com proporção assimétrica = elipse)
    fig_status.add_shape(type="circle", xref="paper", yref="paper",
        x0=0.12, y0=0.01, x1=0.88, y1=0.10,
        fillcolor="rgba(0,0,0,0.13)", line_color="rgba(0,0,0,0)", layer="below")

    fig_status.add_trace(go.Pie(
        labels=ordered_labels, values=vals_pizza,
        hole=0.54,
        pull=[0.09, 0.06, 0.04, 0],
        textinfo="label+percent+value",
        textposition="outside",
        textfont=dict(size=13, family="Inter, sans-serif"),
        insidetextorientation="radial",
        marker=dict(colors=cols_pizza, line=dict(color="#ffffff", width=3)),
        hovertemplate="<b>%{label}</b><br>Avaliações: <b>%{value:,}</b><br>%{percent}<extra></extra>",
        sort=False,
    ))
    fig_status.add_annotation(
        text=f"<b>{fmt_num(n_total)}</b><br><span style='font-size:11px;color:#6b7280'>avaliações</span>",
        x=0.5, y=0.5, showarrow=False,
        font=dict(size=22, color="#1F3864", family="Inter"),
        align="center",
    )
    fig_status.update_layout(
        title=dict(text="<b>Status das Avaliações — AADP 2026</b>",
                   font=dict(size=20, color="#1F3864"), x=0.5, y=0.96),
        height=500, showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.12,
                    xanchor="center", x=0.5, font=dict(size=13)),
        paper_bgcolor="white", margin=dict(t=60, b=80, l=100, r=100),
    )
    st.plotly_chart(fig_status, use_container_width=True)

    # ── LINHA 2: Situação Comissão + Status × Comissão ────────────────────────
    st.markdown("---")
    c1, c2 = st.columns([1, 1.6])

    with c1:
        sit_d = df.groupby("Situação Comissão").size().reset_index(name="Qtd")
        fig_sit = go.Figure(go.Pie(
            labels=sit_d["Situação Comissão"], values=sit_d["Qtd"],
            hole=0.50, pull=[0.05,0],
            textinfo="label+percent+value", textposition="outside",
            textfont=dict(size=12), sort=False,
            marker=dict(colors=[SIT_COLORS.get(s,"#aaa") for s in sit_d["Situação Comissão"]],
                        line=dict(color="#fff", width=3)),
            hovertemplate="<b>%{label}</b><br>%{value:,} avaliações (%{percent})<extra></extra>",
        ))
        fig_sit.update_layout(
            title=dict(text="<b>Situação da Comissão</b>", font_size=15, x=0.5),
            height=380, showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="center", x=0.5),
            paper_bgcolor="white", margin=dict(t=50, b=70, l=50, r=50),
        )
        st.plotly_chart(fig_sit, use_container_width=True)

    with c2:
        cross = df.groupby(["Status Avaliação","Situação Comissão"]).size().reset_index(name="Qtd")
        cross["Status Avaliação"] = pd.Categorical(cross["Status Avaliação"],
                                                    categories=ordered_labels, ordered=True)
        cross = cross.sort_values("Status Avaliação")
        fig_bar = px.bar(cross, x="Status Avaliação", y="Qtd", color="Situação Comissão",
                         color_discrete_map=SIT_COLORS, barmode="group", text="Qtd",
                         title="<b>Status × Situação Comissão</b>")
        fig_bar.update_traces(textposition="outside", textfont_size=11)
        fig_bar.update_layout(height=380, title_font_size=15, title_x=0.5,
                               paper_bgcolor="white", plot_bgcolor="#f9f9f9",
                               xaxis_title="", yaxis_title="Qtd",
                               legend=dict(orientation="h", y=1.1, x=1, xanchor="right"))
        fig_bar.update_xaxes(showgrid=False)
        fig_bar.update_yaxes(showgrid=True, gridcolor="#eee")
        st.plotly_chart(fig_bar, use_container_width=True)

    # ── LINHA 3: Barras por RPM ────────────────────────────────────────────────
    st.markdown("---")
    all_units_sorted = sorted(df["Unidade RPM (Avaliado)"].dropna().unique(), key=rpm_sort_key)
    rpm_cross = df.groupby(["Unidade RPM (Avaliado)","Status Avaliação"]).size().reset_index(name="Qtd")
    fig_rpm = px.bar(
        rpm_cross, x="Unidade RPM (Avaliado)", y="Qtd",
        color="Status Avaliação", color_discrete_map=STATUS_COLORS,
        barmode="stack", text_auto=False,
        title="<b>Distribuição por Unidade RPM e Status</b>",
        category_orders={
            "Unidade RPM (Avaliado)": all_units_sorted,
            "Status Avaliação": STACK_ORDER,
        },
    )
    fig_rpm.update_layout(
        height=440, title_font_size=15, title_x=0.5,
        paper_bgcolor="white", plot_bgcolor="#f9f9f9",
        xaxis_title="", yaxis_title="Avaliações",
        legend=dict(orientation="h", y=1.08, x=1, xanchor="right", font_size=11),
    )
    fig_rpm.update_xaxes(tickangle=45, showgrid=False)
    fig_rpm.update_yaxes(showgrid=True, gridcolor="#eee")
    st.plotly_chart(fig_rpm, use_container_width=True)

    # ── LINHA 4: Certificação + Timeline AV1/AV2/HOM ──────────────────────────
    st.markdown("---")
    c3, c4 = st.columns([1, 1.8])

    with c3:
        cert_d = df["Certificação Homologador"].value_counts().reset_index()
        cert_d.columns = ["Cert","Qtd"]
        cert_map = {"SIM":"#FF6B6B","NÃO":"#70AD47","-":"#AAAAAA"}
        fig_cert = px.bar(cert_d, x="Cert", y="Qtd", color="Cert",
                          color_discrete_map=cert_map,
                          title="<b>Certificação Homologador</b>", text="Qtd")
        fig_cert.update_traces(textposition="outside", textfont_size=12)
        fig_cert.update_layout(height=360, title_x=0.5, showlegend=False,
                                paper_bgcolor="white", plot_bgcolor="#f9f9f9",
                                xaxis_title="", yaxis_title="Qtd")
        fig_cert.update_xaxes(showgrid=False)
        fig_cert.update_yaxes(showgrid=True, gridcolor="#eee")
        st.plotly_chart(fig_cert, use_container_width=True)

    with c4:
        # Timeline com AV1, AV2 e HOM diferenciados por cor
        timeline_frames = []
        for col, label in [("Data AV1","AV1 — Avaliador 1"),
                            ("Data AV2","AV2 — Avaliador 2"),
                            ("Data HOM","HOM — Homologador")]:
            sub = df[df[col] != "-"].copy()
            if sub.empty: continue
            try:
                sub["Data"] = pd.to_datetime(sub[col], dayfirst=True, errors="coerce")
                sub = sub.dropna(subset=["Data"])
                cnt = sub.groupby("Data").size().reset_index(name="Qtd")
                cnt["Função"] = label
                timeline_frames.append(cnt)
            except Exception:
                pass
        if timeline_frames:
            df_time = pd.concat(timeline_frames, ignore_index=True)
            fig_time = px.line(
                df_time, x="Data", y="Qtd", color="Função",
                title="<b>Avaliações por Data e Função</b>",
                markers=True,
                color_discrete_map={
                    "AV1 — Avaliador 1":"#4472C4",
                    "AV2 — Avaliador 2":"#ED7D31",
                    "HOM — Homologador":"#70AD47",
                },
            )
            fig_time.update_traces(
                hovertemplate="<b>%{fullData.name}</b><br>"
                              "📅 <b>%{x|%d/%m/%Y}</b><br>"
                              "Avaliações: <b>%{y}</b><extra></extra>",
                line=dict(width=2.5), marker=dict(size=7),
            )
            fig_time.update_layout(
                height=360, title_x=0.5, paper_bgcolor="white", plot_bgcolor="#f9f9f9",
                xaxis_title="", yaxis_title="Avaliações encerradas",
                legend=dict(orientation="h", y=-0.2, x=0.5, xanchor="center", font_size=11),
                hovermode="x unified",
            )
            fig_time.update_xaxes(showgrid=False, tickformat="%d/%m/%Y",
                                   rangeslider_visible=True)
            fig_time.update_yaxes(showgrid=True, gridcolor="#eee")
            st.plotly_chart(fig_time, use_container_width=True)
        else:
            st.info("Sem dados de datas disponíveis para o gráfico de linha do tempo.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — DADOS GERAIS
# ══════════════════════════════════════════════════════════════════════════════
if active_page == "Dados Gerais":
    st.markdown(f"### 📋 Dados Gerais — {fmt_num(len(df))} avaliações")
    cols_d = [
        # Avaliado
        "nrPM (Avaliado)", "Posto/Grad. (Avaliado)", "Nome (Avaliado)",
        "Unidade RPM (Avaliado)", "Unidade Principal (Avaliado)", "Local/Unidade (Avaliado)",
        "Situação Funcional",
        # Datas e Certificação (sem notas/conceitos)
        "Data AV1", "Data AV2", "Data HOM", "Certificação Homologador",
        # Avaliador 1
        "nrPM (Av1)", "Posto (Av1)", "Nome (Av1)", "RPM (Av1)", "Unid. Principal (Av1)",
        # Avaliador 2
        "nrPM (Av2)", "Posto (Av2)", "Nome (Av2)", "RPM (Av2)", "Unid. Principal (Av2)",
        # Homologador
        "nrPM (Hom)", "Posto (Hom)", "Nome (Hom)", "RPM (Hom)", "Unid. Principal (Hom)",
        # Status
        "Situação Comissão", "Status Avaliação",
    ]
    cols_d = [c for c in cols_d if c in df.columns]
    safe_df(df[cols_d].style.map(color_status,subset=["Status Avaliação"])
                            .map(color_sit,   subset=["Situação Comissão"]), height=540)
    csv_d = df[cols_d].to_csv(index=False, sep=";", encoding="utf-8-sig")
    st.download_button("⬇️ Baixar dados filtrados (CSV)", csv_d.encode("utf-8-sig"),
                        f"avaliacoes_filtradas_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                        mime="text/csv")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — AVALIAÇÕES PENDENTES
# ══════════════════════════════════════════════════════════════════════════════
if active_page == "Avaliações Pendentes":
    st.markdown("### ⏳ Avaliações Pendentes")
    STATUS_PEND = {"Homologação","Parcialmente Encerrada","Aberta"}
    df_pend = df[df["Status Avaliação"].isin(STATUS_PEND)].copy()

    c1, c2 = st.columns(2)
    with c1:
        tipo_pend = st.multiselect("Status:", ["Aberta","Parcialmente Encerrada","Homologação"],
                                   default=["Aberta","Parcialmente Encerrada","Homologação"],
                                   key="tp")
    with c2:
        sc_pend = st.multiselect("Situação:", ["Comissão Atual","Nota Provisória"],
                                  default=["Comissão Atual","Nota Provisória"], key="sp")

    df_pv = df_pend[df_pend["Status Avaliação"].isin(tipo_pend) &
                    df_pend["Situação Comissão"].isin(sc_pend)] if tipo_pend and sc_pend else df_pend

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("🔴 Abertas",          (df_pv["Status Avaliação"]=="Aberta").sum())
    k2.metric("🟠 Parc. Encerradas", (df_pv["Status Avaliação"]=="Parcialmente Encerrada").sum())
    k3.metric("🟡 Homologação",       (df_pv["Status Avaliação"]=="Homologação").sum())
    k4.metric("📊 Total",             len(df_pv))

    cols_pend = [
        # ── Avaliado ──────────────────────────────────────────────────────────
        "nrPM (Avaliado)", "Posto/Grad. (Avaliado)", "Nome (Avaliado)",
        "Unidade RPM (Avaliado)", "Unidade Principal (Avaliado)",
        "Situação Funcional",
        # ── Status e Datas (sem notas/conceitos) ─────────────────────────────
        "Status Avaliação", "Situação Comissão",
        "Data AV1", "Data AV2", "Data HOM", "Certificação Homologador",
        # ── Avaliador 1 ───────────────────────────────────────────────────────
        "nrPM (Av1)", "Posto (Av1)", "Nome (Av1)",
        "RPM (Av1)", "Unid. Principal (Av1)",
        # ── Avaliador 2 ───────────────────────────────────────────────────────
        "nrPM (Av2)", "Posto (Av2)", "Nome (Av2)",
        "RPM (Av2)", "Unid. Principal (Av2)",
        # ── Homologador ───────────────────────────────────────────────────────
        "nrPM (Hom)", "Posto (Hom)", "Nome (Hom)",
        "RPM (Hom)", "Unid. Principal (Hom)",
    ]
    # Manter apenas colunas que existem no DataFrame
    cols_pend = [c for c in cols_pend if c in df_pv.columns]

    safe_df(df_pv[cols_pend]
            .sort_values(["Unidade RPM (Avaliado)", "Status Avaliação", "Nome (Avaliado)"])
            .reset_index(drop=True)
            .style.map(color_status, subset=["Status Avaliação"])
            .map(color_sit,          subset=["Situação Comissão"]),
            height=520)
    csv_p = df_pv[cols_pend].sort_values(
        ["Unidade RPM (Avaliado)", "Status Avaliação", "Nome (Avaliado)"]
    ).to_csv(index=False, sep=";", encoding="utf-8-sig")
    st.download_button("⬇️ Baixar pendentes (CSV)", csv_p.encode("utf-8-sig"),
                        f"avaliacoes_pendentes_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                        mime="text/csv")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — AVALIADORES PENDENTES
# ══════════════════════════════════════════════════════════════════════════════
if active_page == "Avaliadores Pendentes":
    st.markdown("### 👥 Avaliadores Pendentes")

    if rpm_filter:
        df_ab = df_full[(df_full["Status Avaliação"] == "Aberta") & (df_full["RPM (Av1)"].isin(rpm_filter))]
        df_pe = df_full[(df_full["Status Avaliação"].isin(["Aberta", "Parcialmente Encerrada"])) & (df_full["RPM (Av2)"].isin(rpm_filter))]
        df_hom = df_full[(df_full["Status Avaliação"] == "Homologação") & (df_full["RPM (Hom)"].isin(rpm_filter))]
    else:
        df_ab = df[df["Status Avaliação"] == "Aberta"]
        df_pe = df[df["Status Avaliação"].isin(["Aberta", "Parcialmente Encerrada"])]
        df_hom = df[df["Status Avaliação"] == "Homologação"]

    av1 = defaultdict(lambda:{"nome":"","posto":"","rpm":"","unid":"","CA":0,"NP":0})
    for _, r in df_ab.iterrows():
        k = r["nrPM (Av1)"]
        if not k: continue
        av1[k].update(nome=r["Nome (Av1)"],posto=r["Posto (Av1)"],rpm=r["RPM (Av1)"],unid=r["Unid. Principal (Av1)"])
        if r["Situação Comissão"]=="Comissão Atual": av1[k]["CA"]+=1
        else: av1[k]["NP"]+=1

    av2 = defaultdict(lambda:{"nome":"","posto":"","rpm":"","unid":"","CA_ab":0,"CA_pe":0,"NP_ab":0,"NP_pe":0})
    for _, r in df_pe.iterrows():
        k = r["nrPM (Av2)"]
        if not k: continue
        av2[k].update(nome=r["Nome (Av2)"],posto=r["Posto (Av2)"],rpm=r["RPM (Av2)"],unid=r["Unid. Principal (Av2)"])
        is_ca = r["Situação Comissão"]=="Comissão Atual"
        if r["Status Avaliação"]=="Aberta":
            if is_ca: av2[k]["CA_ab"]+=1
            else: av2[k]["NP_ab"]+=1
        else:
            if is_ca: av2[k]["CA_pe"]+=1
            else: av2[k]["NP_pe"]+=1

    # AV1
    st.markdown('<div class="section-hdr">👤 AVALIADOR 1 — Avaliações Em Aberto</div>', unsafe_allow_html=True)
    tb1_rows = [{"Nº PM":k,"Nome":d["nome"],"Posto":d["posto"],"RPM":d["rpm"],
        "Unid. Principal":d["unid"],"CA—Aberta":d["CA"],"NP—Aberta":d["NP"],
        "Total AV1":d["CA"]+d["NP"]} for k,d in av1.items() if d["CA"]+d["NP"]>0]
    tb1 = (pd.DataFrame(tb1_rows).sort_values("Total AV1", ascending=False)
           if tb1_rows else pd.DataFrame())
    k1,k2 = st.columns(2)
    k1.metric("Avaliadores pendentes (AV1)", len(tb1))
    k2.metric("Total avaliações Em Aberto", tb1["Total AV1"].sum() if not tb1.empty else 0)
    if not tb1.empty:
        st.dataframe(tb1.reset_index(drop=True), use_container_width=True, height=260)
        st.download_button("⬇️ AV1 (CSV)", tb1.to_csv(index=False,sep=";",encoding="utf-8-sig").encode("utf-8-sig"),
                            f"av1_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv")
    else: st.success("✅ Nenhum AV1 com pendências!")

    # AV2
    st.markdown('<div class="section-hdr">👤 AVALIADOR 2 — Em Aberto + Parcialmente Encerrada</div>', unsafe_allow_html=True)
    tb2_rows = [{"Nº PM":k,"Nome":d["nome"],"Posto":d["posto"],"RPM":d["rpm"],
        "Unid. Principal":d["unid"],"CA—Aberta":d["CA_ab"],"CA—Parc.Enc.":d["CA_pe"],
        "NP—Aberta":d["NP_ab"],"NP—Parc.Enc.":d["NP_pe"],
        "Total AV2":d["CA_ab"]+d["CA_pe"]+d["NP_ab"]+d["NP_pe"]} for k,d in av2.items()
        if d["CA_ab"]+d["CA_pe"]+d["NP_ab"]+d["NP_pe"]>0]
    tb2 = (pd.DataFrame(tb2_rows).sort_values("Total AV2", ascending=False)
           if tb2_rows else pd.DataFrame())
    k1,k2 = st.columns(2)
    k1.metric("Avaliadores pendentes (AV2)", len(tb2))
    k2.metric("Total pendências AV2", tb2["Total AV2"].sum() if not tb2.empty else 0)
    if not tb2.empty:
        st.dataframe(tb2.reset_index(drop=True), use_container_width=True, height=260)
        st.download_button("⬇️ AV2 (CSV)", tb2.to_csv(index=False,sep=";",encoding="utf-8-sig").encode("utf-8-sig"),
                            f"av2_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv")
    else: st.success("✅ Nenhum AV2 com pendências!")

    # HOMOLOGADOR
    st.markdown(
        '<div class="section-hdr-hom">🏛️ HOMOLOGADOR — Avaliações com Divergência Aguardando Nota de Homologação</div>',
        unsafe_allow_html=True)


    # ── Tabela agregada por Homologador (mesmo modelo AV1/AV2) ──────────────────
    hom_map = defaultdict(lambda: {"nome":"","posto":"","rpm":"","unid":"","CA":0,"NP":0})
    for _, r in df_hom.iterrows():
        k = str(r.get("nrPM (Hom)", "")).strip() or "Não identificado"
        hom_map[k]["nome"]  = r.get("Nome (Hom)", "")
        hom_map[k]["posto"] = r.get("Posto (Hom)", "")
        hom_map[k]["rpm"]   = r.get("RPM (Hom)", "")
        hom_map[k]["unid"]  = r.get("Unid. Principal (Hom)", "")
        if r["Situação Comissão"] == "Comissão Atual":
            hom_map[k]["CA"] += 1
        else:
            hom_map[k]["NP"] += 1

    tb3_rows = [{"Nº PM": k, "Nome": d["nome"], "Posto": d["posto"],
                 "RPM": d["rpm"], "Unid. Principal": d["unid"],
                 "CA—Hom.Pend.": d["CA"], "NP—Hom.Pend.": d["NP"],
                 "Total HOM": d["CA"] + d["NP"]}
                for k, d in hom_map.items() if d["CA"] + d["NP"] > 0]
    tb3 = (pd.DataFrame(tb3_rows).sort_values("Total HOM", ascending=False)
           if tb3_rows else pd.DataFrame())

    k1, k2, k3 = st.columns(3)
    k1.metric("Homologadores com pendência", len(tb3))
    k2.metric("Total aguardando HOM", len(df_hom))
    k3.metric("CA / NP pendentes",
              f"{(df_hom['Situação Comissão']=='Comissão Atual').sum()} / "
              f"{(df_hom['Situação Comissão']=='Nota Provisória').sum()}")

    if not tb3.empty:
        st.dataframe(tb3.reset_index(drop=True), use_container_width=True, height=280)
        st.download_button(
            "⬇️ Homologadores pendentes (CSV)",
            tb3.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig"),
            f"hom_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv")
    else:
        st.success("✅ Nenhum Homologador com pendências!")

    # ── Lista detalhada das avaliações em Homologação ──────────────────────────
    if not df_hom.empty:
        with st.expander(f"📋 Ver {len(df_hom)} avaliação(ões) pendentes de homologação", expanded=False):
            cols_det = ["nrPM (Avaliado)", "Posto/Grad. (Avaliado)", "Nome (Avaliado)",
                        "Unidade RPM (Avaliado)", "Unidade Principal (Avaliado)",
                        "Status Avaliação", "Situação Comissão",
                        "Data AV1", "Data AV2", "Data HOM", "Certificação Homologador",
                        "nrPM (Hom)", "Posto (Hom)", "Nome (Hom)", "RPM (Hom)",
                        "Unid. Principal (Hom)"]
            cols_ok = [c for c in cols_det if c in df_hom.columns]
            safe_df(
                df_hom[cols_ok]
                .sort_values(["Unidade RPM (Avaliado)", "Nome (Avaliado)"])
                .reset_index(drop=True)
                .style.map(color_sit, subset=["Situação Comissão"]),
                height=340)
            st.download_button(
                "⬇️ Lista detalhada (CSV)",
                df_hom[cols_ok].to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig"),
                f"hom_detalhe_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — GERAR RELATÓRIO (geração 100% em memória — funciona local e na nuvem)
# ══════════════════════════════════════════════════════════════════════════════

# ── Motor de geração de Excel em memória ─────────────────────────────────────
STATUS_BG_XL = {
    "Encerrada":             "70AD47",
    "Homologação":           "FFD966",
    "Parcialmente Encerrada":"FF8C00",
    "Aberta":                "FF4444",
}
SIT_BG_XL = {"Comissão Atual": "4472C4", "Nota Provisória": "FFC000"}

# Colunas exportadas para Excel (SEM Conceito Geral, Nota Geral, Nota Homologação)
COLS_XLS = [
    "nrPM (Avaliado)", "Posto/Grad. (Avaliado)", "Nome (Avaliado)",
    "Unidade RPM (Avaliado)", "Unidade Principal (Avaliado)", "Local/Unidade (Avaliado)",
    "Situação Funcional",
    "Data AV1", "Data AV2", "Data HOM", "Certificação Homologador",
    "nrPM (Av1)", "Posto (Av1)", "Nome (Av1)", "RPM (Av1)", "Unid. Principal (Av1)",
    "nrPM (Av2)", "Posto (Av2)", "Nome (Av2)", "RPM (Av2)", "Unid. Principal (Av2)",
    "nrPM (Hom)", "Posto (Hom)", "Nome (Hom)", "RPM (Hom)", "Unid. Principal (Hom)",
    "Situação Comissão", "Status Avaliação",
]


def _xl_styles():
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    thin = Side(border_style="thin", color="CCCCCC")
    return {
        "hdr_fill":  PatternFill("solid", fgColor="1F3864"),
        "hdr_font":  Font(bold=True, color="FFFFFF", name="Calibri", size=10),
        "hdr_al":    Alignment(horizontal="center", vertical="center", wrap_text=True),
        "brd":       Border(left=thin, right=thin, top=thin, bottom=thin),
        "data_font": Font(name="Calibri", size=9),
        "center":    Alignment(horizontal="center", vertical="center"),
        "left":      Alignment(vertical="center"),
        "title_font":Font(bold=True, color="FFFFFF", name="Calibri", size=13),
        "title_al":  Alignment(horizontal="center", vertical="center"),
    }


def _write_title(ws, titulo, n_cols, s):
    from openpyxl.styles import PatternFill
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(n_cols, 1))
    c = ws.cell(1, 1, titulo)
    c.fill = PatternFill("solid", fgColor="1F3864")
    c.font = s["title_font"]; c.alignment = s["title_al"]
    ws.row_dimensions[1].height = 22


def _write_headers(ws, cols, s):
    from openpyxl.utils import get_column_letter
    for i, col in enumerate(cols, 1):
        c = ws.cell(2, i, col)
        c.fill = s["hdr_fill"]; c.font = s["hdr_font"]
        c.alignment = s["hdr_al"]; c.border = s["brd"]
    ws.row_dimensions[2].height = 28
    ws.freeze_panes = "A3"
    return len(cols)


def _write_data_rows(ws, df, cols, s):
    from openpyxl.styles import PatternFill, Font
    for r, (_, row) in enumerate(df.iterrows(), 3):
        for ci, col in enumerate(cols, 1):
            val = row.get(col, "")
            cell = ws.cell(r, ci, str(val) if pd.notna(val) and val != "" else "")
            cell.font = s["data_font"]; cell.border = s["brd"]
            cell.alignment = s["center"] if ci > 7 else s["left"]
            if col == "Status Avaliação" and val in STATUS_BG_XL:
                cell.fill = PatternFill("solid", fgColor=STATUS_BG_XL[val])
                cell.font = Font(bold=True, name="Calibri", size=9,
                                  color="FFFFFF" if val == "Aberta" else "000000")
            elif col == "Situação Comissão" and val in SIT_BG_XL:
                cell.fill = PatternFill("solid", fgColor=SIT_BG_XL[val])
                cell.font = Font(bold=True, name="Calibri", size=9,
                                  color="FFFFFF" if val == "Comissão Atual" else "000000")


def _auto_widths(ws, df, cols):
    from openpyxl.utils import get_column_letter
    for ci, col in enumerate(cols, 1):
        max_len = 0
        if col in df.columns and not df.empty:
            max_val = df[col].astype(str).str.len().max()
            if pd.notna(max_val):
                max_len = int(max_val)
        max_len = max(len(str(col)), max_len)
        ws.column_dimensions[get_column_letter(ci)].width = min(max(max_len * 0.92, 8), 40)


def _write_data_sheet(ws, df, titulo, cols, s):
    """Escreve título + cabeçalhos + dados em uma aba."""
    df_c = df[[c for c in cols if c in df.columns]].reset_index(drop=True)
    actual_cols = list(df_c.columns)
    _write_title(ws, titulo, len(actual_cols), s)
    _write_headers(ws, actual_cols, s)
    _write_data_rows(ws, df_c, actual_cols, s)
    _auto_widths(ws, df_c, actual_cols)


def _write_avaliadores_sheet(ws, df_unit, s, df_global=None, titulo_unidade=""):
    """Aba Avaliadores Pendentes: avaliadores lotados na unidade com pendências."""
    from openpyxl.styles import PatternFill, Font, Alignment
    _write_title(ws, "AVALIADORES PENDENTES — LOTADOS NA UNIDADE", 14, s)

    is_geral = "GERAL" in str(titulo_unidade).upper()

    if df_global is not None and not is_geral:
        df_ab = df_global[(df_global["Status Avaliação"] == "Aberta") & (df_global["RPM (Av1)"] == titulo_unidade)]
        df_pe = df_global[(df_global["Status Avaliação"].isin(["Aberta", "Parcialmente Encerrada"])) & (df_global["RPM (Av2)"] == titulo_unidade)]
        df_hom = df_global[(df_global["Status Avaliação"] == "Homologação") & (df_global["RPM (Hom)"] == titulo_unidade)]
    else:
        df_ab = df_unit[df_unit["Status Avaliação"] == "Aberta"]
        df_pe = df_unit[df_unit["Status Avaliação"].isin(["Aberta", "Parcialmente Encerrada"])]
        df_hom = df_unit[df_unit["Status Avaliação"] == "Homologação"]

    av1 = defaultdict(lambda: {"nome":"","posto":"","rpm":"","unid":"","CA":0,"NP":0})
    for _, r in df_ab.iterrows():
        k = str(r.get("nrPM (Av1)","")).strip()
        if not k: continue
        av1[k].update(nome=r.get("Nome (Av1)",""), posto=r.get("Posto (Av1)",""),
                      rpm=r.get("RPM (Av1)",""), unid=r.get("Unid. Principal (Av1)",""))
        if r["Situação Comissão"] == "Comissão Atual": av1[k]["CA"] += 1
        else: av1[k]["NP"] += 1

    av2 = defaultdict(lambda: {"nome":"","posto":"","rpm":"","unid":"","CA_ab":0,"CA_pe":0,"NP_ab":0,"NP_pe":0})
    for _, r in df_pe.iterrows():
        k = str(r.get("nrPM (Av2)","")).strip()
        if not k: continue
        av2[k].update(nome=r.get("Nome (Av2)",""), posto=r.get("Posto (Av2)",""),
                      rpm=r.get("RPM (Av2)",""), unid=r.get("Unid. Principal (Av2)",""))
        is_ca = r["Situação Comissão"] == "Comissão Atual"
        if r["Status Avaliação"] == "Aberta":
            if is_ca: av2[k]["CA_ab"] += 1
            else: av2[k]["NP_ab"] += 1
        else:
            if is_ca: av2[k]["CA_pe"] += 1
            else: av2[k]["NP_pe"] += 1

    hom = defaultdict(lambda: {"nome":"","posto":"","rpm":"","unid":"","CA":0,"NP":0})
    for _, r in df_hom.iterrows():
        k = str(r.get("nrPM (Hom)","")).strip() or "N/I"
        hom[k].update(nome=r.get("Nome (Hom)",""), posto=r.get("Posto (Hom)",""),
                      rpm=r.get("RPM (Hom)",""), unid=r.get("Unid. Principal (Hom)",""))
        if r["Situação Comissão"] == "Comissão Atual": hom[k]["CA"] += 1
        else: hom[k]["NP"] += 1

    row_num = 3
    # Cabeçalho seção AV1
    titles_av1 = ["Nº PM","Posto","Nome","RPM","Unidade","CA—Aberta","NP—Aberta","Total AV1"]
    fill_av1 = PatternFill("solid", fgColor="1F3864")
    for ci, h in enumerate(titles_av1, 1):
        c = ws.cell(row_num, ci, h)
        c.fill = fill_av1; c.font = s["hdr_font"]; c.alignment = s["hdr_al"]; c.border = s["brd"]
    ws.merge_cells(start_row=row_num-1, start_column=1, end_row=row_num-1, end_column=8)
    lbl = ws.cell(row_num-1, 1, "AVALIADOR 1 — Avaliações Em Aberto")
    lbl.fill = PatternFill("solid", fgColor="2E5090"); lbl.font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    lbl.alignment = Alignment(horizontal="center", vertical="center")
    row_num += 1
    for k, d in sorted(av1.items(), key=lambda x: -(x[1]["CA"]+x[1]["NP"])):
        tot = d["CA"]+d["NP"]
        if tot == 0: continue
        for ci, val in enumerate([k, d["posto"], d["nome"], d["rpm"], d["unid"],
                                    d["CA"], d["NP"], tot], 1):
            c = ws.cell(row_num, ci, val)
            c.font = s["data_font"]; c.border = s["brd"]; c.alignment = s["center"]
        row_num += 1

    row_num += 2
    # Cabeçalho seção AV2
    titles_av2 = ["Nº PM","Posto","Nome","RPM","Unidade","CA—Ab.","CA—PE","NP—Ab.","NP—PE","Total AV2"]
    ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=10)
    lbl2 = ws.cell(row_num, 1, "AVALIADOR 2 — Em Aberto + Parcialmente Encerrada")
    lbl2.fill = PatternFill("solid", fgColor="2E5090"); lbl2.font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    lbl2.alignment = Alignment(horizontal="center", vertical="center")
    row_num += 1
    for ci, h in enumerate(titles_av2, 1):
        c = ws.cell(row_num, ci, h)
        c.fill = fill_av1; c.font = s["hdr_font"]; c.alignment = s["hdr_al"]; c.border = s["brd"]
    row_num += 1
    for k, d in sorted(av2.items(), key=lambda x: -(x[1]["CA_ab"]+x[1]["CA_pe"]+x[1]["NP_ab"]+x[1]["NP_pe"])):
        tot = d["CA_ab"]+d["CA_pe"]+d["NP_ab"]+d["NP_pe"]
        if tot == 0: continue
        for ci, val in enumerate([k, d["posto"], d["nome"], d["rpm"], d["unid"],
                                    d["CA_ab"], d["CA_pe"], d["NP_ab"], d["NP_pe"], tot], 1):
            c = ws.cell(row_num, ci, val)
            c.font = s["data_font"]; c.border = s["brd"]; c.alignment = s["center"]
        row_num += 1

    row_num += 2
    # Cabeçalho seção HOM
    titles_hom = ["Nº PM","Posto","Nome","RPM","Unidade","CA—Pend.","NP—Pend.","Total HOM"]
    ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=8)
    lbl3 = ws.cell(row_num, 1, "HOMOLOGADOR — Avaliações com Divergência Aguardando Nota")
    lbl3.fill = PatternFill("solid", fgColor="7B3F00"); lbl3.font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    lbl3.alignment = Alignment(horizontal="center", vertical="center")
    row_num += 1
    for ci, h in enumerate(titles_hom, 1):
        c = ws.cell(row_num, ci, h)
        c.fill = PatternFill("solid", fgColor="7B3F00"); c.font = s["hdr_font"]
        c.alignment = s["hdr_al"]; c.border = s["brd"]
    row_num += 1
    for k, d in sorted(hom.items(), key=lambda x: -(x[1]["CA"]+x[1]["NP"])):
        tot = d["CA"]+d["NP"]
        if tot == 0: continue
        for ci, val in enumerate([k, d["posto"], d["nome"], d["rpm"], d["unid"],
                                    d["CA"], d["NP"], tot], 1):
            c = ws.cell(row_num, ci, val)
            c.font = s["data_font"]; c.border = s["brd"]; c.alignment = s["center"]
        row_num += 1

    from openpyxl.utils import get_column_letter
    for ci in range(1, 15):
        ws.column_dimensions[get_column_letter(ci)].width = 18


def _write_resumo_sheet(ws, df, titulo, s):
    """Aba Resumo: grade CA × Status e NP × Status com cores."""
    from openpyxl.styles import PatternFill, Font, Alignment
    _write_title(ws, f"RESUMO — {titulo}", 5, s)

    STATUS_ORD = ["Aberta", "Parcialmente Encerrada", "Homologação", "Encerrada"]
    ca = df[df["Situação Comissão"] == "Comissão Atual"]
    np_ = df[df["Situação Comissão"] == "Nota Provisória"]

    fill_ca  = PatternFill("solid", fgColor="4472C4")   # azul  — Comissão Atual
    fill_np  = PatternFill("solid", fgColor="FFC000")   # amarelo — Nota Provisória
    fill_tot = PatternFill("solid", fgColor="1F3864")   # azul escuro — Total
    fill_st  = {
        "Aberta":                 PatternFill("solid", fgColor="FF4444"),
        "Parcialmente Encerrada": PatternFill("solid", fgColor="FF8C00"),
        "Homologação":            PatternFill("solid", fgColor="FFD966"),
        "Encerrada":              PatternFill("solid", fgColor="70AD47"),
    }
    white_bold  = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    black_bold  = Font(bold=True, color="000000", name="Calibri", size=11)
    center_bold = Alignment(horizontal="center", vertical="center")
    thin = s["brd"]

    r = 3
    # Cabeçalho da tabela
    headers = ["STATUS / SITUAÇÃO", "COMISSÃO ATUAL ✅", "NOTA PROVISÓRIA ⚠️",
               "TOTAL STATUS", "% do Total"]
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(r, ci, h)
        cell.fill = fill_tot; cell.font = white_bold
        cell.alignment = center_bold; cell.border = thin
        ws.column_dimensions[ws.cell(r, ci).column_letter].width = 26 if ci == 1 else 18
    ws.row_dimensions[r].height = 30
    r += 1

    total = len(df)
    for st in STATUS_ORD:
        ca_n  = (ca["Status Avaliação"] == st).sum()
        np_n  = (np_["Status Avaliação"] == st).sum()
        tot_n = ca_n + np_n
        pct   = f"{tot_n/total*100:.1f}%" if total > 0 else "0%"

        # Coluna Status
        c0 = ws.cell(r, 1, st)
        c0.fill = fill_st.get(st, PatternFill()); c0.border = thin
        c0.font = black_bold if st in ("Homologação","Parcialmente Encerrada","Encerrada") else white_bold
        c0.alignment = center_bold

        # CA
        c1 = ws.cell(r, 2, ca_n)
        c1.fill = fill_ca; c1.font = white_bold; c1.alignment = center_bold; c1.border = thin

        # NP
        c2 = ws.cell(r, 3, np_n)
        c2.fill = fill_np; c2.font = black_bold; c2.alignment = center_bold; c2.border = thin

        # Total
        c3 = ws.cell(r, 4, tot_n)
        c3.fill = fill_tot; c3.font = white_bold; c3.alignment = center_bold; c3.border = thin

        # %
        c4 = ws.cell(r, 5, pct)
        c4.font = Font(name="Calibri", size=11); c4.alignment = center_bold; c4.border = thin

        ws.row_dimensions[r].height = 24
        r += 1

    # Linha TOTAL GERAL
    ca_total  = len(ca); np_total  = len(np_)
    for ci, val in enumerate(["TOTAL GERAL", ca_total, np_total, total, "100%"], 1):
        c = ws.cell(r, ci, val)
        c.fill = fill_tot; c.font = white_bold; c.alignment = center_bold; c.border = thin
    ws.row_dimensions[r].height = 28

    # Legenda
    r += 2
    for ci, (txt, fill, fnt) in enumerate([
        ("COMISSÃO ATUAL — Policial está lotado na unidade avaliadora (CA)", fill_ca, white_bold),
        ("NOTA PROVISÓRIA — Policial transferido; nota pode mudar (NP)",     fill_np, black_bold),
    ], 1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        lc = ws.cell(r, 1, txt)
        lc.fill = fill; lc.font = fnt
        lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[r].height = 20
        r += 1


def _write_analise_sheet(ws, df, titulo, s):
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.chart import PieChart3D, Reference
    from openpyxl.chart.series import DataPoint

    _write_title(ws, f"ANÁLISE — {titulo}", 6, s)

    STATUS_ORD = ["Aberta", "Parcialmente Encerrada", "Homologação", "Encerrada"]
    ca  = df[df["Situação Comissão"] == "Comissão Atual"]
    np_ = df[df["Situação Comissão"] == "Nota Provisória"]
    total = len(df)

    fill_ca  = PatternFill("solid", fgColor="4472C4")
    fill_np  = PatternFill("solid", fgColor="FFC000")
    fill_tot = PatternFill("solid", fgColor="1F3864")
    fill_st  = {
        "Aberta":                 PatternFill("solid", fgColor="FF4444"),
        "Parcialmente Encerrada": PatternFill("solid", fgColor="FF8C00"),
        "Homologação":            PatternFill("solid", fgColor="FFD966"),
        "Encerrada":              PatternFill("solid", fgColor="70AD47"),
    }
    white_bold  = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    black_bold  = Font(bold=True, color="000000", name="Calibri", size=11)
    center_al   = Alignment(horizontal="center", vertical="center")
    thin = s["brd"]

    # ── Cabeçalho da tabela ────────────────────────────────────────────────────
    r = 3
    headers = ["STATUS", "COMISSÃO ATUAL", "NOTA PROVISÓRIA", "TOTAL", "%"]
    col_fills = [fill_tot, fill_ca, fill_np, fill_tot, fill_tot]
    col_fonts = [white_bold, white_bold, black_bold, white_bold, white_bold]
    for ci, (h, fll, fnt) in enumerate(zip(headers, col_fills, col_fonts), 1):
        c = ws.cell(r, ci, h)
        c.fill = fll; c.font = fnt; c.alignment = center_al; c.border = thin
        ws.column_dimensions[ws.cell(r, ci).column_letter].width = 26 if ci == 1 else 18
    ws.row_dimensions[r].height = 28
    r += 1

    # ── Dados por Status ───────────────────────────────────────────────────────
    data_start_row = r  # usado pelo gráfico
    for st in STATUS_ORD:
        ca_n  = int((ca["Status Avaliação"] == st).sum())
        np_n  = int((np_["Status Avaliação"] == st).sum())
        tot_n = ca_n + np_n
        pct   = f"{tot_n/total*100:.1f}%" if total > 0 else "0%"

        fll_st = fill_st.get(st, PatternFill())
        use_white = st == "Aberta"

        c0 = ws.cell(r, 1, st)
        c0.fill = fll_st; c0.border = thin
        c0.font = white_bold if use_white else black_bold
        c0.alignment = center_al

        c1 = ws.cell(r, 2, ca_n)
        c1.fill = fill_ca; c1.font = white_bold; c1.alignment = center_al; c1.border = thin

        c2 = ws.cell(r, 3, np_n)
        c2.fill = fill_np; c2.font = black_bold; c2.alignment = center_al; c2.border = thin

        c3 = ws.cell(r, 4, tot_n)
        c3.fill = fill_tot; c3.font = white_bold; c3.alignment = center_al; c3.border = thin

        c4 = ws.cell(r, 5, pct)
        c4.font = Font(name="Calibri", size=11); c4.alignment = center_al; c4.border = thin

        ws.row_dimensions[r].height = 22
        r += 1
    data_end_row = r - 1

    # Linha TOTAL GERAL
    ca_tot = len(ca); np_tot = len(np_)
    for ci, val in enumerate(["TOTAL GERAL", ca_tot, np_tot, total, "100%"], 1):
        c = ws.cell(r, ci, val)
        c.fill = fill_tot; c.font = white_bold; c.alignment = center_al; c.border = thin
    ws.row_dimensions[r].height = 26
    total_row = r
    r += 2

    # Legenda
    for txt, fll, fnt in [
        ("COMISSÃO ATUAL — Policial lotado na unidade avaliadora", fill_ca, white_bold),
        ("NOTA PROVISÓRIA — Policial transferido; nota pode mudar", fill_np, black_bold),
    ]:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        lc = ws.cell(r, 1, txt)
        lc.fill = fll; lc.font = fnt
        lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[r].height = 18
        r += 1

    # ── Gráfico 1: Pizza por Status (Total por status - 3D) ─────────────────────────
    pie1 = PieChart3D()
    pie1.title  = "Status das Avaliações"
    pie1.style  = 10
    pie1.width  = 14
    pie1.height = 10

    # Dados: coluna 4 (TOTAL) linhas data_start_row até data_end_row
    data1 = Reference(ws, min_col=4, min_row=data_start_row, max_row=data_end_row)
    cats1 = Reference(ws, min_col=1, min_row=data_start_row, max_row=data_end_row)
    pie1.add_data(data1)
    pie1.set_categories(cats1)
    pie1.series[0].title = None

    # Cores manuais para cada fatia (ordem: Aberta, Parc.Enc, Hom, Encerrada)
    SLICE_COLORS = ["FF4444", "FF8C00", "FFD966", "70AD47"]
    for idx, hex_color in enumerate(SLICE_COLORS):
        pt = DataPoint(idx=idx)
        pt.graphicalProperties.solidFill = hex_color
        pie1.series[0].dPt.append(pt)

    # Legenda na lateral direita
    pie1.legend.position = "r"

    from openpyxl.chart.label import DataLabelList
    pie1.dataLabels = DataLabelList()
    pie1.dataLabels.showPercent     = False
    pie1.dataLabels.showCatName     = False
    pie1.dataLabels.showVal         = True
    pie1.dataLabels.showLeaderLines = True

    # Posiciona o gráfico na coluna G linha 3
    ws.add_chart(pie1, "G3")


def _build_workbook(df_unit: pd.DataFrame, titulo: str, df_global: pd.DataFrame = None) -> bytes:
    """Monta workbook completo com 4 abas para uma unidade (sem Resumo duplicado)."""
    from openpyxl import Workbook
    s = _xl_styles()
    wb = Workbook()

    is_geral = "GERAL" in str(titulo).upper()
    cols = [c for c in COLS_XLS if c in df_unit.columns]
    
    if is_geral:
        sensitive_cols = [c for c in [
            "Conceito Geral", "Nota Geral", "Nota Homologação",
            "Competência 1", "Conceito Comp.1", "Nota Comp.1",
            "Competência 2", "Conceito Comp.2", "Nota Comp.2",
            "Competência 3", "Conceito Comp.3", "Nota Comp.3",
            "Competência 4", "Conceito Comp.4", "Nota Comp.4"
        ] if c in df_unit.columns]
        
        if "Data HOM" in cols:
            idx = cols.index("Data HOM") + 1
            cols = cols[:idx] + sensitive_cols + cols[idx:]
        else:
            cols.extend(sensitive_cols)

    # Aba 1 — Geral
    ws1 = wb.active; ws1.title = "Geral"
    _write_data_sheet(ws1, df_unit, f"AVALIAÇÕES — {titulo}", cols, s)

    # Aba 2 — Avaliações Pendentes (Aberta, Parcialmente Encerrada e Homologação)
    ws2 = wb.create_sheet("Avaliações Pendentes")
    df_pend = df_unit[df_unit["Status Avaliação"].isin(["Aberta", "Parcialmente Encerrada", "Homologação"])]
    _write_data_sheet(ws2, df_pend, f"AVALIAÇÕES PENDENTES — {titulo}", cols, s)

    # Aba 3 — Avaliadores Pendentes
    ws3 = wb.create_sheet("Avaliadores Pendentes")
    _write_avaliadores_sheet(ws3, df_unit, s, df_global=df_global, titulo_unidade=titulo)

    # Aba 4 — Análise (tabela + gráficos de pizza)
    ws4 = wb.create_sheet("Análise")
    _write_analise_sheet(ws4, df_unit, titulo, s)

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf.read()


def _gerar_zip_bytes(df_full: pd.DataFrame, modo: str, units_sel: list) -> tuple:
    """Gera ZIP com planilhas Excel em memória e retorna (bytes, filename)."""
    zip_buf  = io.BytesIO()
    zip_name = f"AADP_2026_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        if modo in ("all", "geral"):
            zf.writestr("Analise_Avaliacoes_Geral.xlsx",
                        _build_workbook(df_full, "GERAL — AADP 2026", df_full))
        if modo in ("all", "units"):
            targets = units_sel if units_sel else sorted(
                df_full["Unidade RPM (Avaliado)"].dropna().unique(), key=rpm_sort_key)
            for rpm in targets:
                mask   = df_full["Unidade RPM (Avaliado)"] == rpm
                df_rpm = df_full[mask].copy()
                safe   = re.sub(r'[^\w]', '_', str(rpm))
                zf.writestr(f"Analise_Avaliacoes_{safe}.xlsx",
                            _build_workbook(df_rpm, rpm, df_full))

    zip_buf.seek(0)
    return zip_buf.read(), zip_name

# ── Interface da aba ──────────────────────────────────────────────────────────
if active_page == "Gerar Relatório":
    st.markdown("### 📥 Gerar Relatório Excel + ZIP")
    st.markdown('<div class="info-box">'
                '✅ Geração <b>100% em memória</b> — funciona tanto no servidor local '
                'quanto no <b>Streamlit Community Cloud</b> (online). '
                'O arquivo ZIP é baixado diretamente no seu navegador.</div>',
                unsafe_allow_html=True)

    modo_rel = st.radio("📋 Tipo de relatório:", [
        "🌐 Completo — Geral + todas as Unidades RPM",
        "📋 Somente Planilha Geral",
        "🎯 Unidades RPM específicas",
    ], horizontal=False)

    units_sel = []
    if "específicas" in modo_rel:
        all_rpms_sorted = sorted(df_full["Unidade RPM (Avaliado)"].dropna().unique(),
                                  key=rpm_sort_key)
        units_sel = st.multiselect("Selecione as Unidades RPM:", all_rpms_sorted,
                                    placeholder="Escolha uma ou mais unidades...")
        if units_sel:
            n_prev = sum((df_full["Unidade RPM (Avaliado)"]==u).sum() for u in units_sel)
            st.markdown(f"<small>📊 {len(units_sel)} unidade(s) · {fmt_num(n_prev)} registros</small>",
                        unsafe_allow_html=True)

    st.markdown("---")

    if "específicas" in modo_rel and not units_sel:
        st.warning("⚠️ Selecione ao menos uma Unidade RPM.")
    else:
        if st.button("🚀 Gerar e Baixar Relatório", type="primary", use_container_width=True):
            modo_code = ("geral" if "Somente" in modo_rel
                         else "units" if "específicas" in modo_rel else "all")
            with st.spinner("⏳ Gerando planilhas Excel... aguarde."):
                try:
                    zip_bytes, zip_name = _gerar_zip_bytes(df_full, modo_code, units_sel)
                    st.success(f"✅ ZIP gerado com sucesso! ({len(zip_bytes)//1024:,} KB)")
                    st.download_button(
                        label=f"⬇️ Baixar {zip_name}",
                        data=zip_bytes,
                        file_name=zip_name,
                        mime="application/zip",
                        use_container_width=True,
                    )
                except Exception as ex:
                    st.error(f"❌ Erro na geração: {ex}")

    st.markdown("---")
    st.markdown("#### 🔄 Ciclo de Atualização")
    st.markdown("""
    **Modo local/servidor:**
    1. Substitua `avaliacoes.csv` / `SIGEF.csv` na pasta **dados/**
    2. Clique **🔄 Recarregar Dados** na barra lateral

    **Modo Google Drive:**
    1. Substitua os arquivos na pasta do Drive
    2. Clique **🔄 Recarregar Dados** — o app baixa automaticamente a versão mais recente
    """)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — HOMOLOGAÇÃO RELATÓRIO WORD (.DOCX)
# ══════════════════════════════════════════════════════════════════════════════
if active_page == "Relatório Word":
    st.markdown("### 📄 Relatório Word (.docx)")
    st.markdown('<div class="info-box">'
                '📄 <b>Geração de Relatório Executivo:</b> Esta aba gera um relatório gerencial completo '
                'em formato Word com base no arquivo Geral.xlsx.</div>',
                unsafe_allow_html=True)
                
    df_word = None
    
    if "Drive" in fonte:
        # ── Modo Google Drive ──────────────────────────────────────────────
        drive_geral_id = cfg.get("drive_geral_id", "")
        if not drive_geral_id:
            st.warning("⚠️ O ID do arquivo Geral.xlsx no Google Drive não foi configurado. "
                       "Por favor, informe-o no menu 'Fonte dos Dados' na barra lateral esquerda.")
        else:
            @st.cache_data(show_spinner="⏳ Baixando planilha Geral.xlsx do Google Drive...")
            def load_geral_from_drive(file_id):
                tmp_dir = tempfile.mkdtemp(prefix="aadp_word_")
                dest_path = os.path.join(tmp_dir, "Geral.xlsx")
                _baixar_drive(file_id, dest_path)
                df = pd.read_excel(dest_path, sheet_name="Geral", skiprows=1)
                try:
                    os.remove(dest_path)
                    os.rmdir(tmp_dir)
                except:
                    pass
                return df
                
            try:
                df_word = load_geral_from_drive(drive_geral_id)
                st.success("✅ Planilha Geral.xlsx baixada e processada do Google Drive!")
                st.info(f"📊 {len(df_word):,} registros carregados com sucesso!")
            except Exception as e:
                st.error(f"❌ Erro ao obter planilha do Google Drive: {e}")
    else:
        # ── Modo Pasta Local ───────────────────────────────────────────────
        default_excel_path = r"c:\Users\guilh\Downloads\analise AADP 2026\Geral.xlsx"
        src_option = st.radio("Selecione a fonte de dados:", [
            "📂 Usar arquivo padrão local (Geral.xlsx)",
            "📤 Fazer upload de outra planilha Geral.xlsx"
        ])
        
        if "padrão" in src_option:
            if os.path.exists(default_excel_path):
                st.success(f"✅ Arquivo local Geral.xlsx encontrado!")
                try:
                    df_word = pd.read_excel(default_excel_path, sheet_name="Geral", skiprows=1)
                    st.info(f"📊 {len(df_word):,} registros carregados com sucesso!")
                except Exception as e:
                    st.error(f"❌ Erro ao ler a aba 'Geral' do arquivo padrão: {e}")
            else:
                st.error(f"❌ Arquivo padrão não encontrado em: {default_excel_path}")
        else:
            uploaded_file = st.file_uploader("Escolha a planilha Geral.xlsx", type=["xlsx"])
            if uploaded_file is not None:
                try:
                    df_word = pd.read_excel(uploaded_file, sheet_name="Geral", skiprows=1)
                    st.success(f"✅ {len(df_word):,} registros carregados do arquivo enviado!")
                except Exception as e:
                    st.error(f"❌ Erro ao ler o arquivo enviado: {e}")
                
    if df_word is not None:
        st.markdown("---")
        st.markdown("#### Configurações do Relatório")
        
        # Opções solicitadas pelo negócio
        rel_scope = st.radio("Escopo do Relatório:", [
            "🏢 Geral RPM (Somente consolidados das unidades principais UDI/UDG)",
            "🌐 Geral Subordinadas (Completo com UDI/UDG e unidades subordinadas)",
            "🎯 Por RPM específica (Filtrar por unidades específicas)"
        ])
        
        selected_rpms = []
        if "específica" in rel_scope:
            unique_rpms = sorted(df_word["Unidade RPM (Avaliado)"].dropna().unique().tolist(), key=rpm_sort_key)
            selected_rpms = st.multiselect("Selecione as Unidades UDI/UDG para o relatório:", unique_rpms)
            
        st.markdown("---")
        
        # Validação de botão de geração
        if "específica" in rel_scope and not selected_rpms:
            st.warning("⚠️ Selecione ao menos uma unidade para gerar o relatório.")
        else:
            if st.button("🚀 Gerar e Baixar Relatório Word", key="btn_word_gen"):
                with st.spinner("⏳ Gerando relatório executivo Word com gráficos... (Isso pode levar alguns instantes)"):
                    try:
                        from gerar_relatorio_word import generate_word_report
                        
                        # Mapear escopo para código
                        mode_code = "geral_rpm" if "Geral RPM" in rel_scope else "geral_subordinadas" if "Geral Subordinadas" in rel_scope else "especifica"
                        
                        doc_bytes = generate_word_report(df_word, mode_code, selected_rpms)
                        
                        st.success("✅ Relatório Word gerado com sucesso!")
                        
                        doc_name = f"Relatorio_Executivo_AADP2026_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                        
                        st.download_button(
                            label=f"⬇️ Baixar {doc_name}",
                            data=doc_bytes,
                            file_name=doc_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    except Exception as ex:
                        st.error(f"❌ Erro ao gerar o relatório: {ex}")

# ─────────────────────── RODAPÉ ───────────────────────────────────────────────
st.markdown("---")
st.markdown(f"<center><small>AADP 2026 · Polícia Militar de Minas Gerais · "
            f"Resolução 5458/2025 · {datetime.now().strftime('%d/%m/%Y')}</small></center>",
            unsafe_allow_html=True)
