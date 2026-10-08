"""


AADP 2026 — Dashboard de Análise de Avaliações


Versão 3.0 — Google Drive + Streamlit Cloud + Geração em memória


"""


import time


_prof_start = time.time()


import streamlit as st


import pandas as pd


import plotly.express as px


import plotly.graph_objects as go


import io, os, re, json, subprocess, unicodedata, csv, tempfile, zipfile, sqlite3, hashlib


from pathlib import Path


from collections import defaultdict


from datetime import datetime, timezone, timedelta





def now_br():


    """Retorna datetime no fuso de Brasília (UTC-3)."""


    return datetime.now(timezone(timedelta(hours=-3)))


@st.cache_data(ttl=600)


def get_last_updated_time(av_f, drive_av_id=None):


    """Retorna a data e hora de consolidação dos dados em horário de Brasília."""


    import requests, email.utils, os


    dt_utc = None


    if drive_av_id:


        try:


            url = f"https://drive.google.com/uc?id={drive_av_id}&export=download"


            r = requests.head(url, allow_redirects=True, timeout=5)


            last_mod = r.headers.get("Last-Modified")


            if last_mod:


                dt_utc = email.utils.parsedate_to_datetime(last_mod)


        except Exception:


            pass


            


    if not dt_utc and os.path.exists(av_f):


        try:


            mtime = os.path.getmtime(av_f)


            dt_utc = datetime.fromtimestamp(mtime, timezone.utc)


        except Exception:


            pass


            


    if dt_utc:


        tz_br = timezone(timedelta(hours=-3))


        dt_br = dt_utc.astimezone(tz_br)


        return dt_br.strftime("%d/%m/%Y, às %H:%M horas")


        


    return "Data/Hora indisponível"







def normalize_pm(v):
    if pd.isna(v): return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.lstrip("0")

def _c_aadp_type(sit):
    sit = str(sit).strip().upper()
    if sit in ["OF. QOR RECONDUZIDO", "OFICIAL QOR DESIGNAD", "PRACA QPR RECONDUZID", "PRACA QPR DESIGNADO"]:
        return "Reconduzido"
    if sit in ["MATRICULADO EM CURSO"]:
        return "Discente"
    return "Ativa"


def parse_float(v):
    if pd.isna(v): return None
    s = str(v).strip().replace(",", ".")
    if s in ("", "-", "nan", "none"): return None
    try:
        return float(s)
    except ValueError:
        return None

@st.cache_resource(show_spinner=False)
def run_grades_audit(xlsx_path, csv_path, drive_geral_id=None, drive_master_xlsx_id=None):
    import pandas as pd
    import numpy as np
    import csv
    import os
    
    # Se estamos no modo Drive e temos os IDs configurados, baixar os arquivos primeiro
    if drive_geral_id:
        try:
            if not os.path.exists(csv_path) or os.path.getsize(csv_path) == 0:
                _baixar_drive(drive_geral_id, csv_path)
        except Exception as e:
            return None, f"Falha ao baixar geral.csv do Google Drive (ID: {drive_geral_id}): {str(e)}"
            
    if drive_master_xlsx_id:
        try:
            if not os.path.exists(xlsx_path) or os.path.getsize(xlsx_path) == 0:
                _baixar_drive(drive_master_xlsx_id, xlsx_path)
        except Exception as e:
            return None, f"Falha ao baixar master Excel do Google Drive (ID: {drive_master_xlsx_id}): {str(e)}"

    CONCEITO_FAIXA = {
        "nivel superior de desempenho":       (9.00, 10.00),
        "nivel alto de desempenho":           (7.00,  8.99),
        "nivel intermediario de desempenho":  (6.00,  6.99),
        "nivel baixo de desempenho":          (3.00,  5.99),
        "nivel inferior de desempenho":       (0.00,  2.99),
    }

    def normaliza(texto: str) -> str:
        if not isinstance(texto, str): return ""
        import unicodedata
        t = unicodedata.normalize("NFD", texto.lower())
        return "".join(c for c in t if unicodedata.category(c) != "Mn")

    try:
        df_master = pd.read_excel(xlsx_path)
        df_master['NR PM'] = df_master['NR PM'].apply(normalize_pm)
    except Exception as e:
        return None, f"Erro ao ler excel master (caminho: {xlsx_path}): {str(e)}"

    if not os.path.exists(csv_path):
        return None, f"Arquivo geral.csv nao encontrado no caminho: {csv_path}."

    cols_to_keep = {
        'nrPM (Avaliado)': 1, 'Nome Completo (Avaliado)': 2, 'Conceito Geral': 46, 'Nota Geral': 47,
        'Nota (Competência 1)': 50, 'Nota (Competência 2)': 53, 'Nota (Competência 3)': 56, 'Nota (Competência 4)': 59,
        'Nota da Homologação': 70, 'Data da Homologação': 71,
        'Recurso Fase 2': 77, 'Nota (Fase 2)': 78, 'Recurso Fase 3': 81, 'Nota (Fase 3)': 82,
        'Data da Avaliação 1': 36, 'Data da Avaliação 2': 45
    }
    
    rows = []
    try:
        with open(csv_path, "r", encoding="cp1252", errors="ignore") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            col_indices = {col: header.index(col) for col in cols_to_keep if col in header}
            for r in reader:
                row_clean = r[:len(header)]
                while len(row_clean) < len(header):
                    row_clean.append("")
                extracted_row = {}
                for col_name, col_idx in col_indices.items():
                    extracted_row[col_name] = row_clean[col_idx]
                rows.append(extracted_row)
        df_geral = pd.DataFrame(rows)
        df_geral['nrPM (Avaliado)'] = df_geral['nrPM (Avaliado)'].apply(normalize_pm)
    except Exception as e:
        return None, f"Erro ao ler geral.csv (caminho: {csv_path}): {str(e)}"

    discrepancies = []

    # 1. Auditoria de Qtd de Avaliações
    master_counts = df_master.set_index('NR PM')['Qtd Avaliações'].to_dict()
    geral_counts = df_geral['nrPM (Avaliado)'].value_counts().to_dict()
    all_pms = set(master_counts.keys()).union(set(geral_counts.keys()))

    for pm in sorted(list(all_pms)):
        if pm in ('nan', '', 'None'): continue
        m_c = master_counts.get(pm, 0)
        g_c = geral_counts.get(pm, 0)
        if m_c != g_c:
            name = ""
            if pm in master_counts:
                name = df_master[df_master['NR PM'] == pm]['Nome Completo'].values[0]
            else:
                matching_rows = df_geral[df_geral['nrPM (Avaliado)'] == pm]
                if not matching_rows.empty:
                    name = matching_rows['Nome Completo (Avaliado)'].values[0]
            discrepancies.append({
                "PM": pm,
                "Nome": name,
                "Tipo": "Divergência de Qtd de Avaliações",
                "Detalhe": f"Excel mestre diz {m_c} avaliações, mas geral.csv possui {g_c} registros."
            })

    # 2. Auditorias de Notas por registro de geral.csv
    for idx, row in df_geral.iterrows():
        pm = row['nrPM (Avaliado)']
        name = row['Nome Completo (Avaliado)']
        
        n_g = parse_float(row['Nota Geral'])
        c1 = parse_float(row['Nota (Competência 1)'])
        c2 = parse_float(row['Nota (Competência 2)'])
        c3 = parse_float(row['Nota (Competência 3)'])
        c4 = parse_float(row['Nota (Competência 4)'])
        
        if n_g is not None and all(x is not None for x in [c1, c2, c3, c4]):
            avg_comp = (c1 + c2 + c3 + c4) / 4.0
            if abs(avg_comp - n_g) > 0.01:
                discrepancies.append({
                    "PM": pm,
                    "Nome": name,
                    "Tipo": "Divergência de Média de Competências",
                    "Detalhe": f"Média das Competências = {avg_comp:.2f} (C1={c1}, C2={c2}, C3={c3}, C4={c4}) vs Nota Geral informada = {n_g}"
                })
                
        concept = row['Conceito Geral']
        n_hom = parse_float(row['Nota da Homologação'])
        
        if n_g is not None and not pd.isna(concept) and concept != '-':
            concept_norm = normaliza(str(concept))
            faixa = CONCEITO_FAIXA.get(concept_norm)
            if faixa:
                is_divergent = not (faixa[0] <= n_g <= faixa[1])
                if is_divergent:
                    if n_hom is None:
                        discrepancies.append({
                            "PM": pm,
                            "Nome": name,
                            "Tipo": "Divergência de Nota de Homologação",
                            "Detalhe": f"Divergência entre Conceito Geral ('{concept}') e Nota Geral ({n_g}), mas sem Nota de Homologação cadastrada."
                        })

    return pd.DataFrame(discrepancies), None




def build_audit_data_from_geral(csv_path):
    import csv
    import os
    import pandas as pd
    import tempfile
    import numpy as np
    import math
    import unicodedata

    com_map = {}
    try:
        cache_dir = os.path.join(tempfile.gettempdir(), "aadp_drive_cache")
        sirh_paths = [
            os.path.join(cache_dir, "COM_AADP_2026.xlsx"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "COM_AADP_2026.xlsx")
        ]
        sirh_file = next((p for p in sirh_paths if os.path.exists(p)), None)
        if sirh_file:
            df_com = pd.read_excel(sirh_file, dtype=str)
            df_com.columns = [str(c).strip() for c in df_com.columns]
            for _, row in df_com.iterrows():
                nrpm = str(row.get("MATRICULA", "")).strip().split('-')[0].lstrip("0") or "0"
                if not nrpm or nrpm == "0": continue
                n_sirh = str(row.get("NOTA DA AADP", "-")).strip()
                if n_sirh in ("nan", "None", "", "0.00", "0,00", "0"):
                    n_sirh = "-"
                o_sirh = str(row.get("DESCRICAO MOTIVO AADP", "")).strip()
                if o_sirh in ("nan", "None"):
                    o_sirh = ""
                com_map[nrpm] = {"nota": n_sirh, "obs": o_sirh}
    except Exception:
        pass

    SITUACOES_ALVO = {
        "ATIV. DIRECAO GERAL", "ATIV. FIM DESTACADO", "ATIV. FIM NA SEDE",
        "ATIV. MEIO", "ATIVIDADE MEIO", "DISP MED DEFINITIVA", "QUADRO ESPECIALISTA"
    }

    CONCEITO_FAIXA = {
        "nivel superior de desempenho":       (9.00, 10.00),
        "nivel alto de desempenho":           (7.00,  8.99),
        "nivel intermediario de desempenho":  (6.00,  6.99),
        "nivel baixo de desempenho":          (3.00,  5.99),
        "nivel inferior de desempenho":       (0.00,  2.99),
    }

    def normaliza(texto: str) -> str:
        if not isinstance(texto, str): return ""
        t = unicodedata.normalize("NFD", texto.lower())
        return "".join(c for c in t if unicodedata.category(c) != "Mn")

    def is_empty(v) -> bool:
        if v == 0 or v == 0.0:
            return False
        return not v or str(v).strip() in ("", "-", "nan", "none", "None", "<NA>")

    def parse_float(s):
        try:
            if is_empty(s): return None
            return float(str(s).replace(",", "."))
        except ValueError:
            return None

    def concordam(conceito: str, nota_str: str):
        if is_empty(conceito) or is_empty(nota_str):
            return None
        nota = parse_float(nota_str)
        if nota is None:
            return None
        faixa = CONCEITO_FAIXA.get(normaliza(conceito.strip()))
        if faixa is None:
            return None
        return faixa[0] <= nota <= faixa[1]

    def calc_status(j: str, l: str, n: str) -> str:
        if is_empty(j):
            return "Aberta"
        if is_empty(l):
            return "Parcialmente Encerrada"
        c = concordam(j, l)
        if c is True:
            return "Encerrada"
        elif c is False:
            return "Encerrada" if not is_empty(n) else "Homologação"
        return "Encerrada"

    def normalize_pm(pm):
        try:
            if is_empty(pm): return ""
            return str(int(float(str(pm).strip())))
        except Exception:
            return str(pm).strip()

    def find_col(header, pattern):
        pattern_norm = normaliza(pattern)
        for idx, col in enumerate(header):
            if pattern_norm in normaliza(col):
                return idx
        raise ValueError(f"Coluna contendo '{pattern}' não encontrada no geral.csv.")

    def parse_date(d_str):
        if is_empty(d_str): return None
        for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(str(d_str).strip(), fmt).date()
            except Exception:
                continue
        return None

    def add_business_days(start_date, num_days):
        curr = start_date
        added = 0
        while added < num_days:
            curr += timedelta(days=1)
            if curr.weekday() < 5: # Mon-Fri
                added += 1
        return curr

    pm_evals = {}
    
    with open(csv_path, "r", encoding="cp1252", errors="ignore") as f:
        reader = csv.reader(f, delimiter=";")
        header = next(reader)
        
        c_pm = find_col(header, "nrPM (Avaliado)")
        c_name = find_col(header, "Nome Completo (Avaliado)")
        c_rank = find_col(header, "Posto/Grad")
        c_rpm = find_col(header, "Unidade RPM (Avaliado)")
        c_unit = find_col(header, "Unidade Principal (Avaliado)")
        c_quadro = find_col(header, "Quadro Atual (Avaliado)")
        c_sit = find_col(header, "Situação Funcional")
        c_dt_av1 = find_col(header, "Data da Avaliação 1")
        c_dt_av2 = find_col(header, "Data da Avaliação 2")
        c_concept = find_col(header, "Conceito Geral")
        c_grade = find_col(header, "Nota Geral")
        c_n_hom = find_col(header, "Nota da Homologação")
        c_dt_hom = find_col(header, "Data da Homologação")
        
        c_n_f1 = find_col(header, "Nota (Fase 1)")
        c_n_f2 = find_col(header, "Nota (Fase 2)")
        c_n_f3 = find_col(header, "Nota (Fase 3)")
        c_n_f4 = find_col(header, "Nota (Fase 4)")
        c_r_f1 = find_col(header, "Recurso Fase 1")
        c_r_f2 = find_col(header, "Recurso Fase 2")
        c_r_f3 = find_col(header, "Recurso Fase 3")
        c_r_f4 = find_col(header, "Recurso Fase 4")
        c_dt_f1 = find_col(header, "Data Cadastro (Fase 1)")
        c_dt_f2 = find_col(header, "Data Cadastro (Fase 2)")
        c_dt_f3 = find_col(header, "Data Cadastro (Fase 3)")
        c_dt_f4 = find_col(header, "Data Cadastro (Fase 4)")
        
        for row in reader:
            if len(row) > 38 and row[38] == "Disciplina":
                row = row[:18] + [""] * 10 + row[18:]
            while len(row) < len(header):
                row.append("")
                
            sit = row[c_sit].strip()
            pm = normalize_pm(row[c_pm])
            if not pm:
                continue
                
            j = row[c_concept].strip()
            l = row[c_grade].strip()
            n = row[c_n_hom].strip()
            status = calc_status(j, l, n)
            c = concordam(j, l)
            
            n_f4 = parse_float(row[c_n_f4])
            n_f3 = parse_float(row[c_n_f3])
            n_f2 = parse_float(row[c_n_f2])
            n_f1 = parse_float(row[c_n_f1])
            
            r_f4 = row[c_r_f4].strip()
            r_f3 = row[c_r_f3].strip()
            r_f2 = row[c_r_f2].strip()
            r_f1 = row[c_r_f1].strip()
            
            dt_f4 = row[c_dt_f4].strip()
            dt_f3 = row[c_dt_f3].strip()
            dt_f2 = row[c_dt_f2].strip()
            dt_f1 = row[c_dt_f1].strip()
            
            original_grade = parse_float(n) if not is_empty(n) else parse_float(l)
            has_appeal = (r_f1 not in ("", "-")) or (n_f1 is not None)
            ref_date = now_br().date()
            
            final_grade = None
            houve_recurso = "-"
            fase_recurso = "-"
            nota_recurso = "-"
            
            if status in ("Aberta", "Parcialmente Encerrada"):
                final_grade = None
            elif not has_appeal:
                if c is False:
                    # Houve discordância: necessita passar para o homologador
                    if is_empty(n):
                        status = "Homologação"
                        final_grade = None
                    else:
                        dt_base = parse_date(row[c_dt_hom])
                        if dt_base is not None:
                            deadline = add_business_days(dt_base, 5)
                            if ref_date <= deadline:
                                status = "EM PRAZO DE RECURSO"
                                final_grade = None
                            else:
                                status = "Encerrada"
                                final_grade = parse_float(n)
                        else:
                            status = "Encerrada"
                            final_grade = parse_float(n)
                else:
                    # Não houve discordância: prazo de 5 dias úteis a partir da data de AV2
                    dt_base = parse_date(row[c_dt_av2])
                    if dt_base is not None:
                        deadline = add_business_days(dt_base, 5)
                        if ref_date <= deadline:
                            status = "EM PRAZO DE RECURSO"
                            final_grade = None
                        else:
                            status = "Encerrada"
                            final_grade = parse_float(l)
                    else:
                        status = "Encerrada"
                        final_grade = parse_float(l)
            else:
                # Houve recurso! (r_f1 registrado = militar interpôs recurso)
                houve_recurso = "SIM"

                if n_f4 is not None:
                    # Fase 4 (encerra definitivamente)
                    final_grade = n_f4
                    fase_recurso = "FASE 4"
                    nota_recurso = str(n_f4)
                    status = "Encerrada"

                elif r_f3 not in ("", "-") or n_f3 is not None:
                    # Fase 3 registrada = Autoridade Recursal DECIDIU → Encerrada
                    # (com nota → nova nota; sem nota + com data → indeferido c/ nota original)
                    fase_recurso = "FASE 3"
                    if n_f3 is not None:
                        final_grade = n_f3
                        nota_recurso = str(n_f3)
                    else:
                        final_grade = original_grade   # indeferido: mantém nota da comissão
                        nota_recurso = "-"
                    status = "Encerrada"

                elif r_f2 not in ("", "-"):
                    # Fase 2 registrada = comissão promoveu para Autoridade Recursal
                    fase_recurso = "FASE 2"
                    if n_f2 is not None:
                        # Autoridade deferiu com nova nota (sem Fase 3 explícita)
                        final_grade = n_f2
                        nota_recurso = str(n_f2)
                        status = "Encerrada"
                    else:
                        # Autoridade ainda não decidiu
                        final_grade = None
                        nota_recurso = "-"
                        status = "AUTORIDADE RECURSAL"

                else:
                    # Só Fase 1 registrada = comissão analisando (Reconsideração)
                    fase_recurso = "FASE 1"
                    if n_f1 is not None:
                        # Comissão deferiu com nova nota
                        final_grade = n_f1
                        nota_recurso = str(n_f1)
                        status = "Encerrada"
                    else:
                        # Comissão ainda analisando
                        final_grade = None
                        nota_recurso = "-"
                        status = "RECONSIDERAÇÃO COMISSÃO"
                
            dt_av = row[c_dt_hom].strip() if not is_empty(n) and row[c_dt_hom].strip() else (row[c_dt_av2].strip() or row[c_dt_av1].strip() or "-")
            
            eval_data = {
                "date": dt_av,
                "status": status.upper(),
                "grade": parse_float(n) if not is_empty(n) else (parse_float(l) if not is_empty(l) else "-"),
                "houve_recurso": houve_recurso,
                "fase_recurso": fase_recurso,
                "nota_recurso": nota_recurso,
                "final_grade": final_grade
            }
            
            if pm not in pm_evals:
                c_data = com_map.get(pm, {})
                pm_evals[pm] = {
                    "NR PM": pm,
                    "Posto/Graduação": row[c_rank].strip(),
                    "Nome Completo": row[c_name].strip(),
                    "Nome RPM": row[c_rpm].strip(),
                    "Nome Unidade Principal": row[c_unit].strip(),
                    "Quadro": row[c_quadro].strip(),
                    "Sit. Funcional": sit,
                    "Nota SIRH": c_data.get("nota", "-"),
                    "Observação": c_data.get("obs", ""),
                    "evals": []
                }
            pm_evals[pm]["evals"].append(eval_data)
            
    rows_audit = []
    for pm, data in pm_evals.items():
        evals = data["evals"]
        qtd = len(evals)
        obs_str = str(data.get("Observação", "")).strip()
        obs_lower = obs_str.lower()
        motivo = "-"
        if "artigo 17" in obs_lower or "art 17" in obs_lower or "art. 17" in obs_lower or "art.17" in obs_lower:
            motivo = "ARTIGO 17"
        elif "artigo 20" in obs_lower or "art 20" in obs_lower or "art. 20" in obs_lower or "art.20" in obs_lower or "revis" in obs_lower:
            motivo = "ARTIGO 20"
        elif obs_str not in ("", "-", "nan", "None"):
            motivo = obs_str
            
        has_motivo = (motivo != "-")
        
        todas_encerradas = "SIM" if (all(e["status"] == "ENCERRADA" for e in evals) or has_motivo) else "NAO"

        ignore_sits = [
            "RESER.NAO REMUNERADA", "RES. TEMPO SERVICO", "EXCLUIDO",
            "RES.TEMPO EFET.SERV.", "REFORMA INCAP.FISICA", 
            "REFORMA P/ INVALIDEZ", "REF.LIM.IDAD.QOR/QPR"
        ]
        if str(data.get("Sit. Funcional", "")).strip().upper() in ignore_sits:
            if not has_motivo and all(e.get("status", "") == "ABERTA" for e in evals):
                continue

        if todas_encerradas == "SIM":
            grades = [e["final_grade"] for e in evals if e["final_grade"] is not None]
            if len(grades) == qtd and qtd > 0:
                final_avg = sum(grades) / float(qtd)
                final_avg_rounded = math.floor(final_avg * 100 + 0.5) / 100.0
            else:
                final_avg_rounded = "-"
        else:
            final_avg_rounded = "-"
            
        if final_avg_rounded != "-":
            try:
                val_geral = float(str(final_avg_rounded).replace(",", "."))
            except:
                val_geral = None
        else:
            val_geral = None
            
        nota_sirh = data["Nota SIRH"]
        try:
            val_sirh = float(str(nota_sirh).replace(",", "."))
        except:
            val_sirh = None
            
        if val_geral is not None and val_sirh is not None and abs(val_geral - val_sirh) < 0.01:
            auditoria = "IGUAL"
        elif nota_sirh == "-" or final_avg_rounded == "-":
            auditoria = ""
        else:
            auditoria = "DIVERGENTE"
            
        if has_motivo and auditoria in ("DIVERGENTE", ""):
            auditoria = motivo
            
        r_audit = {
            "NR PM": data["NR PM"],
            "Posto/Graduação": data["Posto/Graduação"],
            "Nome Completo": data["Nome Completo"],
            "Nome RPM": data["Nome RPM"],
            "Nome Unidade Principal": data["Nome Unidade Principal"],
            "Quadro": data["Quadro"],
            "Sit. Funcional": data["Sit. Funcional"],
            "Qtd Avaliações": qtd,
            "Todas Avaliações Foram Encerradas?": todas_encerradas,
            "Nota Final - Média Aritmética": final_avg_rounded,
            "Nota SIRH": nota_sirh,
            "Auditoria": auditoria,
        }
        
        for i in range(1, 5):
            if i <= qtd:
                ev = evals[i-1]
                r_audit[f"Data Avaliação {i}"] = ev["date"]
                r_audit[f"Fase Avaliação {i}"] = ev["status"]
                r_audit[f"Nota Avaliação {i}"] = ev["grade"]
                r_audit[f"Houve Recurso? {i}"] = ev["houve_recurso"]
                r_audit[f"Fase Recurso {i}"] = ev["fase_recurso"]
                r_audit[f"Nota Fase 2 ou 3 {i}"] = ev["nota_recurso"]
            else:
                r_audit[f"Data Avaliação {i}"] = np.nan
                r_audit[f"Fase Avaliação {i}"] = np.nan
                r_audit[f"Nota Avaliação {i}"] = np.nan
                r_audit[f"Houve Recurso? {i}"] = np.nan
                r_audit[f"Fase Recurso {i}"] = np.nan
                r_audit[f"Nota Fase 2 ou 3 {i}"] = np.nan
        
        r_audit["Observação"] = data["Observação"]
        
        r_audit["Motivo"] = motivo

        rows_audit.append(r_audit)
        
    missing_pms = set(com_map.keys()) - set(pm_evals.keys())
    if missing_pms:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        cache_dir = os.path.join(tempfile.gettempdir(), "aadp_drive_cache")
        possible_paths = [
            os.path.join(base_dir, "dados", "SIGEF.csv"),
            os.path.join(base_dir, "SIGEF.csv"),
            os.path.join(cache_dir, "SIGEF.csv")
        ]
        si_path = next((p for p in possible_paths if os.path.exists(p) and os.path.getsize(p) > 0), None)
        
        if si_path:
            with open(si_path, encoding="cp1252", errors="replace") as f_si:
                reader_si = csv.reader(f_si, delimiter=";")
                header_si = next(reader_si)
                for row_si in reader_si:
                    if len(row_si) > 20:
                        pm_si = str(row_si[0]).strip().lstrip("0")
                        if pm_si in missing_pms:
                            idx_offset = 0
                            if len(row_si) > 16 and row_si[15] in ["A", "I"]:
                                idx_offset = -1
                            elif len(row_si) > 17 and row_si[16] in ["A", "I"]:
                                idx_offset = 0
                            elif len(row_si) > 18 and row_si[17] in ["A", "I"]:
                                idx_offset = 1
                                
                            try:
                                posto = row_si[2 + idx_offset].strip()
                                nome = row_si[3 + idx_offset].strip()
                                rpm = row_si[5 + idx_offset].strip()
                                unid = row_si[7 + idx_offset].strip()
                                quadro = row_si[14 + idx_offset].strip()
                                sit_func = row_si[27 + idx_offset].strip()
                            except IndexError:
                                continue

                            c_data = com_map.get(pm_si, {})
                            nota_sirh = c_data.get("nota", "-")

                            obs_str = str(c_data.get("obs", "")).strip()
                            obs_lower = obs_str.lower()
                            motivo_si = "-"
                            if "artigo 17" in obs_lower or "art 17" in obs_lower or "art. 17" in obs_lower or "art.17" in obs_lower:
                                motivo_si = "ARTIGO 17"
                            elif "artigo 20" in obs_lower or "art 20" in obs_lower or "art. 20" in obs_lower or "art.20" in obs_lower or "revis" in obs_lower:
                                motivo_si = "ARTIGO 20"
                            elif obs_str not in ("", "-", "nan", "None"):
                                motivo_si = obs_str

                            has_motivo_si = (motivo_si != "-")

                            ignore_sits = [
                                "RESER.NAO REMUNERADA", "RES. TEMPO SERVICO", "EXCLUIDO",
                                "RES.TEMPO EFET.SERV.", "REFORMA INCAP.FISICA", 
                                "REFORMA P/ INVALIDEZ", "REF.LIM.IDAD.QOR/QPR"
                            ]
                            if str(sit_func).strip().upper() in ignore_sits:
                                if not has_motivo_si:
                                    missing_pms.remove(pm_si)
                                    if not missing_pms:
                                        break
                                    continue

                            
                            auditoria_val = "DIVERGENTE" if (nota_sirh != "-" and nota_sirh != "") else ""
                            if has_motivo_si and auditoria_val in ("DIVERGENTE", ""):
                                auditoria_val = motivo_si
                                
                            r_audit = {
                                "NR PM": pm_si,
                                "Posto/Graduação": posto,
                                "Nome Completo": nome,
                                "Nome RPM": rpm,
                                "Nome Unidade Principal": unid,
                                "Quadro": quadro,
                                "Sit. Funcional": sit_func,
                                "Qtd Avaliações": 0,
                                "Todas Avaliações Foram Encerradas?": "SIM" if has_motivo_si else "NAO",
                                "Nota Final - Média Aritmética": "-",
                                "Nota SIRH": nota_sirh,
                                "Auditoria": auditoria_val,
                            }
                            
                            for i in range(1, 5):
                                r_audit[f"Data Avaliação {i}"] = np.nan
                                r_audit[f"Fase Avaliação {i}"] = np.nan
                                r_audit[f"Nota Avaliação {i}"] = np.nan
                                r_audit[f"Houve Recurso? {i}"] = np.nan
                                r_audit[f"Fase Recurso {i}"] = np.nan
                                r_audit[f"Nota Fase 2 ou 3 {i}"] = np.nan

                            r_audit["Observação"] = c_data.get("obs", "")
                            r_audit["Motivo"] = motivo_si

                            rows_audit.append(r_audit)
                            missing_pms.remove(pm_si)
                            if not missing_pms:
                                break

    return pd.DataFrame(rows_audit)


def append_sigef_to_audit(df):
    import os, csv
    import pandas as pd
    import tempfile
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    cache_dir = os.path.join(tempfile.gettempdir(), "aadp_drive_cache")
    
    possible_paths = [
        os.path.join(base_dir, "dados", "SIGEF.csv"),
        os.path.join(base_dir, "SIGEF.csv"),
        os.path.join(cache_dir, "SIGEF.csv")
    ]
    
    si_path = next((p for p in possible_paths if os.path.exists(p) and os.path.getsize(p) > 0), None)
        
    if not si_path:
        return df  # If SIGEF not found, return original df
        
    sigef_map = {}
    try:
        with open(si_path, encoding="cp1252", errors="replace") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            
            try:
                idx_quinq = header.index("NUM. QUINQUENIOS")
            except ValueError:
                idx_quinq = 20
            try:
                idx_ade = header.index("DATA ADE")
            except ValueError:
                idx_ade = 24
            try:
                idx_ano_base = header.index("ANO BASE")
            except ValueError:
                idx_ano_base = 32
            try:
                idx_ord_alm = header.index("ORD. ALMANAQUE")
            except ValueError:
                idx_ord_alm = 33
            try:
                idx_conceito = header.index("CONCEITO")
            except ValueError:
                idx_conceito = 34
            try:
                idx_sinal = header.index("SINAL")
            except ValueError:
                idx_sinal = 35
            try:
                idx_pontuacao = header.index("PONTUACAO")
            except ValueError:
                idx_pontuacao = 36

            for row in reader:
                if len(row) > 20:
                    pm_clean = row[0].strip().lstrip("0")
                    if not pm_clean: continue
                    
                    idx_offset = 0
                    if len(row) > 16 and row[15] in ["A", "I"]:
                        idx_offset = -1
                    elif len(row) > 17 and row[16] in ["A", "I"]:
                        idx_offset = 0
                    elif len(row) > 18 and row[17] in ["A", "I"]:
                        idx_offset = 1
                        
                    try:
                        num_quinquenio = row[idx_quinq + idx_offset].strip() if len(row) > idx_quinq + idx_offset else ""
                        data_ade = row[idx_ade + idx_offset].strip() if len(row) > idx_ade + idx_offset else ""
                        ano_base = row[idx_ano_base + idx_offset].strip() if len(row) > idx_ano_base + idx_offset else ""
                        ord_almanaque = row[idx_ord_alm + idx_offset].strip() if len(row) > idx_ord_alm + idx_offset else ""
                        ah = row[idx_conceito + idx_offset].strip() if len(row) > idx_conceito + idx_offset else ""
                        ai = row[idx_sinal + idx_offset].strip() if len(row) > idx_sinal + idx_offset else ""
                        aj = row[idx_pontuacao + idx_offset].strip() if len(row) > idx_pontuacao + idx_offset else ""
                    except IndexError:
                        continue
                    
                    # Reg Adicional: Se X (Data ADE) tem dado válido -> ADE. Se não, T -> QQ. Se ambos -> ADE.
                    reg_adicional = ""
                    if data_ade and str(data_ade).strip() not in ("-", "", "nan", "None"):
                        reg_adicional = "ADE"
                    elif num_quinquenio and str(num_quinquenio).strip() not in ("-", "", "nan", "None", "0"):
                        reg_adicional = "QQ"
                        
                    # Conceito
                    conceito = f"{ah} {ai}{aj}".strip()
                    
                    sigef_map[pm_clean] = {
                        "Conceito": conceito,
                        "Reg. Adicional": reg_adicional,
                        "Ano Base": ano_base,
                        "Ord. Almanaque": ord_almanaque
                    }
    except Exception:
        pass
        
    if not sigef_map:
        return df
        
    col_pm = "NR PM" if "NR PM" in df.columns else "nrPM (Avaliado)"
    if col_pm in df.columns:
        df["Conceito"] = df[col_pm].apply(lambda x: sigef_map.get(str(x).strip().lstrip("0"), {}).get("Conceito", "N/A") if pd.notnull(x) else "N/A")
        df["Reg. Adicional"] = df[col_pm].apply(lambda x: sigef_map.get(str(x).strip().lstrip("0"), {}).get("Reg. Adicional", "N/A") if pd.notnull(x) else "N/A")
        df["Ano Base"] = df[col_pm].apply(lambda x: sigef_map.get(str(x).strip().lstrip("0"), {}).get("Ano Base", "N/A") if pd.notnull(x) else "N/A")
        df["Ord. Almanaque"] = df[col_pm].apply(lambda x: sigef_map.get(str(x).strip().lstrip("0"), {}).get("Ord. Almanaque", "N/A") if pd.notnull(x) else "N/A")
    
    # Reorder columns
    cols = list(df.columns)
    
    # Find "Sit. Funcional"
    if "Sit. Funcional" in cols:
        idx = cols.index("Sit. Funcional") + 1
        
        # Remove the newly added cols from their current end position
        for c in ["Conceito", "Reg. Adicional", "Ano Base", "Ord. Almanaque"]:
            if c in cols:
                cols.remove(c)
                
        # Insert them right after Situação Funcional
        cols.insert(idx, "Conceito")
        cols.insert(idx + 1, "Reg. Adicional")
        cols.insert(idx + 2, "Ano Base")
        cols.insert(idx + 3, "Ord. Almanaque")
        
        df = df[cols]
        
    return df

def load_audit_excel(xlsx_path, drive_master_xlsx_id=None, ano="2026"):
    import pandas as pd
    import os
    import tempfile
    from pathlib import Path
    
    cfg_to_use = load_config()
    y_cfg = get_active_year_config(ano, cfg_to_use)
    drive_geral_id = y_cfg.get("drive_geral_id", "")
    drive_com_id = y_cfg.get("drive_com_id", "")
    drive_sirh_id = y_cfg.get("drive_sirh_id", "")
    if not drive_master_xlsx_id:
        drive_master_xlsx_id = y_cfg.get("drive_master_xlsx_id", "")
        
    cache_dir = os.path.join(tempfile.gettempdir(), f"aadp_drive_cache_{ano}")
    os.makedirs(cache_dir, exist_ok=True)
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # 1. Procura se a planilha consolidada já existe localmente
    possible_master = [
        xlsx_path,
        os.path.join(base_dir, f"DADOS AADP {ano}", "Analise avaliacoes completa.xlsx"),
        os.path.join(y_cfg.get("db_path", ""), "Analise avaliacoes completa.xlsx"),
        os.path.join(cache_dir, "Analise avaliacoes completa.xlsx"),
        os.path.join(base_dir, "Analise avaliacoes completa.xlsx")
    ]
    resolved_xlsx = next((p for p in possible_master if p and os.path.exists(p) and os.path.getsize(p) > 0), None)

    # 2. Se não encontrou localmente e possui ID no Drive, baixa a planilha consolidada
    if not resolved_xlsx and drive_master_xlsx_id:
        target_dest = os.path.join(cache_dir, "Analise avaliacoes completa.xlsx")
        try:
            _baixar_drive(drive_master_xlsx_id, target_dest)
            if os.path.exists(target_dest) and os.path.getsize(target_dest) > 0:
                resolved_xlsx = target_dest
        except Exception as e:
            pass

    # 3. Se a planilha consolidada existe, carrega diretamente
    if resolved_xlsx and os.path.exists(resolved_xlsx):
        try:
            df = pd.read_excel(resolved_xlsx)
            col_media = next((c for c in df.columns if "Aritm" in str(c)), None)
            if col_media:
                import math
                def round_half_up_2(x):
                    try:
                        if pd.isna(x) or x is None: return x
                        val = float(str(x).replace(",", "."))
                        return math.floor(val * 100 + 0.5) / 100.0
                    except Exception:
                        return x
                df[col_media] = df[col_media].apply(round_half_up_2)
            return df, None
        except Exception:
            pass

    # 4. Fallback: se não houver planilha consolidada, tenta gerar a partir de geral.csv
    drive_geral_path = os.path.join(cache_dir, "geral.csv")
    local_geral_paths = [
        os.path.join(base_dir, f"DADOS AADP {ano}", "geral.csv"),
        os.path.join(y_cfg.get("db_path", ""), "geral.csv"),
        os.path.join(base_dir, "geral.csv")
    ]
    local_geral_path = next((p for p in local_geral_paths if os.path.exists(p) and os.path.getsize(p) > 0), None)
    
    if drive_com_id:
        try:
            com_path = os.path.join(cache_dir, "comissao.csv")
            if not os.path.exists(com_path) or os.path.getsize(com_path) == 0:
                _baixar_drive(drive_com_id, com_path)
        except Exception:
            pass

    if drive_sirh_id:
        try:
            sirh_path = os.path.join(cache_dir, f"COM_AADP_{ano}.xlsx")
            if not os.path.exists(sirh_path) or os.path.getsize(sirh_path) == 0:
                _baixar_drive(drive_sirh_id, sirh_path)
        except Exception:
            pass
    
    csv_to_use = None
    if drive_geral_id:
        try:
            if not os.path.exists(drive_geral_path) or os.path.getsize(drive_geral_path) == 0:
                _baixar_drive(drive_geral_id, drive_geral_path)
            if os.path.exists(drive_geral_path) and os.path.getsize(drive_geral_path) > 0:
                csv_to_use = drive_geral_path
        except Exception:
            pass
            
    if not csv_to_use and local_geral_path:
        csv_to_use = local_geral_path
        
    if csv_to_use:
        try:
            df = build_audit_data_from_geral(csv_to_use)
            df = append_sigef_to_audit(df)
            return df, None
        except Exception:
            pass

    return None, f"Arquivo 'Analise avaliacoes completa.xlsx' ou 'geral.csv' não encontrado para AADP {ano}."

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
<script>
(function() {
    let changed = false;
    const key = "streamlit:themeConfiguration";
    const expected = '{"themePreset":"dark"}';
    const lightTheme = '{"themePreset":"light"}';
    
    const forceDark = (win) => {
        try {
            const current = win.localStorage.getItem(key);
            if (current !== lightTheme) {
                if (current !== expected) {
                    win.localStorage.setItem(key, expected);
                    changed = true;
                }
                const legacyKeys = ["stActiveTheme-light", "stActiveTheme-dark", "stActiveTheme", "stActiveThemeType", "stTheme", "theme"];
                legacyKeys.forEach(k => {
                    if (win.localStorage.getItem(k)) {
                        win.localStorage.removeItem(k);
                        changed = true;
                    }
                });
            }
        } catch(e) {}
    };

    forceDark(window);
    if (window.parent) {
        forceDark(window.parent);
    }
    if (changed) {
        window.location.reload();
    }
})();
</script>

<style>


@import url('https://fonts.googleapis.com/css2?family=Inter:wght=300;400;500;600;700&display=swap');


html,body,[class*="css"]{font-family:'Inter',sans-serif;}





[data-testid="stSidebar"]{


  background:linear-gradient(180deg,#0c0b07 0%,#1c1c1c 100%);


  border-right:1px solid #9b8a5c;


}


[data-testid="stSidebar"] *{color:#e5dccb!important;}


[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,


[data-testid="stSidebar"] h3{color:#9b8a5c!important;}


.main{background:#121212;}





/* Garante que a seta de recolhimento e expansão da barra lateral esteja sempre visível */


[data-testid="collapsedControl"], button[data-testid="stSidebarCollapseButton"] {


  color: #9b8a5c !important;


  display: flex !important;


  visibility: visible !important;


  opacity: 1 !important;


}





.main-title{


  background:linear-gradient(135deg,#0c0b07 0%,#282828 100%);


  color:#9b8a5c;padding:24px 28px;border-radius:12px;margin-bottom:20px;


  display:flex;flex-direction:column;align-items:center;text-align:center;gap:16px;


  box-shadow:0 4px 20px rgba(0,0,0,.15);


  border-bottom:3px solid #9b8a5c;


}


.main-title h1{margin:0;font-size:2.1rem;font-weight:800;color:#9b8a5c;letter-spacing:0.02em;}


.main-title p{margin:0;font-size:.9rem;opacity:.9;color:#e5dccb;margin-top:6px;}





.kpi-card{background:#1c1c1c;border-radius:12px;padding:16px 20px;


  box-shadow:0 4px 15px rgba(0,0,0,0.3);border-left:5px solid;


  transition:transform .2s,box-shadow .2s;text-align:center;}


.kpi-card:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,0,0,0.4);}


.kpi-card .label{font-size:.7rem;font-weight:700;text-transform:uppercase;


  letter-spacing:.06em;color:#a0a0a0!important;margin-bottom:6px;}


.kpi-card .value{font-size:1.9rem;font-weight:800;line-height:1;}


.kpi-card .sub{font-size:.72rem;color:#777777!important;margin-top:4px;}


.kpi-total    {border-color:#9b8a5c;} .kpi-total    .value{color:#e5dccb!important;}


.kpi-ca       {border-color:#9b8a5c;} .kpi-ca       .value{color:#e5dccb!important;}


.kpi-np       {border-color:#8c6e42;} .kpi-np       .value{color:#9b8a5c!important;}


.kpi-aberta   {border-color:#FF4444;} .kpi-aberta   .value{color:#ff6b6b!important;}


.kpi-parc     {border-color:#FF8C00;} .kpi-parc     .value{color:#ff9f43!important;}


.kpi-hom      {border-color:#FFD966;} .kpi-hom      .value{color:#ffd257!important;}


.kpi-enc      {border-color:#70AD47;} .kpi-enc      .value{color:#7bed9f!important;}
.kpi-recurso  {border-color:#9B59B6;} .kpi-recurso  .value{color:#d6a2e8!important;}
.kpi-active-total  { background: #282828 !important; box-shadow: 0 0 15px rgba(155, 138, 92, 0.45) !important; border: 1.5px solid #9b8a5c !important; border-left: 5px solid #9b8a5c !important; }
.kpi-active-ca     { background: #282828 !important; box-shadow: 0 0 15px rgba(155, 138, 92, 0.45) !important; border: 1.5px solid #9b8a5c !important; border-left: 5px solid #9b8a5c !important; }
.kpi-active-np     { background: #282828 !important; box-shadow: 0 0 15px rgba(140, 110, 66, 0.45) !important; border: 1.5px solid #8c6e42 !important; border-left: 5px solid #8c6e42 !important; }
.kpi-active-enc    { background: #282828 !important; box-shadow: 0 0 15px rgba(112, 173, 71, 0.45) !important; border: 1.5px solid #70AD47 !important; border-left: 5px solid #70AD47 !important; }
.kpi-active-aberta { background: #282828 !important; box-shadow: 0 0 15px rgba(255, 68, 68, 0.45) !important; border: 1.5px solid #FF4444 !important; border-left: 5px solid #FF4444 !important; }
.kpi-active-parc   { background: #282828 !important; box-shadow: 0 0 15px rgba(255, 140, 0, 0.45) !important; border: 1.5px solid #FF8C00 !important; border-left: 5px solid #FF8C00 !important; }
.kpi-active-hom    { background: #282828 !important; box-shadow: 0 0 15px rgba(255, 217, 102, 0.45) !important; border: 1.5px solid #FFD966 !important; border-left: 5px solid #FFD966 !important; }
.kpi-active-recurso { background: #282828 !important; box-shadow: 0 0 15px rgba(155, 89, 182, 0.45) !important; border: 1.5px solid #9B59B6 !important; border-left: 5px solid #9B59B6 !important; }

/* Glassmorphic Crystal Style Button */
button[aria-label="👁️ Mostrar Encerradas"],
button[aria-label="🙈 Ocultar Encerradas"] {
    background: rgba(155, 138, 92, 0.15) !important;
    backdrop-filter: blur(8px) !important;
    -webkit-backdrop-filter: blur(8px) !important;
    border: 1px solid rgba(155, 138, 92, 0.4) !important;
    border-radius: 30px !important;
    color: var(--text-color) !important;
    font-weight: 600 !important;
    box-shadow: 0 4px 20px rgba(155, 138, 92, 0.25), inset 0 1px 2px rgba(255, 255, 255, 0.1) !important;
    text-shadow: 0 1px 2px rgba(0, 0, 0, 0.3) !important;
    transition: all 0.3s ease !important;
}

button[aria-label="👁️ Mostrar Encerradas"]:hover,
button[aria-label="🙈 Ocultar Encerradas"]:hover {
    background: rgba(155, 138, 92, 0.25) !important;
    border-color: rgba(155, 138, 92, 0.6) !important;
    box-shadow: 0 6px 25px rgba(155, 138, 92, 0.4), inset 0 1px 3px rgba(255, 255, 255, 0.2) !important;
    transform: translateY(-1px) !important;
}

button[aria-label="👁️ Mostrar Encerradas"]:active,
button[aria-label="🙈 Ocultar Encerradas"]:active {
    background: rgba(128, 128, 128, 0.02) !important;
    transform: translateY(0px) !important;
}

/* Crystal Liquid styles for interactive legend buttons */
button[aria-label="🟢 Encerrada"],
button[aria-label="🔴 Aberta"],
button[aria-label="🟠 Parcialmente Encerrada"],
button[aria-label="🟡 Homologação"],
button[aria-label="⚪ Encerrada"],
button[aria-label="⚪ Aberta"],
button[aria-label="⚪ Parcialmente Encerrada"],
button[aria-label="⚪ Homologação"] {
    backdrop-filter: blur(8px) !important;
    -webkit-backdrop-filter: blur(8px) !important;
    border-radius: 30px !important;
    color: var(--text-color) !important;
    font-weight: 600 !important;
    text-shadow: 0 1px 2px rgba(0, 0, 0, 0.2) !important;
    transition: all 0.3s ease !important;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15) !important;
}

/* Hover effects */
button[aria-label="🟢 Encerrada"]:hover,
button[aria-label="🔴 Aberta"]:hover,
button[aria-label="🟠 Parcialmente Encerrada"]:hover,
button[aria-label="🟡 Homologação"]:hover,
button[aria-label="⚪ Encerrada"]:hover,
button[aria-label="⚪ Aberta"]:hover,
button[aria-label="⚪ Parcialmente Encerrada"]:hover,
button[aria-label="⚪ Homologação"]:hover {
    transform: translateY(-1px) !important;
}

/* Colored glowing styles for active buttons */
button[aria-label="🟢 Encerrada"] {
    background: rgba(112, 173, 71, 0.15) !important;
    border: 1px solid #70AD47 !important;
    box-shadow: 0 0 10px rgba(112, 173, 71, 0.3) !important;
}
button[aria-label="🔴 Aberta"] {
    background: rgba(255, 68, 68, 0.15) !important;
    border: 1px solid #FF4444 !important;
    box-shadow: 0 0 10px rgba(255, 68, 68, 0.3) !important;
}
button[aria-label="🟠 Parcialmente Encerrada"] {
    background: rgba(255, 140, 0, 0.15) !important;
    border: 1px solid #FF8C00 !important;
    box-shadow: 0 0 10px rgba(255, 140, 0, 0.3) !important;
}
button[aria-label="🟡 Homologação"] {
    background: rgba(255, 217, 102, 0.15) !important;
    border: 1px solid #FFD966 !important;
    box-shadow: 0 0 10px rgba(255, 217, 102, 0.3) !important;
}

/* Muted look for inactive buttons */
button[aria-label="⚪ Encerrada"],
button[aria-label="⚪ Aberta"],
button[aria-label="⚪ Parcialmente Encerrada"],
button[aria-label="⚪ Homologação"] {
    background: rgba(128, 128, 128, 0.03) !important;
    border: 1px solid rgba(128, 128, 128, 0.15) !important;
    opacity: 0.55 !important;
}

/* Glassmorphic Crystal Style for Page Navigation / Horizontal Tab Buttons */
div.element-container:has(.main-nav-marker) + div[data-testid="stHorizontalBlock"] button,
div.element-container:has(.main-nav-marker) + div.element-container button,
div.element-container:has(.main-nav-marker) + div button {
    background: rgba(128, 128, 128, 0.06) !important;
    backdrop-filter: blur(8px) !important;
    -webkit-backdrop-filter: blur(8px) !important;
    border: 1px solid rgba(128, 128, 128, 0.2) !important;
    border-radius: 8px !important;
    color: var(--text-color) !important;
    font-weight: 600 !important;
    padding: 6px 4px !important;
    transition: all 0.25s ease !important;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08) !important;
    
    /* Fixed unified dimensions and vertical centering */
    min-height: 85px !important;
    height: 85px !important;
    display: flex !important;
    flex-direction: column !important;
    justify-content: center !important;
    align-items: center !important;
}

/* Force inner text tags to wrap on pre-line and have uniform small font size */
div.element-container:has(.main-nav-marker) + div[data-testid="stHorizontalBlock"] button *,
div.element-container:has(.main-nav-marker) + div.element-container button *,
div.element-container:has(.main-nav-marker) + div button * {
    color: var(--text-color) !important;
    white-space: pre-line !important;
    text-align: center !important;
    font-size: 0.8rem !important;
    line-height: 1.25 !important;
}

/* Highlight the icon (first line / emoji) by making it significantly larger */
div.element-container:has(.main-nav-marker) + div[data-testid="stHorizontalBlock"] button p::first-line,
div.element-container:has(.main-nav-marker) + div.element-container button p::first-line,
div.element-container:has(.main-nav-marker) + div button p::first-line,
div.element-container:has(.main-nav-marker) + div[data-testid="stHorizontalBlock"] button span::first-line,
div.element-container:has(.main-nav-marker) + div.element-container button span::first-line,
div.element-container:has(.main-nav-marker) + div button span::first-line {
    font-size: 1.65rem !important;
    line-height: 1.45 !important;
    font-weight: normal !important;
}

div.element-container:has(.main-nav-marker) + div[data-testid="stHorizontalBlock"] button:hover,
div.element-container:has(.main-nav-marker) + div.element-container button:hover,
div.element-container:has(.main-nav-marker) + div button:hover {
    background: rgba(128, 128, 128, 0.12) !important;
    border-color: rgba(128, 128, 128, 0.3) !important;
    transform: translateY(-1px) !important;
}

div.element-container:has(.main-nav-marker) + div[data-testid="stHorizontalBlock"] button:hover *,
div.element-container:has(.main-nav-marker) + div.element-container button:hover *,
div.element-container:has(.main-nav-marker) + div button:hover * {
    color: var(--text-color) !important;
}

/* Make active page buttons glow gold */
div.element-container:has(.main-nav-marker) + div[data-testid="stHorizontalBlock"] button[kind="primary"],
div.element-container:has(.main-nav-marker) + div.element-container button[kind="primary"],
div.element-container:has(.main-nav-marker) + div button[kind="primary"] {
    background: rgba(155, 138, 92, 0.2) !important;
    border: 1.5px solid #9b8a5c !important;
    box-shadow: 0 0 12px rgba(155, 138, 92, 0.35) !important;
    color: var(--text-color) !important;
    font-weight: 700 !important;
}

div.element-container:has(.main-nav-marker) + div[data-testid="stHorizontalBlock"] button[kind="primary"] *,
div.element-container:has(.main-nav-marker) + div.element-container button[kind="primary"] *,
div.element-container:has(.main-nav-marker) + div button[kind="primary"] * {
    color: var(--text-color) !important;
}








.stTabs [data-baseweb="tab-list"]{background:#1c1c1c;border-radius:10px;padding:6px;


  box-shadow:0 2px 8px rgba(0,0,0,.3);gap:4px;}


.stTabs [data-baseweb="tab"]{border-radius:8px;font-weight:600;padding:8px 20px;font-size:.85rem;}


.stTabs [aria-selected="true"]{background:#9b8a5c!important;color:#000!important;}





.section-hdr{background:#9b8a5c;color:#000;padding:10px 16px;border-radius:8px;


  font-weight:600;font-size:.9rem;margin:16px 0 8px 0;}


.section-hdr-hom{background:#8c6e42;color:#fff;padding:10px 16px;border-radius:8px;


  font-weight:600;font-size:.9rem;margin:24px 0 8px 0;}


.info-box{background:#1a1a1a;border:1px solid #9b8a5c;border-radius:8px;


  padding:12px 16px;font-size:.85rem;color:#e5dccb;margin-bottom:12px;}


.warn-box{background:#2a1b00;border:1px solid #ffd257;border-radius:8px;


  padding:12px 16px;font-size:.85rem;color:#ffd257;margin-bottom:12px;}


div[data-testid="metric-container"]{background:#1c1c1c;border-radius:10px;


  padding:12px;box-shadow:0 2px 8px rgba(0,0,0,.3);}


/* Botões Gerais - Efeito Cristal Líquido (Glassmorphism Premium) */
/* Botões Gerais - Efeito Cristal Líquido (Glassmorphism Premium) */
.stButton button {
  background: rgba(128, 128, 128, 0.06) !important;
  color: var(--text-color) !important;
  border: 1px solid rgba(128, 128, 128, 0.2) !important;
  border-top: 1px solid rgba(255, 255, 255, 0.3) !important;
  border-radius: 20px !important;
  padding: 8px 24px !important;
  font-weight: 600 !important;
  font-size: 0.9rem !important;
  box-shadow: inset 0 1px 1px rgba(255, 255, 255, 0.1), 0 4px 12px rgba(0, 0, 0, 0.1) !important;
  backdrop-filter: blur(10px) !important;
  -webkit-backdrop-filter: blur(10px) !important;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
  letter-spacing: 0.02em !important;
}

.stButton button * {
  color: var(--text-color) !important;
}

.stButton button:hover {
  background: rgba(128, 128, 128, 0.12) !important;
  color: var(--text-color) !important;
  border-color: rgba(128, 128, 128, 0.3) !important;
  box-shadow: inset 0 1px 2px rgba(255, 255, 255, 0.2), 0 0 15px rgba(155, 138, 92, 0.2) !important;
  transform: translateY(-2px) !important;
}

.stButton button:hover * {
  color: var(--text-color) !important;
}

.stButton button:active {
  transform: translateY(1px) !important;
  box-shadow: inset 0 1px 3px rgba(0, 0, 0, 0.6) !important;
}

/* Botões Primários - Efeito Cristal Ouro (Ativo) */
.stButton button[data-testid="baseButton-primary"] {
  background: linear-gradient(135deg, rgba(155, 138, 92, 0.25) 0%, rgba(0, 0, 0, 0.6) 100%) !important;
  color: var(--text-color) !important;
  border: 1px solid rgba(155, 138, 92, 0.4) !important;
  border-top: 1px solid rgba(255, 255, 255, 0.4) !important;
  box-shadow: inset 0 1px 2px rgba(255, 255, 255, 0.25), 0 4px 15px rgba(155, 138, 92, 0.25) !important;
}

.stButton button[data-testid="baseButton-primary"] * {
  color: var(--text-color) !important;
}

.stButton button[data-testid="baseButton-primary"]:hover {
  background: linear-gradient(135deg, rgba(155, 138, 92, 0.45) 0%, rgba(255, 255, 255, 0.1) 100%) !important;
  border-color: rgba(155, 138, 92, 0.7) !important;
  box-shadow: inset 0 1px 3px rgba(255, 255, 255, 0.4), 0 0 25px rgba(155, 138, 92, 0.6) !important;
  color: var(--text-color) !important;
}

.stButton button[data-testid="baseButton-primary"]:hover * {
  color: var(--text-color) !important;
}

/* Sidebar Botões Secundários */
div[data-testid="stSidebar"] button[data-testid="baseButton-secondary"] {
  background: rgba(128, 128, 128, 0.05) !important;
  color: var(--text-color) !important;
  border: 1px solid rgba(128, 128, 128, 0.2) !important;
  border-top: 1px solid rgba(255, 255, 255, 0.25) !important;
  border-radius: 20px !important;
  box-shadow: inset 0 1px 1px rgba(255, 255, 255, 0.05), 0 4px 10px rgba(0,0,0,0.2) !important;
  font-weight: 500 !important;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
  backdrop-filter: blur(8px) !important;
  -webkit-backdrop-filter: blur(8px) !important;
}

div[data-testid="stSidebar"] button[data-testid="baseButton-secondary"] * {
  color: var(--text-color) !important;
}

div[data-testid="stSidebar"] button[data-testid="baseButton-secondary"]:hover {
  background: rgba(155, 138, 92, 0.15) !important;
  color: var(--text-color) !important;
  border-color: rgba(155, 138, 92, 0.4) !important;
  border-top-color: rgba(155, 138, 92, 0.6) !important;
  box-shadow: inset 0 1px 2px rgba(255, 255, 255, 0.15), 0 0 15px rgba(155, 138, 92, 0.3) !important;
  transform: translateY(-2px) !important;
}

div[data-testid="stSidebar"] button[data-testid="baseButton-secondary"]:hover * {
  color: var(--text-color) !important;
}

/* Sidebar Botões Primários */
div[data-testid="stSidebar"] button[data-testid="baseButton-primary"] {
  background: linear-gradient(135deg, rgba(155, 138, 92, 0.3) 0%, rgba(0, 0, 0, 0.7) 100%) !important;
  color: #ffffff !important;
  border: 1px solid rgba(155, 138, 92, 0.5) !important;
  border-top: 1px solid rgba(255, 255, 255, 0.4) !important;
  border-radius: 20px !important;
  box-shadow: inset 0 1px 3px rgba(255,255,255,0.3), 0 4px 15px rgba(188,163,116,0.3) !important;
  font-weight: 700 !important;
  text-shadow: 0 1px 2px rgba(0,0,0,0.8) !important;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
}

div[data-testid="stSidebar"] button[data-testid="baseButton-primary"]:hover {
  background: linear-gradient(135deg, rgba(255, 255, 255, 0.25) 0%, rgba(155, 138, 92, 0.45) 100%) !important;
  color: #ffffff !important;
  border-color: rgba(255, 255, 255, 0.7) !important;
  box-shadow: inset 0 1px 3px rgba(255,255,255,0.4), 0 0 25px rgba(188,163,116,0.5) !important;
  transform: translateY(-2px) !important;
}


/* Glassmorphic Crystal Style for Report Scope Selection Buttons */
div.element-container:has(.report-scope-marker) + div[data-testid="stHorizontalBlock"] button,
div.element-container:has(.report-scope-marker) + div.element-container button,
div.element-container:has(.report-scope-marker) + div button {
    background: rgba(128, 128, 128, 0.05) !important;
    backdrop-filter: blur(8px) !important;
    -webkit-backdrop-filter: blur(8px) !important;
    border: 1px solid rgba(128, 128, 128, 0.18) !important;
    border-radius: 8px !important;
    color: var(--text-color) !important;
    font-weight: 600 !important;
    padding: 6px 4px !important;
    transition: all 0.25s ease !important;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08) !important;
    
    /* Fixed unified dimensions and vertical centering */
    min-height: 95px !important;
    height: 95px !important;
    display: flex !important;
    flex-direction: column !important;
    justify-content: center !important;
    align-items: center !important;
}

div.element-container:has(.report-scope-marker) + div[data-testid="stHorizontalBlock"] button *,
div.element-container:has(.report-scope-marker) + div.element-container button *,
div.element-container:has(.report-scope-marker) + div button * {
    color: var(--text-color) !important;
    white-space: pre-line !important;
    text-align: center !important;
    font-size: 0.85rem !important;
    line-height: 1.3 !important;
    font-weight: 600 !important;
}

/* Highlight the emoji icon as larger */
div.element-container:has(.report-scope-marker) + div[data-testid="stHorizontalBlock"] button p::first-line,
div.element-container:has(.report-scope-marker) + div.element-container button p::first-line,
div.element-container:has(.report-scope-marker) + div button p::first-line,
div.element-container:has(.report-scope-marker) + div[data-testid="stHorizontalBlock"] button span::first-line,
div.element-container:has(.report-scope-marker) + div.element-container button span::first-line,
div.element-container:has(.report-scope-marker) + div button span::first-line {
    font-size: 1.75rem !important;
    line-height: 1.45 !important;
    font-weight: normal !important;
}

div.element-container:has(.report-scope-marker) + div[data-testid="stHorizontalBlock"] button:hover,
div.element-container:has(.report-scope-marker) + div.element-container button:hover,
div.element-container:has(.report-scope-marker) + div button:hover {
    background: rgba(128, 128, 128, 0.12) !important;
    border-color: rgba(128, 128, 128, 0.3) !important;
    transform: translateY(-1px) !important;
}

div.element-container:has(.report-scope-marker) + div[data-testid="stHorizontalBlock"] button:hover *,
div.element-container:has(.report-scope-marker) + div.element-container button:hover *,
div.element-container:has(.report-scope-marker) + div button:hover * {
    color: var(--text-color) !important;
}

/* Selected active button (primary) */
div.element-container:has(.report-scope-marker) + div[data-testid="stHorizontalBlock"] button[kind="primary"],
div.element-container:has(.report-scope-marker) + div.element-container button[kind="primary"],
div.element-container:has(.report-scope-marker) + div button[kind="primary"] {
    background: rgba(155, 138, 92, 0.22) !important;
    border: 1.5px solid #9b8a5c !important;
    box-shadow: 0 0 14px rgba(155, 138, 92, 0.45) !important;
    color: var(--text-color) !important;
}

div.element-container:has(.report-scope-marker) + div[data-testid="stHorizontalBlock"] button[kind="primary"] *,
div.element-container:has(.report-scope-marker) + div.element-container button[kind="primary"] *,
div.element-container:has(.report-scope-marker) + div button[kind="primary"] * {
    color: var(--text-color) !important;
}

/* Glassmorphic Crystal Style for Excel Scope Selection Buttons */
div.element-container:has(.excel-scope-marker) + div[data-testid="stHorizontalBlock"] button,
div.element-container:has(.excel-scope-marker) + div.element-container button,
div.element-container:has(.excel-scope-marker) + div button {
    background: rgba(128, 128, 128, 0.05) !important;
    backdrop-filter: blur(8px) !important;
    -webkit-backdrop-filter: blur(8px) !important;
    border: 1px solid rgba(128, 128, 128, 0.18) !important;
    border-radius: 8px !important;
    color: var(--text-color) !important;
    font-weight: 600 !important;
    padding: 6px 4px !important;
    transition: all 0.25s ease !important;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08) !important;
    
    /* Fixed unified dimensions and vertical centering */
    min-height: 95px !important;
    height: 95px !important;
    display: flex !important;
    flex-direction: column !important;
    justify-content: center !important;
    align-items: center !important;
}

div.element-container:has(.excel-scope-marker) + div[data-testid="stHorizontalBlock"] button *,
div.element-container:has(.excel-scope-marker) + div.element-container button *,
div.element-container:has(.excel-scope-marker) + div button * {
    color: var(--text-color) !important;
    white-space: pre-line !important;
    text-align: center !important;
    font-size: 0.85rem !important;
    line-height: 1.3 !important;
    font-weight: 600 !important;
}

/* Highlight the emoji icon as larger */
div.element-container:has(.excel-scope-marker) + div[data-testid="stHorizontalBlock"] button p::first-line,
div.element-container:has(.excel-scope-marker) + div.element-container button p::first-line,
div.element-container:has(.excel-scope-marker) + div button p::first-line,
div.element-container:has(.excel-scope-marker) + div[data-testid="stHorizontalBlock"] button span::first-line,
div.element-container:has(.excel-scope-marker) + div.element-container button span::first-line,
div.element-container:has(.excel-scope-marker) + div button span::first-line {
    font-size: 1.75rem !important;
    line-height: 1.45 !important;
    font-weight: normal !important;
}

div.element-container:has(.excel-scope-marker) + div[data-testid="stHorizontalBlock"] button:hover,
div.element-container:has(.excel-scope-marker) + div.element-container button:hover,
div.element-container:has(.excel-scope-marker) + div button:hover {
    background: rgba(128, 128, 128, 0.12) !important;
    border-color: rgba(128, 128, 128, 0.3) !important;
    transform: translateY(-1px) !important;
}

div.element-container:has(.excel-scope-marker) + div[data-testid="stHorizontalBlock"] button:hover *,
div.element-container:has(.excel-scope-marker) + div.element-container button:hover *,
div.element-container:has(.excel-scope-marker) + div button:hover * {
    color: var(--text-color) !important;
}

/* Selected active button (primary) */
div.element-container:has(.excel-scope-marker) + div[data-testid="stHorizontalBlock"] button[kind="primary"],
div.element-container:has(.excel-scope-marker) + div.element-container button[kind="primary"],
div.element-container:has(.excel-scope-marker) + div button[kind="primary"] {
    background: rgba(155, 138, 92, 0.22) !important;
    border: 1.5px solid #9b8a5c !important;
    box-shadow: 0 0 14px rgba(155, 138, 92, 0.45) !important;
    color: var(--text-color) !important;
}

div.element-container:has(.excel-scope-marker) + div[data-testid="stHorizontalBlock"] button[kind="primary"] *,
div.element-container:has(.excel-scope-marker) + div.element-container button[kind="primary"] *,
div.element-container:has(.excel-scope-marker) + div button[kind="primary"] * {
    color: var(--text-color) !important;
}

/* Force caqui theme colors globally for primary elements (buttons, radios, inputs) */
button[kind="primary"], 
button[data-testid="stBaseButton-primary"],
div[data-testid="stBaseButton-primary"] button,
div.stButton > button[kind="primary"],
div.stButton > button[data-testid="stBaseButton-primary"],
.st-emotion-cache-12w0qpk,
.st-emotion-cache-1f3w0dd,
.st-emotion-cache-zt52cr {
    background-color: #9b8a5c !important;
    background: #9b8a5c !important;
    color: #ffffff !important;
    border: 1px solid #9b8a5c !important;
    border-color: #9b8a5c !important;
}

button[kind="primary"]:hover, 
button[data-testid="stBaseButton-primary"]:hover,
div[data-testid="stBaseButton-primary"] button:hover,
div.stButton > button[kind="primary"]:hover,
div.stButton > button[data-testid="stBaseButton-primary"]:hover,
.st-emotion-cache-12w0qpk:hover,
.st-emotion-cache-1f3w0dd:hover,
.st-emotion-cache-zt52cr:hover {
    background-color: #83744c !important;
    background: #83744c !important;
    border-color: #83744c !important;
    color: #ffffff !important;
}

button[kind="primary"] *, 
button[data-testid="stBaseButton-primary"] *,
div[data-testid="stBaseButton-primary"] button *,
div.stButton > button[kind="primary"] *,
div.stButton > button[data-testid="stBaseButton-primary"] * {
    color: #ffffff !important;
}

div[data-checked="true"] > div,
div[data-testid="stRadio"] label[data-baseweb="radio"] div[data-checked="true"] > div,
.st-emotion-cache-16xl6g5[data-checked="true"] > div {
    background-color: #9b8a5c !important;
    border-color: #9b8a5c !important;
}

div[data-checked="true"],
div[data-testid="stRadio"] label[data-baseweb="radio"] div[data-checked="true"],
.st-emotion-cache-16xl6g5[data-checked="true"] {
    border-color: #9b8a5c !important;
}

div[data-baseweb="input"] > div:focus-within {
    border-color: #9b8a5c !important;
}

</style>


""", unsafe_allow_html=True)





# ─────────────────────── CONSTANTES ───────────────────────────────────────────


THIS_DIR  = Path(__file__).parent


DADOS_DIR = THIS_DIR / "dados"


DADOS_DIR.mkdir(exist_ok=True)


CONFIG_FILE = THIS_DIR / "config_aadp.json"





SITUACOES_ALVO = {
    "ATIV. DIRECAO GERAL", "ATIV. FIM DESTACADO", "ATIV. FIM NA SEDE",
    "ATIV. MEIO", "ATIVIDADE MEIO", "DISP MED DEFINITIVA", "QUADRO ESPECIALISTA"
}




CONCEITO_FAIXA = {


    "nivel superior de desempenho":       (9.00, 10.00),


    "nivel alto de desempenho":           (7.00,  8.99),


    "nivel intermediario de desempenho":  (6.00,  6.99),


    "nivel baixo de desempenho":          (3.00,  5.99),


    "nivel inferior de desempenho":       (0.00,  2.99),


}


STATUS_COLORS = {
    "Encerrada":                  "#70AD47",
    "Homologação":              "#FFD966",
    "Parcialmente Encerrada":     "#FF8C00",
    "Aberta":                     "#FF4444",
    "EM PRAZO DE RECURSO":         "#BDC3C7",
    "RECONSIDERAÇÃO COMISSÃO":    "#D6A2E8",
    "AUTORIDADE RECURSAL":        "#AF7AC5",
}


SIT_COLORS = {"Comissão Atual":"#4472C4","Nota Provisória":"#FFC000"}


# Ordem para empilhamento: Encerradas embaixo, pendentes em cima


STACK_ORDER  = [
    "Encerrada", "Aberta", "Parcialmente Encerrada", "Homologação",
    "EM PRAZO DE RECURSO", "RECONSIDERAÇÃO COMISSÃO", "AUTORIDADE RECURSAL"
]





# ─────────────────────── CONFIGURAÇÃO ─────────────────────────────────────────


def load_config():


    cfg = {"db_path": str(DADOS_DIR)}


    if CONFIG_FILE.exists():


        try:


            cfg.update(json.loads(CONFIG_FILE.read_text(encoding="utf-8")))


        except Exception:


            pass


    # Carrega do st.secrets do Streamlit para evitar perda de IDs/links após reinicializações


    base_drive_keys = [
        "drive_av_id", "drive_si_id", "drive_geral_id", "drive_com_id",
        "drive_sirh_id", "drive_master_xlsx_id", "drive_metas_id", "db_path"
    ]
    for b_key in base_drive_keys:
        for suffix in ["", "_2026", "_2027"]:
            k = f"{b_key}{suffix}"
            try:
                if k in st.secrets:
                    cfg[k] = st.secrets[k]
            except Exception:
                pass

    for key in ["sheet_api_url", "fonte_dados", "smtp_host", "smtp_port", "smtp_user", "smtp_pass", "alert_receiver_email", "alert_webhook_url"]:


        try:


            if key in st.secrets:


                cfg[key] = st.secrets[key]


        except Exception:


            pass


    return cfg






def get_active_year_config(selected_year: str, cfg: dict):
    """Retorna os caminhos de banco local e IDs do Google Drive para o ano selecionado (2026 ou 2027)."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    y = str(selected_year).strip()

    # Procura pasta local dedicada do ano
    possible_local_paths = [
        os.path.join(base_dir, f"DADOS AADP {y}"),
        os.path.join(base_dir, f"DADOS_AADP_{y}"),
        os.path.join(base_dir, f"dados AADP {y}"),
        os.path.join(base_dir, "dados", y),
        os.path.join(base_dir, y),
        os.path.join(base_dir, f"dados_{y}"),
        os.path.join(str(DADOS_DIR), f"DADOS AADP {y}"),
        os.path.join(str(DADOS_DIR), y)
    ]
    db_path_year = next((p for p in possible_local_paths if p and os.path.isdir(p)), None)
    if not db_path_year:
        db_path_year = cfg.get(f"db_path_{y}", cfg.get("db_path", str(DADOS_DIR)))

    def _safe_secret(sec_key):
        try:
            if hasattr(st, "secrets") and sec_key in st.secrets:
                return str(st.secrets[sec_key])
        except Exception:
            return ""
        return ""

    def _get_id(key):
        val = cfg.get(f"{key}_{y}")
        if not val:
            val = _safe_secret(f"{key}_{y}")
        if not val:
            val = cfg.get(key)
        if not val:
            val = _safe_secret(key)
        return str(val).strip() if val else ""

    return {
        "year": y,
        "db_path": db_path_year,
        "drive_av_id": _get_id("drive_av_id"),
        "drive_si_id": _get_id("drive_si_id"),
        "drive_geral_id": _get_id("drive_geral_id"),
        "drive_com_id": _get_id("drive_com_id"),
        "drive_sirh_id": _get_id("drive_sirh_id"),
        "drive_master_xlsx_id": _get_id("drive_master_xlsx_id"),
        "drive_metas_id": _get_id("drive_metas_id"),
    }


def save_config(cfg):


    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")





DB_FILE = str(Path(__file__).parent / "aadp_secure.db")





def init_db():


    conn = sqlite3.connect(DB_FILE)


    c = conn.cursor()


    c.execute("""


        CREATE TABLE IF NOT EXISTS users (


            pm TEXT PRIMARY KEY,


            name TEXT,


            rank TEXT,


            rpm TEXT,


            unit TEXT,


            function TEXT,


            role TEXT,


            status TEXT,


            password TEXT,


            created_at TEXT


        )


    """)


    c.execute("""


        CREATE TABLE IF NOT EXISTS logs (


            id INTEGER PRIMARY KEY AUTOINCREMENT,


            timestamp TEXT,


            pm TEXT,


            action TEXT,


            details TEXT


        )


    """)


    c.execute("SELECT * FROM users WHERE pm = 'ADM'")


    if not c.fetchone():


        adm_pass = hashlib.sha256("arquivosDRH2026".encode()).hexdigest()


        c.execute("""


            INSERT INTO users (pm, name, rank, rpm, unit, function, role, status, password, created_at)


            VALUES ('ADM', 'Administrador Geral', 'Desenvolvedor', 'Geral', 'DRH', 'Administrador', 'ADMINISTRADOR', 'Ativo', ?, ?)


        """, (adm_pass, now_br().strftime("%Y-%m-%d %H:%M:%S")))


    conn.commit()


    conn.close()





def log_action(pm: str, action: str, details: str = ""):


    timestamp = now_br().strftime("%Y-%m-%d %H:%M:%S")


    if check_use_cloud():


        run_sheet_api("add_log", {"log": {"timestamp": timestamp, "pm": pm, "action": action, "details": details}})


    else:


        try:


            conn = sqlite3.connect(DB_FILE)


            c = conn.cursor()


            c.execute("INSERT INTO logs (timestamp, pm, action, details) VALUES (?, ?, ?, ?)",


                      (timestamp, pm, action, details))


            conn.commit()


            conn.close()


        except Exception:


            pass





init_db()





cfg = load_config()





# ─────────────────────── DATABASE WRAPPERS (SQLite / Google Sheets Cloud) ───────


def check_use_cloud():


    url = cfg.get("sheet_api_url", "")


    return bool(url and url.strip().lower().startswith("http"))





def run_sheet_api(action, payload=None):


    url = cfg.get("sheet_api_url", "")


    if not url:


        return None


    import requests


    try:


        body = {"action": action}


        if payload:


            body.update(payload)


        r = requests.post(url, json=body, timeout=30)


        if r.status_code == 200:


            res = r.json()


            if res.get("status") == "success":


                return res.get("data")


            else:


                st.error(f"Erro na Planilha: {res.get('message')}")


    except Exception as e:


        st.error(f"Erro ao conectar com Google Sheets: {e}")


    return None





def refresh_db_cache():


    if check_use_cloud():


        users = run_sheet_api("get_users")


        if users is None:


            users = []


    else:


        try:


            conn = sqlite3.connect(DB_FILE)


            c = conn.cursor()


            c.execute("SELECT pm, name, rank, rpm, unit, function, password, role, status, created_at FROM users")


            rows = c.fetchall()


            conn.close()


            users = []


            for r in rows:


                users.append({


                    "pm": r[0], "name": r[1], "rank": r[2], "rpm": r[3], "unit": r[4],


                    "function": r[5], "password": r[6], "role": r[7], "status": r[8],


                    "created_at": r[9]


                })


        except Exception:


            users = []


    st.session_state.db_users = users


    return users


def send_new_user_alert(pm, name, rank, rpm, unit, function):
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header
    import urllib.request
    import json
    
    cfg = load_config()
    
    subject = f"⚠️ Novo Cadastro Pendente AADP 2026 - PM: {pm}"
    body = f"""Olá Administrador,

Um novo pedido de cadastro de acesso ao AADP 2026 foi realizado:

• Nº PM: {pm}
• Nome: {name}
• Posto/Graduação: {rank}
• RPM/UDG: {rpm}
• Unidade: {unit}
• Função: {function}

Por favor, acesse o Painel Administrador do sistema para aprovar ou recusar esta solicitação.
"""
    
    # 1. Webhook Alert (Zapier/Make/etc)
    webhook_url = cfg.get("alert_webhook_url", "").strip()
    if webhook_url:
        try:
            payload = {
                "event": "new_user_registration",
                "pm": pm,
                "name": name,
                "rank": rank,
                "rpm": rpm,
                "unit": unit,
                "function": function,
                "message": body
            }
            req = urllib.request.Request(
                webhook_url,
                data=json.dumps(payload).encode("utf-8"),
                headers={"Content-Type": "application/json"}
            )
            with urllib.request.urlopen(req, timeout=5) as response:
                pass
        except Exception as e:
            log_action("SYSTEM", "ALERT_WEBHOOK_ERROR", f"Falha no webhook de alerta: {str(e)}")

    # 2. SMTP Email Alert
    smtp_host = cfg.get("smtp_host", "").strip()
    smtp_port = str(cfg.get("smtp_port", "")).strip()
    smtp_user = cfg.get("smtp_user", "").strip()
    smtp_pass = cfg.get("smtp_pass", "").replace(" ", "").strip()
    receiver = cfg.get("alert_receiver_email", "").strip()
    
    if smtp_host and smtp_port and smtp_user and smtp_pass and receiver:
        try:
            msg = MIMEText(body, "plain", "utf-8")
            msg["Subject"] = Header(subject, "utf-8")
            msg["From"] = smtp_user
            msg["To"] = receiver
            
            port = int(smtp_port)
            if port == 465:
                server = smtplib.SMTP_SSL(smtp_host, port, timeout=10)
                server.login(smtp_user, smtp_pass)
                server.sendmail(smtp_user, [receiver], msg.as_string())
                server.quit()
            else:
                server = smtplib.SMTP(smtp_host, port, timeout=10)
                server.ehlo()
                server.starttls()
                server.login(smtp_user, smtp_pass)
                server.sendmail(smtp_user, [receiver], msg.as_string())
                server.quit()
            log_action("SYSTEM", "ALERT_EMAIL_SENT", f"Email de alerta enviado para {receiver}")
        except Exception as e:
            log_action("SYSTEM", "ALERT_EMAIL_ERROR", f"Falha no envio de email: {str(e)}")


def get_cached_users():


    if "db_users" not in st.session_state:


        refresh_db_cache()
    
    users = st.session_state.db_users
    latest = {}
    for u in users:
        pm = str(u["pm"]).strip()
        latest[pm] = u
    return list(latest.values())





def db_get_user_for_login(pm, password_hash):


    if check_use_cloud():


        users = run_sheet_api("get_users")


        if users:


            for u in users:


                if str(u["pm"]).strip() == str(pm).strip() and str(u["password"]).strip() == str(password_hash).strip():


                    return (u["name"], u["role"], u["rpm"], u["unit"], u["status"])


        return None


    else:


        conn = sqlite3.connect(DB_FILE)


        c = conn.cursor()


        c.execute("SELECT name, role, rpm, unit, status FROM users WHERE pm = ? AND password = ?", (pm, password_hash))


        res = c.fetchone()


        conn.close()


        return res





def db_register_user(pm, name, rank, rpm, unit, function, password_hash):


    created_at = now_br().strftime("%Y-%m-%d %H:%M:%S")


    success = False


    if check_use_cloud():


        user = {


            "pm": pm, "name": name, "rank": rank, "rpm": rpm, "unit": unit,


            "function": function, "role": "AGUARDANDO", "status": "Pendente",


            "password": password_hash, "created_at": created_at


        }


        res = run_sheet_api("add_user", {"user": user})


        success = res is not None


    else:


        try:


            conn = sqlite3.connect(DB_FILE)


            c = conn.cursor()


            c.execute("""


                INSERT INTO users (pm, name, rank, rpm, unit, function, role, status, password, created_at)


                VALUES (?, ?, ?, ?, ?, ?, 'AGUARDANDO', 'Pendente', ?, ?)


            """, (pm, name, rank, rpm, unit, function, password_hash, created_at))


            conn.commit()


            conn.close()


            success = True


        except Exception:


            success = False


    if success:


        refresh_db_cache()


    return success





def db_get_all_pms():


    users = get_cached_users()


    return [str(u["pm"]) for u in users]





def db_update_user_info(pm, name, rank, rpm, unit, sector):


    if check_use_cloud():


        updates = {"name": name, "rank": rank, "rpm": rpm, "unit": unit, "function": sector}


        run_sheet_api("update_user", {"pm": pm, "updates": updates})


    else:


        conn = sqlite3.connect(DB_FILE)


        c = conn.cursor()


        c.execute("""


            UPDATE users


            SET name = ?, rank = ?, rpm = ?, unit = ?, function = ?


            WHERE pm = ?


        """, (name, rank, rpm, unit, sector, pm))


        conn.commit()


        conn.close()


    refresh_db_cache()





def db_get_user_password(pm):


    users = get_cached_users()


    for u in users:


        if str(u["pm"]).strip() == str(pm).strip():


            return u["password"]


    return None





def db_update_password(pm, password_hash):


    if check_use_cloud():


        run_sheet_api("update_user", {"pm": pm, "updates": {"password": password_hash}})


    else:


        conn = sqlite3.connect(DB_FILE)


        c = conn.cursor()


        c.execute("UPDATE users SET password = ? WHERE pm = ?", (password_hash, pm))


        conn.commit()


        conn.close()


    refresh_db_cache()





def db_get_simulator_users():


    users = get_cached_users()


    sim_users = []


    for u in users:


        if u["status"] == "Ativo" and str(u["pm"]) != "ADM":


            sim_users.append((u["pm"], u["name"], u["rank"], u["role"], u["rpm"], u["unit"]))


    sim_users.sort(key=lambda x: x[1])


    return sim_users





def db_get_pending_users():


    users = get_cached_users()
    
    # Mapeia PMs que já possuem um status resolvido (Ativo, Recusado, Bloqueado)
    non_pending_pms = {str(u["pm"]).strip() for u in users if u["status"] != "Pendente"}


    pend_users = []
    seen = set()


    for u in users:
        pm = str(u["pm"]).strip()
        # Se o PM não tem status resolvido e ainda não foi listado nas pendências
        if u["status"] == "Pendente" and pm not in non_pending_pms and pm not in seen:


            pend_users.append((u["pm"], u["name"], u["rank"], u["rpm"], u["unit"], u["function"], u["created_at"]))
            seen.add(pm)


    return pend_users





def db_get_active_users():
    users = get_cached_users()
    active = []
    for u in users:
        if u["status"] == "Ativo" and str(u["pm"]) != "ADM":
            active.append((
                u["pm"],
                u["rank"],
                u["name"],
                u.get("rpm", ""),
                u.get("unit", ""),
                u.get("function", ""),  # Setor = NOME UNIDADE (col J SIGEF)
                u["role"],
                u["created_at"]
            ))
    return active





def db_approve_user(pm, role, rpm):


    if check_use_cloud():


        run_sheet_api("update_user", {"pm": pm, "updates": {"status": "Ativo", "role": role, "rpm": rpm}})


    else:


        conn = sqlite3.connect(DB_FILE)


        c = conn.cursor()


        c.execute("UPDATE users SET status = 'Ativo', role = ?, rpm = ? WHERE pm = ?", (role, rpm, pm))


        conn.commit()


        conn.close()


    refresh_db_cache()





def db_reject_user(pm):


    if check_use_cloud():


        run_sheet_api("update_user", {"pm": pm, "updates": {"status": "Recusado"}})


    else:


        conn = sqlite3.connect(DB_FILE)


        c = conn.cursor()


        c.execute("UPDATE users SET status = 'Recusado' WHERE pm = ?", (pm,))


        conn.commit()


        conn.close()


    refresh_db_cache()





def db_update_user_role_rpm(pm, role, rpm):


    if check_use_cloud():


        run_sheet_api("update_user", {"pm": pm, "updates": {"role": role, "rpm": rpm}})


    else:


        conn = sqlite3.connect(DB_FILE)


        c = conn.cursor()


        c.execute("UPDATE users SET role = ?, rpm = ? WHERE pm = ?", (role, rpm, pm))


        conn.commit()


        conn.close()


    refresh_db_cache()





def db_revoke_user(pm):


    if check_use_cloud():


        run_sheet_api("update_user", {"pm": pm, "updates": {"status": "Bloqueado"}})


    else:


        conn = sqlite3.connect(DB_FILE)


        c = conn.cursor()


        c.execute("UPDATE users SET status = 'Bloqueado' WHERE pm = ?", (pm,))


        conn.commit()


        conn.close()


    refresh_db_cache()





def db_get_users_df():


    users = get_cached_users()


    rows = []


    for u in users:


        if str(u["pm"]) != "ADM":


            rows.append({


                "pm": u["pm"], "rank": u["rank"], "name": u["name"],


                "role": u["role"], "rpm": u["rpm"], "unit": u["unit"],


                "status": u["status"]


            })


    return pd.DataFrame(rows)





def format_log_timestamp(ts_str):
    if not ts_str:
        return ts_str
    try:
        ts_clean = str(ts_str).strip()
        if "T" in ts_clean or ts_clean.endswith("Z"):
            dt = pd.to_datetime(ts_clean)
            if dt.tzinfo is None:
                dt = dt.tz_localize("UTC")
            dt_br = dt.tz_convert(timezone(timedelta(hours=-3)))
            return dt_br.strftime("%Y-%m-%d %H:%M:%S")
        else:
            dt = pd.to_datetime(ts_clean)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return ts_str

def refresh_logs_cache():
    if check_use_cloud():
        logs = run_sheet_api("get_logs")
        if logs is None:
            logs = []
        else:
            for l in logs:
                if "timestamp" in l:
                    l["timestamp"] = format_log_timestamp(l["timestamp"])
    else:
        try:
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT timestamp, pm, action, details FROM logs ORDER BY id DESC")
            rows = c.fetchall()
            conn.close()
            logs = []
            for r in rows:
                logs.append({
                    "timestamp": format_log_timestamp(r[0]), "pm": r[1], "action": r[2], "details": r[3]
                })
        except Exception:
            logs = []
            
    st.session_state.db_logs = logs
    return logs





def get_cached_logs():


    if "db_logs" not in st.session_state:


        refresh_logs_cache()


    return st.session_state.db_logs





def db_get_logs_pms():


    logs = get_cached_logs()


    pms = set(str(l["pm"]) for l in logs)


    return list(pms)





def db_get_user_info(pm):


    users = get_cached_users()


    for u in users:


        if str(u["pm"]).strip() == str(pm).strip():


            return (u["rank"], u["name"])


    return None





def db_get_logs_df(sel_log_user, start_str, end_str):


    logs = get_cached_logs()


    rows = []


    for l in logs:


        t_date = l["timestamp"][:10]


        if start_str <= t_date <= end_str:


            if sel_log_user == "Todos" or str(l["pm"]).strip() == str(sel_log_user).strip():


                rows.append({


                    "timestamp": l["timestamp"],


                    "pm": l["pm"],


                    "action": l["action"],


                    "details": l["details"]


                })


    return pd.DataFrame(rows)





def db_get_pending_count():


    return len(db_get_pending_users())





def db_check_user_status(pm):


    users = get_cached_users()


    for u in users:


        if str(u["pm"]).strip() == str(pm).strip():


            return u["status"]


    return None





def _auto_detect_role(sector: str):
    """
    Determina automaticamente o perfil do usuário com base no NOME UNIDADE (col J do SIGEF).
    Regras:
      - Se contiver 'SADM' → perfil 'SADM' (liberação automática)
      - Se contiver 'P1'   → perfil 'P1'   (liberação automática)
      - Qualquer outro valor → None (cadastro fica PENDENTE para o administrador)
    NOTA: Perfis GESTOR e ADMINISTRADOR NUNCA são liberados automaticamente.
    """
    s = str(sector).strip().upper()
    if "SADM" in s:
        return "SADM"
    if "P1" in s:
        return "P1"
    return None


def db_re_request_access(pm, name, rank, rpm, unit, sector, password_hash):


    created_at = now_br().strftime("%Y-%m-%d %H:%M:%S")


    success = False

    # Detecta perfil automático pelo NOME UNIDADE (col J SIGEF)
    auto_role = _auto_detect_role(sector)
    final_role   = auto_role if auto_role else "PENDENTE"
    final_status = "Ativo"   if auto_role else "Pendente"


    if check_use_cloud():


        updates = {


            "name": name, "rank": rank, "rpm": rpm, "unit": unit, "function": sector,


            "role": final_role, "status": final_status, "password": password_hash, "created_at": created_at


        }


        res = run_sheet_api("update_user", {"pm": pm, "updates": updates})


        success = res is not None


    else:


        try:


            conn = sqlite3.connect(DB_FILE)


            c = conn.cursor()


            c.execute("""


                UPDATE users 


                SET name = ?, rank = ?, rpm = ?, unit = ?, function = ?, role = ?, status = ?, password = ?, created_at = ?


                WHERE pm = ?


            """, (name, rank, rpm, unit, sector, final_role, final_status, password_hash, created_at, pm))


            conn.commit()


            conn.close()


            success = True


        except Exception:


            success = False


    if success:
        refresh_db_cache()
        if not auto_role:
            # Só envia alerta quando não há liberação automática
            try:
                send_new_user_alert(pm, name, rank, rpm, unit, sector)
            except Exception:
                pass

    return success





def db_create_new_request(pm, name, rank, rpm, unit, sector, password_hash):


    created_at = now_br().strftime("%Y-%m-%d %H:%M:%S")


    success = False

    # Detecta perfil automático pelo NOME UNIDADE (col J SIGEF)
    auto_role = _auto_detect_role(sector)
    final_role   = auto_role if auto_role else "PENDENTE"
    final_status = "Ativo"   if auto_role else "Pendente"


    if check_use_cloud():


        user = {


            "pm": pm, "name": name, "rank": rank, "rpm": rpm, "unit": unit, "function": sector,


            "role": final_role, "status": final_status, "password": password_hash, "created_at": created_at


        }


        res = run_sheet_api("add_user", {"user": user})


        success = res is not None


    else:


        try:


            conn = sqlite3.connect(DB_FILE)


            c = conn.cursor()


            c.execute("""


                INSERT INTO users (pm, name, rank, rpm, unit, function, role, status, password, created_at)


                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)


            """, (pm, name, rank, rpm, unit, sector, final_role, final_status, password_hash, created_at))


            conn.commit()


            conn.close()


            success = True


        except Exception:


            success = False


    if success:
        refresh_db_cache()
        if not auto_role:
            # Só envia alerta quando não há liberação automática
            try:
                send_new_user_alert(pm, name, rank, rpm, unit, sector)
            except Exception:
                pass

    return success





# ─────────────────────── LÓGICA DADOS ─────────────────────────────────────────


def normaliza(t):


    s = unicodedata.normalize("NFD", t.lower())


    return "".join(c for c in s if unicodedata.category(c) != "Mn")





def is_empty(v):


    if v == 0 or v == 0.0:


        return False


    return not v or str(v).strip() in ("", "-", "nan", "none")





def concordam(j, l):


    if is_empty(j) or is_empty(l): return None


    try: nota = float(str(l).replace(",","."))


    except: return None


    faixa = CONCEITO_FAIXA.get(normaliza(j.strip()))


    if faixa is None: return None


    return faixa[0] <= nota <= faixa[1]





def matches_rpm(reg_rpm, csv_rpm):


    s_reg = str(reg_rpm).strip()


    s_csv = str(csv_rpm).strip()


    if s_reg.lower() == "gestor":


        return True


    if s_reg.lower() == s_csv.lower():


        return True


    


    # Extrair dígitos de forma ultra rápida


    reg_digits = "".join(c for c in s_reg if c.isdigit())


    csv_digits = "".join(c for c in s_csv if c.isdigit())


    


    if reg_digits and csv_digits:


        return int(reg_digits) == int(csv_digits)


    return False





def find_sigef_user(pm_number: str) -> dict:


    """Busca os dados do militar no SIGEF.csv pelo Nº PM (6 dígitos)."""


    pm_clean = pm_number.strip().lstrip("0")


    if not pm_clean:


        return None


    try:


        import os, csv


        si_path = "SIGEF.csv"


        # Se não existe localmente, tenta baixar do Google Drive


        if not os.path.exists(si_path):


            sigef_drive_id = "10Ld_4XEz9b4kI_T6TC9W19tQdtBJCz5F"


            try:


                _baixar_drive(sigef_drive_id, si_path)


            except Exception as e:


                # Tenta olhar na subpasta dados/


                si_path = os.path.join("dados", "SIGEF.csv")


                if not os.path.exists(si_path):


                    st.error(f"Erro ao baixar base SIGEF do Google Drive: {e}")


                    return None


        with open(si_path, encoding="cp1252", errors="replace") as f:


            reader = csv.reader(f, delimiter=";")


            header = next(reader)


            try:
                idx_birth = header.index("DATA NASCIMENTO")
            except ValueError:
                idx_birth = 17
            try:
                idx_cpf = header.index("NUMERO CPF")
            except ValueError:
                idx_cpf = 25

            for row in reader:
                if len(row) > max(idx_birth, idx_cpf):
                    curr_pm = row[0].strip().lstrip("0")
                    if curr_pm == pm_clean:
                        
                        idx_offset = 0
                        if len(row) > 16 and row[15] in ["A", "I"]:
                            idx_offset = -1
                        elif len(row) > 17 and row[16] in ["A", "I"]:
                            idx_offset = 0
                        elif len(row) > 18 and row[17] in ["A", "I"]:
                            idx_offset = 1
                            
                        # Ajustar indices reais para este row
                        real_birth = idx_birth + idx_offset
                        real_cpf = idx_cpf + idx_offset
                        
                        return {
                            "pm": row[0].strip(),
                            "rank": row[2].strip().title(),
                            "name": row[3].strip().title(),
                            "rpm": row[5].strip(),      # UDI/UDG (NOME RPM)
                            "unit": row[7].strip(),     # Unidade Principal (NOME UNIDADE PRINCIPAL)
                            "sector": row[9].strip(),    # Setor (NOME UNIDADE)
                            "birthdate": row[real_birth].strip() if len(row) > real_birth else "",
                            "cpf": row[real_cpf].strip() if len(row) > real_cpf else ""
                        }


    except Exception as e:


        st.error(f"Erro ao ler banco SIGEF: {e}")


    return None





def sync_users_with_sigef():


    """Sincroniza os dados cadastrais de todos os usuários com o SIGEF.csv de forma otimizada."""


    try:


        import os, csv


        si_path = "SIGEF.csv"


        if not os.path.exists(si_path):


            si_path = os.path.join("dados", "SIGEF.csv")


            if not os.path.exists(si_path):


                # Tenta baixar


                sigef_drive_id = "10Ld_4XEz9b4kI_T6TC9W19tQdtBJCz5F"


                try:


                    _baixar_drive(sigef_drive_id, "SIGEF.csv")


                    si_path = "SIGEF.csv"


                except Exception:


                    return





        # Carrega todo o SIGEF para um dicionário em memória (uma única leitura de disco)


        sigef_dict = {}


        with open(si_path, encoding="cp1252", errors="replace") as f:


            reader = csv.reader(f, delimiter=";")


            next(reader)  # pula cabeçalho


            for row in reader:


                if len(row) > 24:


                    pm_clean = row[0].strip().lstrip("0")


                    if pm_clean:


                        sigef_dict[pm_clean] = {
                            "name": row[3].strip(),
                            "rank": row[2].strip(),
                            "rpm": row[5].strip(),
                            "unit": row[7].strip(),
                            "sector": row[9].strip()
                        }


        


        pms = db_get_all_pms()


        for pm in pms:


            if pm == "ADM":


                continue


            pm_clean = pm.strip().lstrip("0")


            if pm_clean in sigef_dict:


                info = sigef_dict[pm_clean]


                db_update_user_info(pm, info["name"], info["rank"], info["rpm"], info["unit"], info["sector"])


    except Exception:


        pass





def calc_cert(j, l):


    if is_empty(j) or is_empty(l): return "-"


    c = concordam(j, l)


    return "NÃO" if c is True else ("SIM" if c is False else "-")





def calc_status(j, l, n):


    if is_empty(j):


        return "Aberta"


    if is_empty(l):


        return "Parcialmente Encerrada"


    c = concordam(j, l)


    if c is True:


        return "Encerrada"


    elif c is False:


        return "Encerrada" if not is_empty(n) else "Homologação"


    return "Parcialmente Encerrada"





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
    url = f"https://drive.google.com/uc?id={file_id}"
    
    sig = inspect.signature(gdown.download)
    if "id" in sig.parameters:
        try:
            gdown.download(id=file_id, output=destino, quiet=True)
        except Exception:
            if "fuzzy" in sig.parameters:
                gdown.download(url, destino, quiet=True, fuzzy=True)
            else:
                gdown.download(url, destino, quiet=True)
    elif "fuzzy" in sig.parameters:
        gdown.download(url, destino, quiet=True, fuzzy=True)
    else:
        gdown.download(url, destino, quiet=True)
        
    if not os.path.exists(destino) or os.path.getsize(destino) == 0:
        raise FileNotFoundError(f"Falha ao baixar arquivo do Drive (ID: {file_id})")
        
    try:
        with open(destino, "rb") as f:
            head = f.read(200)
        if b"<!DOCTYPE" in head.upper() or b"<HTML" in head.upper():
            raise ValueError("O Google Drive retornou uma página HTML em vez do arquivo binário. Isso ocorre se o arquivo não estiver compartilhado como público ('Qualquer pessoa com o link pode ler'), se o ID estiver incorreto, ou se você estiver tentando baixar uma Planilha Google (Google Sheets) em vez de um arquivo Excel (.xlsx) carregado no Drive.")
    except ValueError as ve:
        if os.path.exists(destino):
            os.remove(destino)
        raise ve
    except Exception:
        pass


def _parse_csv(av_f: str, si_f: str) -> pd.DataFrame:
    """Processa os dois CSVs e retorna o DataFrame final."""
    sigef = {}
    with open(si_f, encoding="cp1252", errors="replace") as f:
        for row in csv.reader(f, delimiter=";"):
            if len(row) > 9:
                sigef[row[0].strip().lstrip("0") or "0"] = row[9].strip()

    # Carregar o geral.csv se disponível para obter as colunas de recursos e calcular o status de recurso
    geral_f = os.path.join(os.path.dirname(av_f), "geral.csv")
    recourse_map = {}
    cdp_map = {}
    pm_max_cdp = {}
    
    def find_col_index(header_list, name_pattern):
        def norm(s):
            import unicodedata
            t = unicodedata.normalize("NFD", str(s).lower().replace(" ", "").replace("_", "").replace("-", "").replace("(", "").replace(")", "").replace("/", ""))
            return "".join(c for c in t if unicodedata.category(c) != "Mn")
        pat = norm(name_pattern)
        for i, col in enumerate(header_list):
            if pat in norm(col):
                return i
        raise ValueError(f"Coluna contendo '{name_pattern}' nao encontrada.")

    def normalize_pm_str(pm_val):
        try:
            if not pm_val or str(pm_val).strip() in ("", "-", "nan", "none", "None", "<NA>"):
                return ""
            return str(int(float(str(pm_val).strip())))
        except Exception:
            return str(pm_val).strip()

    def parse_float_val(s_val):
        if not s_val or str(s_val).strip() in ("", "-", "nan", "none", "None", "<NA>"):
            return None
        try:
            return float(str(s_val).replace(",", "."))
        except ValueError:
            return None

    if os.path.exists(geral_f):
        try:
            with open(geral_f, "r", encoding="cp1252", errors="ignore") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader)
                
                c_pm_g = find_col_index(header, "nrPM (Avaliado)")
                try:
                    c_dt_cdp_g = find_col_index(header, "Data do CDP")
                except ValueError:
                    c_dt_cdp_g = -1
                c_concept_g = find_col_index(header, "Conceito Geral")
                c_grade_g = find_col_index(header, "Nota Geral")
                c_dt_av1_g = find_col_index(header, "Data da Avaliação 1")
                c_dt_av2_g = find_col_index(header, "Data da Avaliação 2")
                c_dt_hom_g = find_col_index(header, "Data da Homologação")
                c_n_hom_g = find_col_index(header, "Nota da Homologação")
                
                c_n_f1_g = find_col_index(header, "Nota (Fase 1)")
                c_n_f2_g = find_col_index(header, "Nota (Fase 2)")
                c_n_f3_g = find_col_index(header, "Nota (Fase 3)")
                c_n_f4_g = find_col_index(header, "Nota (Fase 4)")
                c_r_f1_g = find_col_index(header, "Recurso Fase 1")
                c_r_f2_g = find_col_index(header, "Recurso Fase 2")
                c_r_f3_g = find_col_index(header, "Recurso Fase 3")
                c_r_f4_g = find_col_index(header, "Recurso Fase 4")
                c_dt_f1_g = find_col_index(header, "Data Cadastro (Fase 1)")
                c_dt_f2_g = find_col_index(header, "Data Cadastro (Fase 2)")
                c_dt_f3_g = find_col_index(header, "Data Cadastro (Fase 3)")
                c_dt_f4_g = find_col_index(header, "Data Cadastro (Fase 4)")
                c_av1_g = find_col_index(header, "nrPM (Avaliador 1)")
                c_av2_g = find_col_index(header, "nrPM (Avaliador 2)")
                
                def parse_date_internal(d_str):
                    if not d_str or str(d_str).strip() in ("", "-", "nan", "none", "None", "<NA>"): return None
                    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
                        try:
                            return datetime.strptime(str(d_str).strip(), fmt).date()
                        except Exception:
                            continue
                    return None

                def add_business_days_internal(start_date, num_days):
                    curr = start_date
                    added = 0
                    while added < num_days:
                        curr += timedelta(days=1)
                        if curr.weekday() < 5: # Mon-Fri
                            added += 1
                    return curr

                for row_g in reader:
                    if len(row_g) > 38 and row_g[38] == "Disciplina":
                        row_g = row_g[:18] + [""] * 10 + row_g[18:]
                    while len(row_g) < len(header): row_g.append("")
                    pm_g = normalize_pm_str(row_g[c_pm_g])
                    av1_g = normalize_pm_str(row_g[c_av1_g])
                    av2_g = normalize_pm_str(row_g[c_av2_g])
                    concept_g = row_g[c_concept_g].strip()
                    
                    if c_dt_cdp_g != -1:
                        dt_cdp_val = parse_date_internal(row_g[c_dt_cdp_g])
                        if dt_cdp_val:
                            cdp_map[(pm_g, av1_g, av2_g)] = dt_cdp_val
                            if pm_g not in pm_max_cdp or dt_cdp_val > pm_max_cdp[pm_g]:
                                pm_max_cdp[pm_g] = dt_cdp_val

                    grade_g = row_g[c_grade_g].strip()
                    n_hom_g = row_g[c_n_hom_g].strip()
                    
                    n_f4_val = parse_float_val(row_g[c_n_f4_g])
                    n_f3_val = parse_float_val(row_g[c_n_f3_g])
                    n_f2_val = parse_float_val(row_g[c_n_f2_g])
                    n_f1_val = parse_float_val(row_g[c_n_f1_g])
                    
                    r_f4_val = row_g[c_r_f4_g].strip()
                    r_f3_val = row_g[c_r_f3_g].strip()
                    r_f2_val = row_g[c_r_f2_g].strip()
                    r_f1_val = row_g[c_r_f1_g].strip()
                    
                    dt_f4_val = row_g[c_dt_f4_g].strip()
                    dt_f3_val = row_g[c_dt_f3_g].strip()
                    dt_f2_val = row_g[c_dt_f2_g].strip()
                    dt_f1_val = row_g[c_dt_f1_g].strip()
                    
                    c_g = concordam(concept_g, grade_g)
                    has_appeal_g = (r_f1_val not in ("", "-")) or (n_f1_val is not None)
                    ref_date_g = now_br().date()
                    
                    final_status_g = calc_status(concept_g, grade_g, n_hom_g)
                    
                    if final_status_g in ("Aberta", "Parcialmente Encerrada"):
                        pass
                    elif not has_appeal_g:
                        if c_g is False:
                            # Houve discordância: necessita passar para o homologador
                            if is_empty(n_hom_g):
                                final_status_g = "Homologação"
                            else:
                                dt_base = parse_date_internal(row_g[c_dt_hom_g])
                                if dt_base is not None:
                                    deadline = add_business_days_internal(dt_base, 5)
                                    if ref_date_g <= deadline:
                                        final_status_g = "EM PRAZO DE RECURSO"
                                    else:
                                        final_status_g = "Encerrada"
                                else:
                                    final_status_g = "Encerrada"
                        else:
                            # Não houve discordância: prazo de 5 dias úteis a partir da data de AV2
                            dt_base = parse_date_internal(row_g[c_dt_av2_g])
                            if dt_base is not None:
                                deadline = add_business_days_internal(dt_base, 5)
                                if ref_date_g <= deadline:
                                    final_status_g = "EM PRAZO DE RECURSO"
                                else:
                                    final_status_g = "Encerrada"
                            else:
                                final_status_g = "Encerrada"
                    else:
                        # Houve recurso (r_f1 registrado = militar interpôs recurso)
                        if n_f4_val is not None:
                            # Fase 4: encerra definitivamente
                            final_status_g = "Encerrada"
                        elif r_f3_val not in ("", "-") or n_f3_val is not None:
                            # Fase 3 registrada = Autoridade Recursal DECIDIU → Encerrada
                            final_status_g = "Encerrada"
                        elif r_f2_val not in ("", "-"):
                            # Fase 2 registrada = comissão promoveu para Autoridade Recursal
                            if n_f2_val is not None:
                                final_status_g = "Encerrada"
                            else:
                                final_status_g = "AUTORIDADE RECURSAL"
                        else:
                            # Só Fase 1 registrada = comissão analisando (Reconsideração)
                            final_status_g = "RECONSIDERAÇÃO COMISSÃO"
                    key = (pm_g, av1_g, av2_g)
                    recourse_map[key] = final_status_g
        except Exception:
            pass

    # Contagem de instâncias de avaliação ativas por PM
    pm_counts = {}
    with open(av_f, encoding="cp1252", errors="replace") as f:
        reader = csv.reader(f, delimiter=";")
        try:
            next(reader)
        except StopIteration:
            pass
        for row in reader:
            if not row:
                continue
            while len(row) < 8:
                row.append("")
            nrpm = row[0].strip()
            pm_counts[nrpm] = pm_counts.get(nrpm, 0) + 1

    rows = []
    with open(av_f, encoding="cp1252", errors="replace") as f:
        reader = csv.reader(f, delimiter=";")
        next(reader)
        for row in reader:
            while len(row) < 50: row.append("")  # CSV tem 50 colunas (colégio até Homologador)
            nrpm = row[0].strip(); local = row[5].strip()
            j = row[9].strip(); l = row[11].strip(); n = row[13].strip()
            
            is_same_location = (local.upper().strip() == sigef.get(nrpm.lstrip("0") or "0", "").upper().strip())
            has_multiple_evals = (pm_counts.get(nrpm, 0) > 1)
            
            pm_norm = normalize_pm_str(nrpm)
            av1_val = row[26].strip()
            av2_val = row[34].strip()
            av1_norm = normalize_pm_str(av1_val)
            av2_norm = normalize_pm_str(av2_val)
            key = (pm_norm, av1_norm, av2_norm)
            
            dt_cdp = cdp_map.get(key)
            if dt_cdp is not None and pm_norm in pm_max_cdp:
                sc = "Comissão Atual" if dt_cdp >= pm_max_cdp[pm_norm] else "Nota Provisória"
            else:
                sc = "Comissão Atual" if (is_same_location or not has_multiple_evals) else "Nota Provisória"
            
            # Verificar se temos o status do recurso no recourse_map
            pm_norm = normalize_pm_str(nrpm)
            av1_val = row[26].strip()
            av2_val = row[34].strip()
            av1_norm = normalize_pm_str(av1_val)
            av2_norm = normalize_pm_str(av2_val)
            key = (pm_norm, av1_norm, av2_norm)
            status_av = recourse_map.get(key)
            if status_av is None:
                status_av = calc_status(j, l, n)


            rows.append({


                "nrPM (Avaliado)":          nrpm,


                "Nome (Avaliado)":           row[1].strip(),


                "Posto/Grad. (Avaliado)":    row[2].strip(),


                "Unidade RPM (Avaliado)":    row[3].strip(),


                "Unidade Principal (Avaliado)": row[4].strip(),


                "Local/Unidade (Avaliado)":  local,


                "Quadro Atual (Avaliado)":   row[6].strip(),


                "Sit. Funcional":        "ATIV. MEIO" if row[7].strip() == "ATIVIDADE MEIO" else row[7].strip(),


                "Data AV1":                  row[8].strip(),


                "Conceito Geral":            j,


                "Data AV2":                  row[12].strip() if not is_empty(n) else row[10].strip(),


                "Nota Geral":                n if not is_empty(n) else l,


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


                "Status Avaliação":  status_av,


            })


    return pd.DataFrame(rows)





@st.cache_resource(show_spinner="⏳ Carregando e processando dados...")
def load_data(db_path: str, drive_av_id: str = "", drive_si_id: str = "", drive_geral_id: str = "", ano: str = "2026"):
    """Carrega dados de pasta local ou Google Drive e gera o Geral.xlsx automaticamente."""
    if drive_av_id and drive_si_id:
        # ── Modo Google Drive ──────────────────────────────────────────────
        cache_dir = os.path.join(tempfile.gettempdir(), f"aadp_drive_cache_{ano}")
        os.makedirs(cache_dir, exist_ok=True)
        av_f = os.path.join(cache_dir, "avaliacoes.csv")
        si_f = os.path.join(cache_dir, "SIGEF.csv")
        
        if not os.path.exists(av_f) or os.path.getsize(av_f) == 0:
            _baixar_drive(drive_av_id, av_f)
        if not os.path.exists(si_f) or os.path.getsize(si_f) == 0:
            _baixar_drive(drive_si_id, si_f)
        if drive_geral_id:
            geral_f = os.path.join(cache_dir, "geral.csv")
            if not os.path.exists(geral_f) or os.path.getsize(geral_f) == 0:
                try:
                    _baixar_drive(drive_geral_id, geral_f)
                except Exception:
                    pass


    else:


        # ── Modo pasta local ───────────────────────────────────────────────


        av_f = os.path.join(db_path, "avaliacoes.csv")


        si_f = os.path.join(db_path, "SIGEF.csv")


        if not os.path.exists(av_f): raise FileNotFoundError(f"Não encontrado: {av_f}")


        if not os.path.exists(si_f): raise FileNotFoundError(f"Não encontrado: {si_f}")


        


    df = _parse_csv(av_f, si_f)


    


    # Gera e substitui o Geral.xlsx na pasta local/servidor apenas se necessário


    try:


        local_dir = db_path if db_path else str(DADOS_DIR)


        geral_out = os.path.join(local_dir, "Geral.xlsx")


        


        build_needed = True


        if os.path.exists(geral_out) and os.path.exists(av_f):


            if os.path.getmtime(geral_out) >= os.path.getmtime(av_f):


                build_needed = False


                


        if build_needed:


            xlsx_bytes = _build_workbook(df, "GERAL — AADP 2026", df)


            with open(geral_out, "wb") as f_out:


                f_out.write(xlsx_bytes)


    except Exception:


        pass


        


    return df





def apply_filters(df, rpm_f, unid_f, sc_f, st_f, cert_f):


    if rpm_f:  df = df[df["Unidade RPM (Avaliado)"].isin(rpm_f)]


    if unid_f: df = df[df["Unidade Principal (Avaliado)"].isin(unid_f)]


    if sc_f:   df = df[df["Situação Comissão"].isin(sc_f)]


    if st_f:   df = df[df["Status Avaliação"].isin(st_f)]


    if cert_f: df = df[df["Certificação Homologador"].isin(cert_f)]


    return df





def fmt_num(n): return f"{n:,}".replace(",",".")





def df_to_xlsx(df: pd.DataFrame) -> bytes:
    import io
    output = io.BytesIO()
    
    is_audit_df = any("Aritm" in str(c) for c in df.columns) and any(str(c).endswith(" 1") for c in df.columns)
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if is_audit_df:
            try:
                styled_df = style_audit_dataframe(df)
                styled_df.to_excel(writer, index=False, sheet_name='Auditoria')
            except Exception:
                df.to_excel(writer, index=False, sheet_name='Planilha')
        else:
            df.to_excel(writer, index=False, sheet_name='Planilha')
            
    output.seek(0)
    return output.read()





MAX_STYLE = 4_000_000


def clean_none_values(df):
    if df is None or not hasattr(df, "columns"):
        return df
    import pandas as pd
    df = df.copy()
    df = df.fillna("-")
    for col in df.columns:
        if df[col].dtype == object or str(df[col].dtype) == "string":
            df[col] = df[col].apply(lambda x: "-" if str(x).strip().lower() in ("none", "nan", "<na>", "nat") else x)
            df[col] = df[col].replace({None: "-", "None": "-"})
    return df


def style_audit_dataframe(df):
    if df is None or not hasattr(df, "columns") or df.empty:
        return df
    col_media = next((c for c in df.columns if "Aritm" in str(c)), None)
    cols_av1 = [c for c in df.columns if str(c).endswith(" 1")]
    cols_av2 = [c for c in df.columns if str(c).endswith(" 2")]
    cols_av3 = [c for c in df.columns if str(c).endswith(" 3")]
    cols_av4 = [c for c in df.columns if str(c).endswith(" 4")]
    
    styles = {}
    if col_media:
        styles[col_media] = "background-color: #ffe599; color: black; font-weight: 500;" # Soft yellow
    for c in cols_av1:
        styles[c] = "background-color: #e2f0d9; color: #2d6a0f; font-weight: 500;" # Soft green
    for c in cols_av2:
        styles[c] = "background-color: #fce4d6; color: #7a3d00; font-weight: 500;" # Soft orange/peach
    for c in cols_av3:
        styles[c] = "background-color: #ebd9eb; color: #4a148c; font-weight: 500;" # Soft purple
    for c in cols_av4:
        styles[c] = "background-color: #e8f0fe; color: #1a0dab; font-weight: 500;" # Soft blue

    def get_column_styles(row):
        row_styles = []
        for col in row.index:
            if col == "Auditoria" and str(row.get("Auditoria", "")).strip() == "DIVERGENTE":
                row_styles.append("background-color: #ffcccc; color: #cc0000; font-weight: bold;")
            else:
                row_styles.append(styles.get(col, ""))
        return row_styles

    styler = df.style.apply(get_column_styles, axis=1)
    if col_media:
        def format_media(val):
            try:
                import math
                if val is None or val == "-":
                    return "-"
                f_val = float(str(val).replace(",", "."))
                if math.isnan(f_val):
                    return "-"
                return f"{f_val:.2f}"
            except Exception:
                return str(val)
        styler = styler.format(formatter={col_media: format_media})
    return styler


def safe_df(styled_or_df, height=520, key_prefix=None, show_download=False, download_name="dados_filtrados", download_type="csv", download_label="Baixar dados"): 
    """Exibe um DataFrame com st.dataframe nativo.
    - Ordenacao crescente/decrescente: clique no cabecalho de qualquer coluna.
    - Filtro rapido: campo de busca global acima da tabela.
    """
    import pandas as pd

    # Extrair DataFrame subjacente
    if hasattr(styled_or_df, "data"):
        styled_or_df.data = clean_none_values(styled_or_df.data)
        raw_df = styled_or_df.data.copy()
        is_styled = True
        is_large = raw_df.size > MAX_STYLE
    elif isinstance(styled_or_df, pd.DataFrame):
        raw_df = clean_none_values(styled_or_df)
        is_styled = False
        is_large = False
    else:
        raw_df = clean_none_values(styled_or_df)
        is_styled = False
        is_large = False

    if not isinstance(raw_df, pd.DataFrame) or raw_df.empty:
        st.info("\u2139\ufe0f Nenhum dado dispon\u00edvel para exibir.")
        return

    # Gerar chave unica para nao colidir entre telas
    if key_prefix is None:
        import hashlib
        key_prefix = "sdf_" + hashlib.md5(
            "".join(str(c) for c in raw_df.columns).encode()
        ).hexdigest()[:8]

    # Layout para campo de busca e botão de download
    st.markdown("<div style='margin-top: -15px; margin-bottom: -15px;'></div>", unsafe_allow_html=True)
    col_search, col_space, col_dl = st.columns([2.5, 0.1, 1], vertical_alignment="bottom")
    with col_search:
        busca = st.text_input(
            "🔍 Busca rápida (filtra qualquer coluna)",
            key=f"{key_prefix}_busca",
            placeholder="Digite o nr PM, Nome ou Unidade para filtrar",
            label_visibility="collapsed",
        )

    df_filtered = raw_df.copy()
    if busca and busca.strip():
        termo = busca.strip().lower()
        mask = df_filtered.apply(
            lambda col: col.astype(str).str.lower().str.contains(termo, na=False)
        ).any(axis=1)
        df_filtered = df_filtered[mask]

    if show_download:
        with col_dl:
            if download_type == "csv":
                d_file = df_filtered.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
                d_ext = "csv"
                d_mime = "text/csv"
            else:
                d_file = df_to_xlsx(df_filtered)
                d_ext = "xlsx"
                d_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            
            st.download_button(
                f"⬇️ {download_label}",
                d_file,
                f"{download_name}.{d_ext}",
                mime=d_mime,
                type="primary",
                use_container_width=True
            )

    total = len(raw_df)
    shown = len(df_filtered)
    limit = 500
    
    if shown > limit:
        st.caption(f"📊 Exibindo as primeiras **{limit}** de **{shown:,}** linhas encontradas (Total: {total:,} registros). Use o campo **Busca rápida** acima para filtrar.")
        df_to_show = df_filtered.head(limit)
    else:
        st.caption(f"📊 Exibindo **{shown:,}** de **{total:,}** registros.")
        df_to_show = df_filtered

    # Exibir com estilo quando possivel
    is_audit_df = any("Aritm" in str(c) for c in df_to_show.columns) and any(str(c).endswith(" 1") for c in df_to_show.columns)
    
    if is_audit_df:
        try:
            styled_df = style_audit_dataframe(df_to_show)
            st.dataframe(styled_df, use_container_width=True, height=height)
        except Exception:
            st.dataframe(df_to_show, use_container_width=True, height=height)
    elif is_styled and not is_large:
        try:
            st.dataframe(styled_or_df.data.loc[df_to_show.index], use_container_width=True, height=height)
        except Exception:
            st.dataframe(df_to_show, use_container_width=True, height=height)
    else:
        st.dataframe(df_to_show, use_container_width=True, height=height)


def color_status(val):


    m = {"Encerrada":"background-color:#e8f5e1;color:#2d6a0f;font-weight:600",


         "Homologação":"background-color:#fff8db;color:#7a5c00;font-weight:600",


         "Parcialmente Encerrada":"background-color:#fff0db;color:#7a3d00;font-weight:600",


         "Aberta":"background-color:#fde8e8;color:#8b0000;font-weight:600",
         "EM PRAZO DE RECURSO":"background-color:#f2f3f4;color:#5d6d7e;font-weight:600",
         "RECONSIDERAÇÃO COMISSÃO":"background-color:#f4ecf7;color:#6c3483;font-weight:600",
         "AUTORIDADE RECURSAL":"background-color:#f4ecf7;color:#6c3483;font-weight:600"}


    return m.get(val, "")





def color_sit(val):


    if val == "Comissão Atual":   return "background-color:#dce8f5;color:#1a3a6a;font-weight:600"


    if val == "Nota Provisória":  return "background-color:#fff9e6;color:#7a5c00;font-weight:600"


    return ""





# ─────────────────────── SEGURANÇA E AUTENTICAÇÃO ──────────────────────────────


if "authenticated" not in st.session_state:


    st.session_state.authenticated = False


    st.session_state.user_pm = ""


    st.session_state.user_name = ""


    st.session_state.user_role = ""


    st.session_state.user_rpm = ""


    st.session_state.user_unit = ""





if not st.session_state.authenticated:


    c1, c2, c3 = st.columns([1, 2, 1])


    with c2:


        if os.path.exists("logo_drh.png"):
            st.image("logo_drh.png", use_container_width=True)
        else:
            st.markdown("<div style='text-align: center; font-size: 4.5rem; margin-bottom: 25px;'>👮</div>", unsafe_allow_html=True)


        st.markdown("<h2 style='text-align: center; color: #9b8a5c; margin-top: -15px;'>Painel de Controle AADP</h2>", unsafe_allow_html=True)


        st.markdown("<h4 style='text-align: center; color: #a0a0a0; font-size: 0.95rem;'>Polícia Militar de Minas Gerais · Resolução 5458/2025</h4>", unsafe_allow_html=True)


        st.markdown("---")


        


        if "auth_mode" not in st.session_state:


            st.session_state.auth_mode = "🔑 Acessar Conta"


            


        auth_mode = st.radio("Selecione uma opção:", ["🔑 Acessar Conta", "📝 Solicitar Cadastro"], horizontal=True, key="auth_mode_radio")


        st.markdown("---")


        


        if auth_mode == "🔑 Acessar Conta":
            if st.session_state.get("forgot_password_mode", False):
                st.markdown("##### ❓ Recuperar Acesso (Esqueci minha senha)")
                
                if st.session_state.get("forgot_step2", False) and st.session_state.get("forgot_target_pm"):
                    target_pm = st.session_state.forgot_target_pm
                    
                    st.warning(
                        f"⚠️ **Atenção (PM: {target_pm})**:\n\n"
                        "Para recuperar seu acesso você precisa revogar o acesso e criar novo acesso.\n\n"
                        "Deseja prosseguir com a revogação do seu cadastro atual?"
                    )
                    
                    col_c1, col_c2 = st.columns(2)
                    with col_c1:
                        if st.button("Cancelar", use_container_width=True, key="btn_forgot_cancel"):
                            st.session_state.forgot_password_mode = False
                            st.session_state.forgot_step2 = False
                            st.session_state.forgot_target_pm = None
                            st.rerun()
                    with col_c2:
                        if st.button("Revogar", use_container_width=True, type="primary", key="btn_forgot_revoke"):
                            db_revoke_user(target_pm)
                            log_action(target_pm, "REVOGAR_AUTOCADASTRO", "Usuario solicitou revogacao por esquecimento de senha")
                            st.success("✅ Seu cadastro foi revogado com sucesso! Agora você pode solicitar um novo cadastro.")
                            st.session_state.forgot_password_mode = False
                            st.session_state.forgot_step2 = False
                            st.session_state.forgot_target_pm = None
                            st.rerun()
                else:
                    st.write("Informe o seu Nº PM para verificar seu cadastro:")
                    forgot_pm = st.text_input("Nº PM:", key="forgot_pm_input", placeholder="Ex: 123456")
                    
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        if st.button("Voltar ao Login", use_container_width=True, key="btn_forgot_back"):
                            st.session_state.forgot_password_mode = False
                            st.rerun()
                    with col_f2:
                        if st.button("Verificar Cadastro", use_container_width=True, type="primary", key="btn_forgot_verify"):
                            if not forgot_pm or not forgot_pm.isdigit():
                                st.error("Por favor, informe um Nº PM válido (apenas números).")
                            else:
                                f_pm = forgot_pm.strip()
                                status = db_check_user_status(f_pm)
                                if not status:
                                    st.error("❌ Nº PM não encontrado no sistema!")
                                elif status == "Bloqueado":
                                    st.warning("⚠️ Seu acesso já está revogado. Você pode ir para a opção '📝 Solicitar Cadastro' e realizar seu novo cadastro.")
                                else:
                                    st.session_state.forgot_target_pm = f_pm
                                    st.session_state.forgot_step2 = True
                                    st.rerun()
            else:
                with st.form("form_login", clear_on_submit=False):
                    login_pm = st.text_input("Nº PM:", key="login_pm_val", placeholder="Ex: 123456 ou ADM")
                    login_pass = st.text_input("Senha:", type="password", key="login_pass_val")
                    submitted_login = st.form_submit_button("Entrar", use_container_width=True, type="primary")

                if submitted_login:
                    spm = login_pm.strip()
                    spass = login_pass
                    
                    if not spm or not spass:
                        st.error("Por favor, preencha todos os campos.")
                    else:
                        h_pass = hashlib.sha256(spass.encode()).hexdigest()
                        res = db_get_user_for_login(spm, h_pass)
                        if res:
                            name, role, rpm, unit, status = res
                            if status == "Ativo":
                                st.session_state.authenticated = True
                                st.session_state.selected_year = None
                                st.session_state.user_pm = spm
                                st.session_state.user_name = name
                                if role in ("P1/SADM", "P1"):
                                    role = "P1"
                                elif role in ("DRH6", "Gestor"):
                                    role = "Gestor"
                                st.session_state.user_role = role
                                st.session_state.user_rpm = rpm
                                st.session_state.user_unit = unit
                                refresh_db_cache()
                                log_action(spm, "LOGIN", "Acesso realizado com sucesso")
                                st.success(f"Bem-vindo, {name}!")
                                st.rerun()
                            elif status == "Pendente":
                                st.warning("⚠️ Sua conta está aguardando liberação do Administrador.")
                            else:
                                st.error("❌ Acesso revogado/bloqueado. Entre em contato com a DRH.")
                        else:
                            st.error("❌ Nº PM ou Senha incorretos.")

                st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
                if st.button("❓ Esqueci minha senha", use_container_width=True, key="btn_forgot_password"):
                    st.session_state.forgot_password_mode = True
                    st.session_state.forgot_step2 = False
                    st.session_state.forgot_target_pm = None
                    st.rerun()


                        


        else:


            st.markdown("##### 📝 Solicitação de acesso - Painel de Controle AADP")


            st.info("⚠️ Informe apenas os **6 primeiros dígitos** do seu Nº PM (sem o dígito verificador).\n\n🔒 **Importante:** Somente Comandante e Subcomandante de unidade, militares da P1 e SADM poderão solicitar e ter autorizado o acesso a este Painel de controle AADP.")


            


            # Inicializa variáveis de estado


            if "sigef_data" not in st.session_state:


                st.session_state.sigef_data = None


            if "sigef_verified" not in st.session_state:


                st.session_state.sigef_verified = False


                


            # Entrada de Nº PM


            reg_pm = st.text_input("Nº PM (Apenas os 6 primeiros dígitos):", max_chars=6, key="reg_pm", placeholder="Ex: 053108")


            


            col_cons, col_clear = st.columns([3, 1])


            with col_cons:


                if st.button("🔍 Consultar", use_container_width=True, type="secondary"):


                    if not reg_pm or not reg_pm.isdigit():


                        st.error("Por favor, informe um Nº PM válido (apenas números, máximo 6 dígitos).")


                        st.session_state.sigef_data = None


                        st.session_state.sigef_verified = False


                    else:


                        with st.spinner("⏳ Consultando banco de dados do SIGEF..."):


                            res = find_sigef_user(reg_pm)


                            if res:


                                st.session_state.sigef_data = res


                                st.session_state.sigef_verified = False # Reseta a verificação para nova busca


                                st.success("✅ Militar encontrado no SIGEF! Prossiga com a verificação de segurança abaixo.")


                            else:


                                st.error("❌ Nº PM não encontrado no banco SIGEF. Verifique se digitou os 6 primeiros dígitos corretamente.")


                                st.session_state.sigef_data = None


                                st.session_state.sigef_verified = False


            with col_clear:


                if st.button("🧹 Limpar", use_container_width=True):


                    st.session_state.sigef_data = None


                    st.session_state.sigef_verified = False


                    st.rerun()





            if st.session_state.sigef_data:


                data = st.session_state.sigef_data


                


                # Etapa 2: Verificação de Segurança (CPF e Nascimento)


                if not st.session_state.sigef_verified:


                    st.markdown("---")


                    st.markdown("##### 🔒 Verificação de Segurança")


                    st.write("Para confirmar se realmente é você, confirme as duas informações abaixo:")


                    


                    v_cpf = st.text_input("Digite seu CPF (apenas números):", max_chars=11, key="v_cpf", placeholder="Ex: 12345678901")


                    v_birth = st.text_input("Digite sua Data de Nascimento (DD/MM/AAAA):", max_chars=10, key="v_birth", placeholder="Ex: 25/09/1957")


                    


                    if st.button("Confirmar Dados", use_container_width=True, type="primary"):


                        # Normaliza CPF: remove pontuação e também todos os zeros à esquerda


                        clean_input_cpf = re.sub(r'\D', '', v_cpf).lstrip('0')


                        clean_sigef_cpf = re.sub(r'\D', '', data["cpf"]).lstrip('0')


                        


                        # Normaliza Data de Nascimento: remove qualquer caractere que não seja número (ex: barras)


                        clean_input_birth = re.sub(r'\D', '', v_birth)


                        clean_sigef_birth = re.sub(r'\D', '', data["birthdate"])


                        


                        if clean_input_cpf == clean_sigef_cpf and clean_input_birth == clean_sigef_birth:


                            st.session_state.sigef_verified = True


                            st.success("✅ Identidade confirmada com sucesso!")


                            st.rerun()


                        else:


                            st.error("❌ CPF ou Data de Nascimento incorretos. Verifique suas informações e tente novamente.")


                


                # Etapa 3: Liberação do Formulário de Senha


                if st.session_state.sigef_verified:


                    st.markdown("---")


                    st.markdown("##### 👤 Dados Funcionais Confirmados:")


                    st.text_input("Posto/Graduação:", value=data["rank"], disabled=True, key="disp_rank")


                    st.text_input("Nome Completo:", value=data["name"], disabled=True, key="disp_name")


                    st.text_input("UDI/UDG:", value=data["rpm"], disabled=True, key="disp_rpm")


                    st.text_input("Unidade Principal:", value=data["unit"], disabled=True, key="disp_unit")


                    st.text_input("Setor:", value=data["sector"], disabled=True, key="disp_sector")


                    


                    st.markdown("##### 🔑 Configuração de Senha de Acesso:")


                    with st.form("form_cadastro_final", clear_on_submit=False):
                        st.warning("⚠️ **Segurança:** Por motivos de segurança, a senha de acesso cadastrada **NÃO** deve ser igual à sua senha da **IntranetPM** ou do **SIRH**.")
                        reg_pass = st.text_input("Escolha uma Senha:", type="password", key="reg_pass")


                        reg_pass_conf = st.text_input("Confirme a Senha:", type="password", key="reg_pass_conf")


                        


                        submitted = st.form_submit_button("Enviar Solicitação", use_container_width=True, type="primary")


                        


                    if submitted:


                        spass = reg_pass


                        if not spass:


                            st.error("Por favor, preencha o campo de senha.")


                        elif spass != reg_pass_conf:


                            st.error("As senhas não coincidem!")


                        elif len(spass) < 6:


                            st.error("A senha deve ter pelo menos 6 caracteres.")


                        else:


                            if st.session_state.get(f"registered_{data['pm']}", False):
                                st.error("❌ Solicitação já processada. Por favor, retorne à tela de Login.")
                            else:
                                st.session_state[f"registered_{data['pm']}"] = True
                                try:


                                    current_status = db_check_user_status(data["pm"])


                                    if current_status:


                                        if current_status in ("Pendente", "Ativo"):


                                            st.error("❌ Este Nº PM já possui solicitação de acesso ativa ou pendente no sistema!")


                                        else:
                                            # Usuário Bloqueado ou Recusado - pode solicitar novamente
                                            h_pass = hashlib.sha256(spass.encode()).hexdigest()
                                            db_re_request_access(data["pm"], data["name"], data["rank"], data["rpm"], data["unit"], data["sector"], h_pass)
                                            log_action(data["pm"], "RE_CADASTRO_SOLICITADO", f"Nome: {data['name']}, Posto: {data['rank']}")
                                            auto_role_check = _auto_detect_role(data["sector"])
                                            if auto_role_check:
                                                st.success(f"✅ Cadastro aprovado automaticamente! Seu perfil **{auto_role_check}** foi liberado. Você já pode fazer login.")
                                            else:
                                                st.success("✅ Nova solicitação enviada com sucesso! Aguarde a liberação do Administrador.")
                                            st.session_state.sigef_data = None
                                            st.session_state.sigef_verified = False
                                    else:


                                        h_pass = hashlib.sha256(spass.encode()).hexdigest()


                                        db_create_new_request(data["pm"], data["name"], data["rank"], data["rpm"], data["unit"], data["sector"], h_pass)


                                        log_action(data["pm"], "CADASTRO_SOLICITADO", f"Nome: {data['name']}, Posto: {data['rank']}, UDI/UDG: {data['rpm']}")


                                        auto_role_check = _auto_detect_role(data["sector"])
                                        if auto_role_check:
                                            st.success(f"✅ Cadastro aprovado automaticamente! Seu perfil **{auto_role_check}** foi liberado. Você já pode fazer login.")
                                        else:
                                            st.success("✅ Solicitação enviada com sucesso! Aguarde a liberação do Administrador.")


                                        st.session_state.sigef_data = None


                                        st.session_state.sigef_verified = False


                                except Exception as e:


                                    st.error(f"Erro ao salvar cadastro: {str(e)}")

    st.stop()


# ─────────────────────── TELA DE ESCOLHA DO ANO AVALIATIVO ────────────────────
if not st.session_state.get("selected_year"):
    c_y1, c_y2, c_y3 = st.columns([1, 2.2, 1])
    with c_y2:
        if os.path.exists("logo_drh.png"):
            st.image("logo_drh.png", use_container_width=True)
        else:
            st.markdown("<div style='text-align: center; font-size: 4rem; margin-bottom: 20px;'>🏛️</div>", unsafe_allow_html=True)

        st.markdown("<h2 style='text-align: center; color: #9b8a5c; margin-top: -10px;'>Painel de Controle AADP</h2>", unsafe_allow_html=True)
        st.markdown("<h4 style='text-align: center; color: #a0a0a0; font-size: 0.95rem;'>Selecione o Ciclo Avaliativo para Prosseguir</h4>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align: center; color: #d4af37;'>Militar conectado: <b>{st.session_state.get('user_name', '')}</b> ({st.session_state.get('user_role', '')})</p>", unsafe_allow_html=True)
        st.markdown("---")

        st.markdown("#### 📅 Escolha o Ano de Avaliação:")
        col_btn26, col_btn27 = st.columns(2)
        with col_btn26:
            if st.button("📅 AADP 2026\n\n(Ciclo 2025/2026)", use_container_width=True, type="primary", key="btn_ano_2026"):
                st.session_state.selected_year = "2026"
                st.cache_data.clear()
                log_action(st.session_state.user_pm, "SELECAO_ANO", "Ano avaliativo 2026 selecionado")
                st.rerun()
        with col_btn27:
            if st.button("📅 AADP 2027\n\n(Ciclo 2026/2027)", use_container_width=True, type="primary", key="btn_ano_2027"):
                st.session_state.selected_year = "2027"
                st.cache_data.clear()
                log_action(st.session_state.user_pm, "SELECAO_ANO", "Ano avaliativo 2027 selecionado")
                st.rerun()

        st.markdown("---")
        if st.button("🚪 Sair / Logoff", use_container_width=True, key="btn_logoff_year_select"):
            log_action(st.session_state.user_pm, "LOGOFF", "Saída na tela de seleção de ano")
            st.session_state.authenticated = False
            st.session_state.selected_year = None
            st.rerun()

    st.stop()


# ─────────────────────── SIDEBAR ──────────────────────────────────────────────


with st.sidebar:


    st.image("logo_drh.png", use_container_width=True)


    active_year = str(st.session_state.get("selected_year", "2026"))
    st.markdown(f"### AADP {active_year}")
    st.markdown("**Sistema de Análise de Avaliações**")
    st.info(f"📅 **Ciclo Ativo:** AADP {active_year}")
    if st.button("🔄 Trocar Ano Avaliativo", use_container_width=True, key="btn_sidebar_trocar_ano", help="Clique para retornar à tela de escolha do ano"):
        st.session_state.selected_year = None
        st.cache_data.clear()
        st.rerun()


    


    # Determina o perfil ativo (real vs simulado) para ajustar as opções da barra lateral


    sidebar_active_role = st.session_state.get("simulated_role", st.session_state.user_role) if st.session_state.get("simulation_active", False) else st.session_state.user_role


    


    # Exibe informações do militar


    st.markdown(f"<small>👤 <b>Militar:</b> {st.session_state.user_name} ({st.session_state.user_pm})</small>", unsafe_allow_html=True)


    if st.session_state.get("simulation_active", False) and st.session_state.user_role == "ADMINISTRADOR":


        st.markdown(f"<small>🔑 <b>Perfil Real:</b> ADMINISTRADOR</small>", unsafe_allow_html=True)


        st.markdown(f"<small>🕵️ <b>Simulado:</b> <span style='color:#ff9f43;'>{st.session_state.simulated_role}</span></small>", unsafe_allow_html=True)


        if st.button("Voltar simulação", key="sidebar_stop_sim", type="primary", use_container_width=True):
            st.session_state.simulation_active = False
            st.session_state.simulated_pm = ""
            st.session_state.simulated_name = ""
            st.session_state.simulated_role = ""
            st.session_state.simulated_rpm = ""
            st.session_state.simulated_unit = ""
            st.session_state.active_page = "Painel Administrador"
            log_action("ADM", "ENCERRAR_SIMULACAO", "Simulacao desativada")
            st.rerun()


    else:


        st.markdown(f"<small>🔑 <b>Perfil:</b> <span style='color:#9b8a5c;'>{sidebar_active_role}</span></small>", unsafe_allow_html=True)


        


    st.markdown("---")
    # Funcionalidade suspensa temporariamente:
    # global_aadp = st.radio("Grupo AADP", ["Ativa", "Reconduzido", "Discente"], index=0, key="global_aadp")
    global_aadp = "Ativa"
    st.markdown("---")





    # 1. Mostrar Filtros (Primeiro)


    if "show_filtros" not in st.session_state:


        st.session_state.show_filtros = False


        


    btn_filtros_label = "🔍 Ocultar Filtros" if st.session_state.show_filtros else "🔍 Mostrar Filtros"


    btn_filtros_type = "primary" if st.session_state.show_filtros else "secondary"


    if st.button(btn_filtros_label, use_container_width=True, key="btn_toggle_filtros", type=btn_filtros_type):


        st.session_state.show_filtros = not st.session_state.show_filtros


        st.rerun()





    container_filtros = st.container()


    st.markdown("---")





    # 2. Páginas / Navegação (Segundo)


    st.markdown("#### 🧭 Páginas")


    pages = []
    
    if sidebar_active_role.upper() in ("ADMINISTRADOR", "GESTOR", "P1", "SADM"):
        pages.append(("🚨 Análise de Comissões", "Comissões"))
        pages.append(("🎯 Controle do CDP", "Controle do CDP"))

    pages.extend([
        ("📊 Análise Gráfica", "Análise Gráfica"),
        ("📋 Dados Gerais", "Dados Gerais"),
        ("⏳ Avaliações Pendentes", "Avaliações Pendentes"),
        ("👥 Avaliadores Pendentes", "Avaliadores Pendentes"),
    ])

    # P1 e SADM não possuem acesso às opções de exportação/relatórios
    if sidebar_active_role not in ("P1", "SADM"):
        pages.append(("📥 Gerar Relatório", "Gerar Relatório"))
    if sidebar_active_role not in ("SADM",):
        pages.append(("📄 Relatório Word", "Relatório Word"))

    # Auditoria de Notas: visível para ADMINISTRADOR, GESTOR, P1 e SADM
    if sidebar_active_role.upper() in ("ADMINISTRADOR", "GESTOR", "P1", "SADM"):
        pages.append(("📊 Auditoria de Notas", "Auditoria de Notas"))

    if sidebar_active_role.upper() in ("ADMINISTRADOR", "GESTOR", "P1", "SADM"):
        pages.append(("📊 Dados Consolidados", "Dados Consolidados"))

    # O administrador real sempre vê o painel administrador


    if st.session_state.user_role == "ADMINISTRADOR":


        pending_count = db_get_pending_count()


            


        if pending_count > 0:


            pages.append((f"⚙️ Painel Administrador (🔴 :red[{pending_count}])", "Painel Administrador"))


        else:


            pages.append(("⚙️ Painel Administrador", "Painel Administrador"))





    if "active_page" not in st.session_state:


        st.session_state.active_page = "Análise Gráfica"


        


    if st.session_state.active_page == "Controle do CDP" and sidebar_active_role.upper() not in ("ADMINISTRADOR", "GESTOR", "P1", "SADM"):
        st.session_state.active_page = "Análise Gráfica"

    if st.session_state.active_page == "Painel Administrador" and st.session_state.user_role != "ADMINISTRADOR":


        st.session_state.active_page = "Análise Gráfica"


        


    if st.session_state.active_page == "Gerar Relatório" and sidebar_active_role in ("P1", "SADM"):
        st.session_state.active_page = "Análise Gráfica"
    elif st.session_state.active_page == "Relatório Word" and sidebar_active_role in ("SADM",):
        st.session_state.active_page = "Análise Gráfica"





    for label, page_name in pages:


        is_active = st.session_state.active_page == page_name


        btn_type = "primary" if is_active else "secondary"


        if st.button(label, key=f"nav_{page_name}", use_container_width=True, type=btn_type):
            st.session_state.active_page = page_name
            st.rerun()





    # Inicializa variáveis para não dar NameError


    active_year = str(st.session_state.get("selected_year", "2026"))
    year_cfg = get_active_year_config(active_year, cfg)
    drive_av_id = year_cfg["drive_av_id"]
    drive_si_id = year_cfg["drive_si_id"]
    drive_geral_id = year_cfg["drive_geral_id"]
    db_path = year_cfg["db_path"]
    fonte = cfg.get("fonte_dados", "📁 Pasta local / Servidor")


    reload = False





    # Recarregar Dados para Administrador


    if st.session_state.user_role == "ADMINISTRADOR":


        st.markdown("---")


        if st.button("🔄 Recarregar Dados", use_container_width=True, type="primary", key="btn_reload"):
            st.cache_data.clear()
            st.cache_resource.clear()
            
            # Limpar arquivos baixados para forçar download novo
            import tempfile
            import shutil
            for y_c in ["", "_2026", "_2027"]:
                c_dir = os.path.join(tempfile.gettempdir(), f"aadp_drive_cache{y_c}")
                if os.path.exists(c_dir):
                    try: shutil.rmtree(c_dir)
                    except: pass
            
            # Limpar planilha mestre
            local_master = os.path.join(str(DADOS_DIR), "Analise avaliacoes completa.xlsx")
            if os.path.exists(local_master):
                try: os.remove(local_master)
                except: pass
                
            local_master_parent = os.path.join(str(Path(DADOS_DIR).parent), "Analise avaliacoes completa.xlsx")
            if os.path.exists(local_master_parent):
                try: os.remove(local_master_parent)
                except: pass

            if "db_users" in st.session_state:
                del st.session_state.db_users
            if "db_logs" in st.session_state:
                del st.session_state.db_logs
            st.success("Dados recarregados com sucesso!")
            st.rerun()





    # Botões de Alterar Senha e Sair/Logoff no final da barra lateral


    st.markdown("---")


    if st.button("🔑 Alterar Senha", use_container_width=True, key="btn_toggle_change_password"):


        st.session_state.show_change_password = not st.session_state.get("show_change_password", False)


        st.rerun()





    if st.button("🚪 Sair / Logoff", use_container_width=True, key="btn_logoff"):


        log_action(st.session_state.user_pm, "LOGOFF", "Saída voluntária")


        st.session_state.authenticated = False
        st.session_state.selected_year = None
        st.session_state.user_pm = ""


        st.session_state.user_name = ""


        st.session_state.user_role = ""


        st.session_state.user_rpm = ""


        st.session_state.user_unit = ""


        st.session_state.simulation_active = False


        st.session_state.simulated_pm = ""


        st.session_state.simulated_name = ""


        st.session_state.simulated_role = ""


        st.session_state.simulated_rpm = ""


        st.session_state.simulated_unit = ""


        st.rerun()

# ─────────────────────── CARREGAR DADOS ───────────────────────────────────────


# Calcula as variáveis ativas considerando simulação


if st.session_state.get("simulation_active", False):


    active_role = st.session_state.get("simulated_role", st.session_state.user_role)


    active_rpm = st.session_state.get("simulated_rpm", st.session_state.user_rpm)


    active_unit = st.session_state.get("simulated_unit", st.session_state.user_unit)


    active_pm = st.session_state.get("simulated_pm", st.session_state.user_pm)


    active_name = st.session_state.get("simulated_name", st.session_state.user_name)


else:


    active_role = st.session_state.user_role


    active_rpm = st.session_state.user_rpm


    active_unit = st.session_state.user_unit


    active_pm = st.session_state.user_pm


    active_name = st.session_state.user_name





try:


    if reload: st.cache_data.clear()


    av_csv_path = os.path.join(db_path or cfg.get("db_path", str(DADOS_DIR)), "avaliacoes.csv")


    last_mod_str = get_last_updated_time(av_csv_path, drive_av_id or cfg.get("drive_av_id", ""))


    


    df_full = load_data(
        db_path   = year_cfg["db_path"],
        drive_av_id = year_cfg["drive_av_id"],
        drive_si_id = year_cfg["drive_si_id"],
        drive_geral_id = year_cfg["drive_geral_id"],
        ano = active_year
    )


    if active_role == "P1":


        df_full = df_full[df_full["Unidade RPM (Avaliado)"].apply(lambda x: matches_rpm(active_rpm, x))]


    elif active_role == "SADM":


        df_full = df_full[df_full["Unidade Principal (Avaliado)"].str.strip().str.upper() == active_unit.strip().upper()]


        


    data_ok = True; ts = now_br().strftime("%d/%m/%Y %H:%M")


except Exception as e:


    data_ok = False; err_msg = str(e); last_mod_str = "Data/Hora indisponível"





# ─────────────────────── FILTROS ──────────────────────────────────────────────


rpm_filter = unid_filter = sit_com_filter = status_filter = cert_filter = []


if data_ok:


    all_rpm   = sorted(df_full["Unidade RPM (Avaliado)"].dropna().unique(), key=rpm_sort_key)


    all_status= ["Aberta", "Parcialmente Encerrada", "Homologação", "Encerrada", "EM PRAZO DE RECURSO", "RECONSIDERAÇÃO COMISSÃO", "AUTORIDADE RECURSAL"]


    all_sit   = ["Comissão Atual","Nota Provisória"]


    all_cert  = ["SIM","NÃO","-"]


    with container_filtros:


        if st.session_state.show_filtros:


            st.markdown("#### 🔍 Filtros de Visualização")


            if active_role not in ("P1", "SADM"):


                rpm_filter = st.multiselect("🏢 Unidade RPM", all_rpm, placeholder="Todas")


            else:


                rpm_filter = []


            df_tmp = df_full[df_full["Unidade RPM (Avaliado)"].isin(rpm_filter)] if rpm_filter else df_full


            all_unid = sorted(df_tmp["Unidade Principal (Avaliado)"].dropna().unique())


            


            if active_role != "SADM":


                unid_filter = st.multiselect("🏛️ Subunidade", all_unid, placeholder="Todas")


            else:


                unid_filter = []


                


            st.markdown("")


            sit_com_filter = st.multiselect("🔵 Situação Comissão", all_sit, placeholder="Todas")


            status_filter  = st.multiselect("📊 Status",            all_status, placeholder="Todos")


            cert_filter    = st.multiselect("✅ Cert. Homologador",  all_cert,   placeholder="Todos")


            st.markdown("---")


            st.markdown(f"<small>🕐 Carregado: {ts}</small>", unsafe_allow_html=True)


            st.markdown(f"<small>📊 {fmt_num(len(df_full))} registros</small>", unsafe_allow_html=True)


    df = apply_filters(df_full, rpm_filter, unid_filter, sit_com_filter, status_filter, cert_filter)
    
    PAGES_WITH_FILTER = [
        "Análise Gráfica", "Dados Gerais", "Avaliações Pendentes", 
        "Avaliadores Pendentes", "Gerar Relatório", "Relatório Word"
    ]
    if st.session_state.get("active_page", "") in PAGES_WITH_FILTER:
        df_full = df_full[df_full["Sit. Funcional"].isin(SITUACOES_ALVO)].copy()
        df = df[df["Sit. Funcional"].isin(SITUACOES_ALVO)].copy()


else:


    df = pd.DataFrame()





# ─────────────────────── CABEÇALHO ────────────────────────────────────────────


logo_base64 = ""


if os.path.exists("logo_drh.png"):


    import base64


    with open("logo_drh.png", "rb") as f:


        logo_base64 = base64.b64encode(f.read()).decode("utf-8")





logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="width: 100%; max-width: 480px; height: auto; max-height: 120px; object-fit: contain; border-radius: 8px; align-self: center;" />' if logo_base64 else ""





st.markdown(f'<div class="main-title">{logo_html}<div style="margin-top: 10px;"><h1 style="font-size: 2.3rem; margin: 0; font-weight: 800; color: #9b8a5c; text-transform: uppercase;">Painel de Controle AADP</h1><p style="font-size: 1.05rem; margin: 5px 0 0 0; color: #e5dccb; font-weight: 500;">Polícia Militar de Minas Gerais · Resolução 5458/2025</p><p style="font-size: 0.9rem; margin-top: 8px; color: #a0a0a0; font-style: italic;">Dados consolidados em {last_mod_str}</p></div></div>', unsafe_allow_html=True)





# --- FORMULÁRIO DE ALTERAÇÃO DE SENHA ---
if st.session_state.get("show_change_password", False):
    st.markdown('<div class="info-box" style="border-left: 5px solid #9b8a5c;">🛡️ <b>Alterar Senha do Usuário</b></div>', unsafe_allow_html=True)
    with st.form("form_change_password", clear_on_submit=True):
        curr_pw = st.text_input("Senha Atual:", type="password", key="chg_curr_pw")
        new_pw = st.text_input("Nova Senha:", type="password", key="chg_new_pw")
        conf_pw = st.text_input("Confirmar Nova Senha:", type="password", key="chg_conf_pw")
        submit_chg = st.form_submit_button("💾 Atualizar Senha", use_container_width=True, type="primary")
            
    if submit_chg:
        if not curr_pw or not new_pw or not conf_pw:
            st.error("❌ Por favor, preencha todos os campos.")
        elif new_pw != conf_pw:
            st.error("❌ A nova senha e a confirmação não coincidem.")
        else:
            h_curr = hashlib.sha256(curr_pw.encode()).hexdigest()
            row_pw = db_get_user_password(st.session_state.user_pm)
            if not row_pw or row_pw != h_curr:
                st.error("❌ Senha atual incorreta.")
            else:
                h_new = hashlib.sha256(new_pw.encode()).hexdigest()
                db_update_password(st.session_state.user_pm, h_new)
                log_action(st.session_state.user_pm, "ALTERAR_SENHA", "Senha alterada com sucesso pelo proprio usuario")
                st.success("✅ Senha alterada com sucesso!")
                st.session_state.show_change_password = False
                st.rerun()
                
    if st.button("❌ Cancelar / Fechar", use_container_width=True, key="btn_cancel_change_pw"):
        st.session_state.show_change_password = False
        st.rerun()
    st.markdown("---")


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
n_recurso = df["Status Avaliação"].isin([
    "RECONSIDERAÇÃO COMISSÃO", "AUTORIDADE RECURSAL", "EM PRAZO DE RECURSO"
]).sum()
n_ca     = (df["Situação Comissão"]=="Comissão Atual").sum()
n_np     = (df["Situação Comissão"]=="Nota Provisória").sum()



if st.session_state.get("active_page", "Análise Gráfica") == "Análise Gráfica":
    col_block1, col_block2 = st.columns([1, 1.25], gap="large")

    with col_block1:
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
                        f'<div class="sub">{n_ca/max(n_total,1)*100:.2f}%</div>'
                        '</div>', unsafe_allow_html=True)
        with cb1_2:
            st.markdown('<div class="kpi-card kpi-np">'
                        '<div class="label">NOTA PROVISÓRIA</div>'
                        f'<div class="value">{fmt_num(n_np)}</div>'
                        f'<div class="sub">{n_np/max(n_total,1)*100:.2f}%</div>'
                        '</div>', unsafe_allow_html=True)

    with col_block2:
        st.markdown('<div class="kpi-card kpi-enc">'
                    '<div class="label">ENCERRADAS</div>'
                    f'<div class="value">{fmt_num(n_enc)}</div>'
                    f'<div class="sub">{n_enc/max(n_total,1)*100:.2f}%</div>'
                    '</div>', unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)

        cb2_1, cb2_2, cb2_3, cb2_4 = st.columns(4)
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
        with cb2_4:
            st.markdown('<div class="kpi-card kpi-recurso">'
                        '<div class="label">EM RECURSO</div>'
                        f'<div class="value">{fmt_num(n_recurso)}</div>'
                        '<div class="sub">Recurso pendente</div>'
                        '</div>', unsafe_allow_html=True)







st.markdown("<div class='main-nav-marker' style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

# ─────────────────────── HORIZONTAL NAVIGATION TABS ──────────────────────────
main_active_role = st.session_state.get("simulated_role", st.session_state.user_role) if st.session_state.get("simulation_active", False) else st.session_state.user_role

main_nav_pages = []

if main_active_role.upper() in ("ADMINISTRADOR", "GESTOR", "P1", "SADM"):
    main_nav_pages.append(("🚨\nAnálise de Comissões", "Comissões"))
    main_nav_pages.append(("🎯\nControle do CDP", "Controle do CDP"))

main_nav_pages.extend([
    ("📊\nAnálise Gráfica", "Análise Gráfica"),
    ("📋\nDados Gerais", "Dados Gerais"),
    ("⏳\nAvaliações Pendentes", "Avaliações Pendentes"),
    ("👥\nAvaliadores Pendentes", "Avaliadores Pendentes"),
])

if main_active_role not in ("P1", "SADM"):
    main_nav_pages.append(("📥\nGerar Relatório", "Gerar Relatório"))
    main_nav_pages.append(("📄\nRelatório Word", "Relatório Word"))

if main_active_role.upper() in ("ADMINISTRADOR", "GESTOR", "P1", "SADM"):
    main_nav_pages.append(("📊\nAuditoria de Notas", "Auditoria de Notas"))

if main_active_role.upper() in ("ADMINISTRADOR", "GESTOR", "P1", "SADM"):
    main_nav_pages.append(("📊\nDados Consolidados", "Dados Consolidados"))

if st.session_state.user_role == "ADMINISTRADOR":
    p_count = db_get_pending_count()
    if p_count > 0:
        main_nav_pages.append((f"⚙️\nPainel Administrador ({p_count} 🔴)", "Painel Administrador"))
    else:
        main_nav_pages.append(("⚙️\nPainel Administrador", "Painel Administrador"))

# Render as horizontal buttons in columns
num_tabs = len(main_nav_pages)
tab_cols = st.columns(num_tabs)
for idx, (label, page_name) in enumerate(main_nav_pages):
    with tab_cols[idx]:
        is_active = (st.session_state.active_page == page_name)
        btn_type = "primary" if is_active else "secondary"
        if st.button(label, key=f"main_tab_{page_name}", use_container_width=True, type=btn_type):
            st.session_state.active_page = page_name
            st.rerun()

st.markdown("<div style='margin-bottom: 25px; border-bottom: 1px solid rgba(255,255,255,0.08);'></div>", unsafe_allow_html=True)





# ─────────────────────── SELEÇÃO DE ABAS VIA SESSION STATE ───────────────────


if "active_page" not in st.session_state:


    st.session_state.active_page = "Análise Gráfica"





active_page = st.session_state.active_page




# TAB 1 — ANÁLISE GRÁFICA


# ══════════════════════════════════════════════════════════════════════════════


if active_page == "Análise Gráfica":


    st.markdown("### 📊 Análise Gráfica das Avaliações")





    # ── LINHA 1: Pizza de Status (destaque, full width) ──────────────────────


    st.markdown("---")


    ordered_labels = ["Aberta", "Parcialmente Encerrada", "Homologação", "Encerrada", "EM PRAZO DE RECURSO", "RECONSIDERAÇÃO COMISSÃO", "AUTORIDADE RECURSAL"]


    sd = df.groupby("Status Avaliação").size()


    vals_pizza  = [int(sd.get(s, 0)) for s in ordered_labels]


    cols_pizza  = [STATUS_COLORS[s] for s in ordered_labels]





    fig_status = go.Figure()





    # Sombra para efeito 3D (circle com proporção assimétrica = elipse)


    fig_status.add_shape(type="circle", xref="paper", yref="paper",


        x0=0.12, y0=0.01, x1=0.88, y1=0.10,


        fillcolor="rgba(0,0,0,0.18)", line_color="rgba(0,0,0,0)", layer="below")





    fig_status.add_trace(go.Pie(
        labels=ordered_labels, values=vals_pizza,
        hole=0.54,
        pull=[0.09, 0.06, 0.04, 0],
        texttemplate="<b>%{label}</b><br>%{value:,} (%{percent:.2%})",
        textposition="outside",
        rotation=90,  # Solicitado pelo usuário
        textfont=dict(size=13, family="Inter, sans-serif"),
        insidetextorientation="radial",
        marker=dict(colors=cols_pizza, line=dict(color="#121212", width=3)),
        hovertemplate="<b>%{label}</b><br>Avaliações: <b>%{value:,}</b><br>%{percent:.2%}<extra></extra>",
        sort=False,
    ))

    fig_status.add_annotation(
        text=f"<b>{fmt_num(n_total)}</b><br><span style='font-size:11px;color:#a0a0a0'>avaliações</span>",
        x=0.5, y=0.5, showarrow=False,
        font=dict(size=22, color="#9b8a5c", family="Inter"),
        align="center",
    )

    fig_status.update_layout(
        template="plotly_dark",
        title=dict(text="<b>Status das Avaliações — AADP 2026</b>",
                   font=dict(size=20, color="#9b8a5c"), x=0.5, y=0.96),
        height=500, showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(t=60, b=80, l=120, r=120),
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


            texttemplate="<b>%{label}</b><br>%{value:,} (%{percent:.2%})", textposition="outside",


            textfont=dict(size=12), sort=False,


            marker=dict(colors=[SIT_COLORS.get(s,"#aaa") for s in sit_d["Situação Comissão"]],


                        line=dict(color="#121212", width=3)),


            hovertemplate="<b>%{label}</b><br>%{value:,} avaliações (%{percent:.2%})<extra></extra>",


        ))


        fig_sit.update_layout(


            template="plotly_dark",


            title=dict(text="<b>Situação da Comissão</b>", font_size=15, x=0.5, font=dict(color="#9b8a5c")),


            height=380, showlegend=False,


            paper_bgcolor="rgba(0,0,0,0)",


            plot_bgcolor="rgba(0,0,0,0)",


            margin=dict(t=50, b=70, l=50, r=50),


        )


        st.plotly_chart(fig_sit, use_container_width=True)

    with c2:
        if "hide_enc_chart" not in st.session_state:
            st.session_state.hide_enc_chart = True
            
        btn_label = "👁️ Mostrar Encerradas" if st.session_state.hide_enc_chart else "🙈 Ocultar Encerradas"
        if st.button(btn_label, key="btn_toggle_enc_chart", use_container_width=True):
            st.session_state.hide_enc_chart = not st.session_state.hide_enc_chart
            st.rerun()
            
        excluir_encerradas = st.session_state.hide_enc_chart
        st.markdown("<p style='font-size: 0.78rem; color: #a0a0a0; margin-top: -8px; margin-bottom: 12px; font-style: italic;'>ℹ️ Oculta avaliações encerradas para ampliar e detalhar a escala dos status pendentes.</p>", unsafe_allow_html=True)
        
        cross = df.groupby(["Status Avaliação","Situação Comissão"]).size().reset_index(name="Qtd")
        
        if excluir_encerradas:
            cross = cross[cross["Status Avaliação"] != "Encerrada"]
            
        current_labels = [l for l in ordered_labels if l != "Encerrada"] if excluir_encerradas else ordered_labels
        cross["Status Avaliação"] = pd.Categorical(cross["Status Avaliação"],
                                                    categories=current_labels, ordered=True)
        cross = cross.sort_values("Status Avaliação")
        
        fig_bar = px.bar(cross, x="Status Avaliação", y="Qtd", color="Situação Comissão",
                         color_discrete_map=SIT_COLORS, barmode="group", text="Qtd",
                         template="plotly_dark",
                         title="<b>Status × Situação Comissão</b>")
        
        fig_bar.update_traces(textposition="outside", textfont_size=11)
        
        fig_bar.update_layout(height=380, title_font_size=15, title_x=0.5,
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                               xaxis_title="", yaxis_title="Qtd",
                               showlegend=False,
                               title_font=dict(color="#9b8a5c"))
        
        fig_bar.update_xaxes(showgrid=False)
        fig_bar.update_yaxes(showgrid=True, gridcolor="#2a2a2a")
        st.plotly_chart(fig_bar, use_container_width=True)





    # ── LINHA 3: Barras por RPM ────────────────────────────────────────────────


    st.markdown("---")


    col_s1, col_s2 = st.columns([1, 1])
    with col_s1:
        sort_option = st.selectbox(
            "Ordenação das Unidades (RPM):",
            [
                "Crescente por Unidade",
                "Decrescente por Unidade",
                "Crescente por Quantidade",
                "Decrescente por Quantidade"
            ],
            key="dist_chart_sort_opt"
        )
    with col_s2:
        st.write("") # spacing

    # Initialize legend button states if not in session state
    if "dist_legend_enc" not in st.session_state:
        st.session_state.dist_legend_enc = True
    if "dist_legend_abe" not in st.session_state:
        st.session_state.dist_legend_abe = True
    if "dist_legend_par" not in st.session_state:
        st.session_state.dist_legend_par = True
    if "dist_legend_hom" not in st.session_state:
        st.session_state.dist_legend_hom = True

    # Legenda Interativa - Checkboxes para selecionar os status
    st.markdown("<p style='font-size: 0.95rem; font-weight: bold; margin-bottom: 6px; color: #9b8a5c;'>Legenda Interativa — Selecione os Status para Exibição e Ordenação:</p>", unsafe_allow_html=True)
    l1, l2, l3, l4 = st.columns(4)
    with l1:
        enc_label = "🟢 Encerrada" if st.session_state.dist_legend_enc else "⚪ Encerrada"
        if st.button(enc_label, key="btn_dist_enc", use_container_width=True):
            st.session_state.dist_legend_enc = not st.session_state.dist_legend_enc
            st.rerun()
        show_enc = st.session_state.dist_legend_enc
    with l2:
        abe_label = "🔴 Aberta" if st.session_state.dist_legend_abe else "⚪ Aberta"
        if st.button(abe_label, key="btn_dist_abe", use_container_width=True):
            st.session_state.dist_legend_abe = not st.session_state.dist_legend_abe
            st.rerun()
        show_abe = st.session_state.dist_legend_abe
    with l3:
        par_label = "🟠 Parcialmente Encerrada" if st.session_state.dist_legend_par else "⚪ Parcialmente Encerrada"
        if st.button(par_label, key="btn_dist_par", use_container_width=True):
            st.session_state.dist_legend_par = not st.session_state.dist_legend_par
            st.rerun()
        show_par = st.session_state.dist_legend_par
    with l4:
        hom_label = "🟡 Homologação" if st.session_state.dist_legend_hom else "⚪ Homologação"
        if st.button(hom_label, key="btn_dist_hom", use_container_width=True):
            st.session_state.dist_legend_hom = not st.session_state.dist_legend_hom
            st.rerun()
        show_hom = st.session_state.dist_legend_hom

    # Mapear status selecionados
    active_statuses = []
    if show_enc: active_statuses.append("Encerrada")
    if show_abe: active_statuses.append("Aberta")
    if show_par: active_statuses.append("Parcialmente Encerrada")
    if show_hom: active_statuses.append("Homologação")

    if not active_statuses:
        active_statuses = ["Encerrada", "Aberta", "Parcialmente Encerrada", "Homologação"]

    # Filtrar o DataFrame pelos status selecionados
    df_filtered = df[df["Status Avaliação"].isin(active_statuses)]

    grp_col = "Unidade RPM (Avaliado)" if main_active_role != "P1" else "Unidade Principal (Avaliado)"
    title_text = "<b>Distribuição por Unidade RPM e Status</b>" if main_active_role != "P1" else f"<b>Distribuição por Subunidade ({active_rpm}) e Status</b>"
    
    # Obter lista de unidades presentes
    all_units = df_filtered[grp_col].dropna().unique()
    if len(all_units) == 0:
        all_units = df[grp_col].dropna().unique()

    # Ordenar as unidades com base na opção selecionada e nos status ativos
    if sort_option == "Crescente por Unidade":
        all_units_sorted = sorted(all_units, key=rpm_sort_key)
    elif sort_option == "Decrescente por Unidade":
        all_units_sorted = sorted(all_units, key=rpm_sort_key, reverse=True)
    elif sort_option == "Crescente por Quantidade":
        unit_totals = df_filtered.groupby(grp_col).size().reset_index(name="Total")
        unit_totals_sorted = unit_totals.sort_values("Total", ascending=True)
        all_units_sorted = list(unit_totals_sorted[grp_col])
        for u in all_units:
            if u not in all_units_sorted:
                all_units_sorted.append(u)
    else: # Decrescente por Quantidade
        unit_totals = df_filtered.groupby(grp_col).size().reset_index(name="Total")
        unit_totals_sorted = unit_totals.sort_values("Total", ascending=False)
        all_units_sorted = list(unit_totals_sorted[grp_col])
        for u in all_units:
            if u not in all_units_sorted:
                all_units_sorted.append(u)

    rpm_cross = df_filtered.groupby([grp_col,"Status Avaliação"]).size().reset_index(name="Qtd")
    
    fig_rpm = px.bar(
        rpm_cross, x=grp_col, y="Qtd",
        color="Status Avaliação", color_discrete_map=STATUS_COLORS,
        barmode="stack", text_auto=True,
        template="plotly_dark",
        title=title_text,
        category_orders={
            grp_col: all_units_sorted,
            "Status Avaliação": STACK_ORDER,
        },
    )
    
    fig_rpm.update_traces(textposition="auto")
    
    num_cols = len(all_units_sorted)
    if num_cols <= 6:
        fig_rpm.update_traces(width=0.05 * num_cols)
    
    fig_rpm.update_layout(
        uirevision="constant_value",
        height=480, title_font_size=15, title_x=0.5,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="", yaxis_title="Avaliações",
        showlegend=False,
        title_font=dict(color="#9b8a5c")
    )
    
    fig_rpm.update_xaxes(tickangle=45, showgrid=False)
    fig_rpm.update_yaxes(showgrid=True, gridcolor="#2a2a2a")
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


                          template="plotly_dark",


                          title="<b>Certificação Homologador</b>", text="Qtd")


        fig_cert.update_traces(textposition="outside", textfont_size=12)


        fig_cert.update_layout(height=360, title_x=0.5, showlegend=False,


                                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",


                                xaxis_title="", yaxis_title="Qtd",


                                title_font=dict(color="#9b8a5c"))


        fig_cert.update_xaxes(showgrid=False)


        fig_cert.update_yaxes(showgrid=True, gridcolor="#2a2a2a")


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
            
            # Filtro de datas
            min_date = df_time["Data"].min().date() if not df_time.empty else datetime.date.today()
            max_date = df_time["Data"].max().date() if not df_time.empty else datetime.date.today()
            
            st.markdown("<br>", unsafe_allow_html=True)
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                dt_ini = st.date_input("📅 Data Início da Pesquisa", value=min_date, min_value=min_date, max_value=max_date, key="dt_ini_timechart")
            with col_d2:
                dt_fim = st.date_input("📅 Data Fim da Pesquisa", value=max_date, min_value=min_date, max_value=max_date, key="dt_fim_timechart")
            # Seletores de Filtros de Função (Cards)
            for k in ["sh_av1", "sh_av2", "sh_hom"]:
                if k not in st.session_state: st.session_state[k] = True
            
            st.markdown("<br>", unsafe_allow_html=True)
            cb1, cb2, cb3 = st.columns(3)
            with cb1:
                if st.button("🔵 AV1" if st.session_state.sh_av1 else "⚪ AV1", use_container_width=True): st.session_state.sh_av1 = not st.session_state.sh_av1
            with cb2:
                if st.button("🟠 AV2" if st.session_state.sh_av2 else "⚪ AV2", use_container_width=True): st.session_state.sh_av2 = not st.session_state.sh_av2
            with cb3:
                if st.button("🟢 HOM" if st.session_state.sh_hom else "⚪ HOM", use_container_width=True): st.session_state.sh_hom = not st.session_state.sh_hom
                
            allowed = []
            if st.session_state.sh_av1: allowed.append("AV1 — Avaliador 1")
            if st.session_state.sh_av2: allowed.append("AV2 — Avaliador 2")
            if st.session_state.sh_hom: allowed.append("HOM — Homologador")
            
            mask = (df_time["Data"].dt.date >= dt_ini) & (df_time["Data"].dt.date <= dt_fim) & (df_time["Função"].isin(allowed))
            df_time_filtered = df_time[mask]

            fig_time = px.line(
                df_time_filtered, x="Data", y="Qtd", color="Função",


                title="<b>Avaliações por Data e Função</b>",


                template="plotly_dark",


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
                height=380, title_x=0.5, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                xaxis_title="", yaxis_title="Avaliações encerradas",
                showlegend=False,
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.45,
                    xanchor="center",
                    x=0.5,
                    font=dict(color="#e5dccb"),
                    bgcolor="rgba(0,0,0,0)"
                ),
                hovermode="x unified",
                title_font=dict(color="#9b8a5c")
            )


            fig_time.update_xaxes(showgrid=False, tickformat="%d/%m/%Y",


                                   rangeslider_visible=True)


            fig_time.update_yaxes(showgrid=True, gridcolor="#2a2a2a")


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


        "Sit. Funcional",


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
                            .map(color_sit,   subset=["Situação Comissão"]), height=540,
            show_download=True, download_name=f"avaliacoes_filtradas_{now_br().strftime('%Y%m%d_%H%M')}")





# ══════════════════════════════════════════════════════════════════════════════


# TAB 3 — AVALIAÇÕES PENDENTES


# ══════════════════════════════════════════════════════════════════════════════


if active_page == "Avaliações Pendentes":


    st.markdown("### ⏳ Avaliações Pendentes")


    STATUS_PEND = {
        "Homologação", "Parcialmente Encerrada", "Aberta",
        "EM PRAZO DE RECURSO", "RECONSIDERAÇÃO COMISSÃO", "AUTORIDADE RECURSAL"
    }


    df_pend = df[df["Status Avaliação"].isin(STATUS_PEND) & df["Sit. Funcional"].isin(SITUACOES_ALVO)].copy()





    c1, c2 = st.columns(2)


    with c1:


        tipo_pend = st.multiselect("Status:", ["Aberta", "Parcialmente Encerrada", "Homologação", "EM RECURSO"],


                                   default=["Aberta", "Parcialmente Encerrada", "Homologação", "EM RECURSO"],


                                   key="tp")


    with c2:


        sc_pend = st.multiselect("Situação:", ["Comissão Atual","Nota Provisória"],


                                  default=["Comissão Atual","Nota Provisória"], key="sp")




    # Mapeamento do filtro para os valores reais do dataframe
    filter_statuses = []
    for tp in tipo_pend:
        if tp == "EM RECURSO":
            filter_statuses.extend(["RECONSIDERAÇÃO COMISSÃO", "AUTORIDADE RECURSAL", "EM PRAZO DE RECURSO"])
        else:
            filter_statuses.append(tp)

    df_pv = df_pend[df_pend["Status Avaliação"].isin(filter_statuses) &


                    df_pend["Situação Comissão"].isin(sc_pend)] if tipo_pend and sc_pend else df_pend





    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f'<div class="kpi-card kpi-aberta">'
                    '<div class="label">🔴 Abertas</div>'
                    f'<div class="value">{fmt_num((df_pv["Status Avaliação"]=="Aberta").sum())}</div>'
                    '<div class="sub">AV1 pendente</div>'
                    '</div>', unsafe_allow_html=True)
    with k2:
        st.markdown(f'<div class="kpi-card kpi-parc">'
                    '<div class="label">🟠 Parc. Encerradas</div>'
                    f'<div class="value">{fmt_num((df_pv["Status Avaliação"]=="Parcialmente Encerrada").sum())}</div>'
                    '<div class="sub">AV2 pendente</div>'
                    '</div>', unsafe_allow_html=True)
    with k3:
        st.markdown(f'<div class="kpi-card kpi-hom">'
                    '<div class="label">🟡 Homologação</div>'
                    f'<div class="value">{fmt_num((df_pv["Status Avaliação"]=="Homologação").sum())}</div>'
                    '<div class="sub">Homologação pendente</div>'
                    '</div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="kpi-card kpi-recurso">'
                    '<div class="label">🟣 Em Recurso</div>'
                    f'<div class="value">{fmt_num(df_pv["Status Avaliação"].isin(["RECONSIDERAÇÃO COMISSÃO", "AUTORIDADE RECURSAL", "EM PRAZO DE RECURSO"]).sum())}</div>'
                    '<div class="sub">Recurso pendente</div>'
                    '</div>', unsafe_allow_html=True)
    with k5:
        st.markdown(f'<div class="kpi-card kpi-total">'
                    '<div class="label">📊 Total</div>'
                    f'<div class="value">{fmt_num(len(df_pv))}</div>'
                    '<div class="sub">Registros filtrados</div>'
                    '</div>', unsafe_allow_html=True)





    cols_pend = [


        # ── Avaliado ──────────────────────────────────────────────────────────


        "nrPM (Avaliado)", "Posto/Grad. (Avaliado)", "Nome (Avaliado)",


        "Unidade RPM (Avaliado)", "Unidade Principal (Avaliado)",


        "Sit. Funcional",


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
            height=520, show_download=True, download_name=f"avaliacoes_pendentes_{now_br().strftime('%Y%m%d_%H%M')}", download_label="Baixar pendentes")








# ══════════════════════════════════════════════════════════════════════════════


# TAB 4 — AVALIADORES PENDENTES


# ══════════════════════════════════════════════════════════════════════════════


if active_page == "Avaliadores Pendentes":


    st.markdown("### 👥 Avaliadores Pendentes")





    df_alvo = df_full[df_full["Sit. Funcional"].isin(SITUACOES_ALVO)]

    if rpm_filter:
        df_ab = df_alvo[(df_alvo["Status Avaliação"] == "Aberta") & (df_alvo["RPM (Av1)"].isin(rpm_filter))]
        df_pe = df_alvo[(df_alvo["Status Avaliação"].isin(["Aberta", "Parcialmente Encerrada"])) & (df_alvo["RPM (Av2)"].isin(rpm_filter))]
        df_hom = df_alvo[(df_alvo["Status Avaliação"] == "Homologação") & (df_alvo["RPM (Hom)"].isin(rpm_filter))]
    else:
        df_alvo_filtered = df[df["Sit. Funcional"].isin(SITUACOES_ALVO)]
        df_ab = df_alvo_filtered[df_alvo_filtered["Status Avaliação"] == "Aberta"]
        df_pe = df_alvo_filtered[df_alvo_filtered["Status Avaliação"].isin(["Aberta", "Parcialmente Encerrada"])]
        df_hom = df_alvo_filtered[df_alvo_filtered["Status Avaliação"] == "Homologação"]



    # Pre-calcular quantitativo de avaliadores/homologadores únicos pendentes (função)
    cnt_av1 = df_ab["nrPM (Av1)"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique()
    cnt_av2 = df_pe["nrPM (Av2)"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique()
    cnt_hom = df_hom["nrPM (Hom)"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique()

    # Renderizar cards de quantitativos funcionais
    col_av1, col_av2, col_av3 = st.columns(3)
    with col_av1:
        st.markdown(f'<div class="kpi-card kpi-total">'
                    '<div class="label">👤 Avaliador 1 (AV1)</div>'
                    f'<div class="value">{fmt_num(cnt_av1)}</div>'
                    '<div class="sub">Avaliadores com pendência de AV1</div>'
                    '</div>', unsafe_allow_html=True)
    with col_av2:
        st.markdown(f'<div class="kpi-card kpi-parc">'
                    '<div class="label">👥 Avaliador 2 (AV2)</div>'
                    f'<div class="value">{fmt_num(cnt_av2)}</div>'
                    '<div class="sub">Avaliadores com pendência de AV2</div>'
                    '</div>', unsafe_allow_html=True)
    with col_av3:
        st.markdown(f'<div class="kpi-card kpi-hom">'
                    '<div class="label">⚖️ Homologador (HOM)</div>'
                    f'<div class="value">{fmt_num(cnt_hom)}</div>'
                    '<div class="sub">Homologadores com pendência de HOM</div>'
                    '</div>', unsafe_allow_html=True)

    st.markdown("<div style='margin-bottom: 20px;'></div>", unsafe_allow_html=True)





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


    k1,k2,k3 = st.columns([2,2,1])
    k1.metric("Avaliadores pendentes (AV1)", len(tb1))
    k2.metric("Total avaliações Em Aberto", tb1["Total AV1"].sum() if not tb1.empty else 0)
    with k3:
        st.markdown("<br>", unsafe_allow_html=True)
        if not tb1.empty:
            st.download_button("⬇️ AV1 (CSV)", tb1.to_csv(index=False,sep=";",encoding="utf-8-sig").encode("utf-8-sig"), f"av1_{now_br().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv", type="primary", use_container_width=True)


    if not tb1.empty:


        tb1_disp = tb1.reset_index(drop=True)
        tb1_disp.index = range(1, len(tb1_disp) + 1)
        tb1_disp = clean_none_values(tb1_disp)


        


        import inspect


        sig = inspect.signature(st.dataframe)


        has_select = "on_select" in sig.parameters


        


        selected_pm = None


        selected_nome = None


        


        if has_select:


            st.write("💡 *Dica: Clique em uma linha da tabela abaixo para abrir as avaliações deste avaliador.*")


            event1 = st.dataframe(


                tb1_disp,


                use_container_width=True,


                height=260,


                on_select="rerun",


                selection_mode="single-row",


                key="select_tb1"


            )


            rows1 = event1.get("selection", {}).get("rows", [])


            if rows1:


                idx = rows1[0]


                selected_pm = tb1.iloc[idx]["Nº PM"]


                selected_nome = tb1.iloc[idx]["Nome"]


        else:


            st.dataframe(tb1_disp, use_container_width=True, height=260)


            sel_nome = st.selectbox("🔎 Selecione um Avaliador 1 para ver as avaliações:", ["-- Selecione --"] + list(tb1["Nome"].unique()), key="sel_tb1")


            if sel_nome != "-- Selecione --":


                row_sel = tb1[tb1["Nome"] == sel_nome].iloc[0]


                selected_pm = row_sel["Nº PM"]


                selected_nome = row_sel["Nome"]


                


        if selected_pm:


            df_det1 = df_ab[df_ab["nrPM (Av1)"] == selected_pm].copy()


            st.markdown(f"#### 📋 Avaliações pendentes de AV1: **{selected_nome}** ({selected_pm})")


            cols_det1 = ["nrPM (Avaliado)", "Posto/Grad. (Avaliado)", "Nome (Avaliado)",


                         "Unidade RPM (Avaliado)", "Unidade Principal (Avaliado)",


                         "Status Avaliação", "Situação Comissão", "Data AV1"]


            cols_ok1 = [c for c in cols_det1 if c in df_det1.columns]


            safe_df(df_det1[cols_ok1].reset_index(drop=True).style.map(color_status, subset=["Status Avaliação"]), height=180)


            


            # Botões de download lado a lado para esta lista filtrada específica


            dl1, dl2 = st.columns(2)


            with dl1:


                st.download_button(


                    "⬇️ Baixar esta lista (Excel .xlsx)",


                    df_to_xlsx(df_det1[cols_ok1]),


                    f"pendencias_AV1_{selected_pm}_{now_br().strftime('%Y%m%d_%H%M')}.xlsx",


                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",


                    key="dl_av1_xlsx"


                )


            with dl2:


                st.download_button(


                    "⬇️ Baixar esta lista (CSV)",


                    df_det1[cols_ok1].to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig"),


                    f"pendencias_AV1_{selected_pm}_{now_br().strftime('%Y%m%d_%H%M')}.csv",


                    mime="text/csv",


                    key="dl_av1_csv"


                )


            



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


    k1,k2,k3 = st.columns([2,2,1])
    k1.metric("Avaliadores pendentes (AV2)", len(tb2))
    k2.metric("Total pendências AV2", tb2["Total AV2"].sum() if not tb2.empty else 0)
    with k3:
        st.markdown("<br>", unsafe_allow_html=True)
        if not tb2.empty:
            st.download_button("⬇️ AV2 (CSV)", tb2.to_csv(index=False,sep=";",encoding="utf-8-sig").encode("utf-8-sig"), f"av2_{now_br().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv", type="primary", use_container_width=True)


    if not tb2.empty:


        tb2_disp = tb2.reset_index(drop=True)
        tb2_disp.index = range(1, len(tb2_disp) + 1)
        tb2_disp = clean_none_values(tb2_disp)


        


        import inspect


        sig = inspect.signature(st.dataframe)


        has_select = "on_select" in sig.parameters


        


        selected_pm2 = None


        selected_nome2 = None


        


        if has_select:


            st.write("💡 *Dica: Clique em uma linha da tabela abaixo para abrir as avaliações deste avaliador.*")


            event2 = st.dataframe(


                tb2_disp,


                use_container_width=True,


                height=260,


                on_select="rerun",


                selection_mode="single-row",


                key="select_tb2"


            )


            rows2 = event2.get("selection", {}).get("rows", [])


            if rows2:


                idx = rows2[0]


                selected_pm2 = tb2.iloc[idx]["Nº PM"]


                selected_nome2 = tb2.iloc[idx]["Nome"]


        else:


            st.dataframe(tb2_disp, use_container_width=True, height=260)


            sel_nome2 = st.selectbox("🔎 Selecione um Avaliador 2 para ver as avaliações:", ["-- Selecione --"] + list(tb2["Nome"].unique()), key="sel_tb2")


            if sel_nome2 != "-- Selecione --":


                row_sel2 = tb2[tb2["Nome"] == sel_nome2].iloc[0]


                selected_pm2 = row_sel2["Nº PM"]


                selected_nome2 = row_sel2["Nome"]


                


        if selected_pm2:


            df_det2 = df_pe[df_pe["nrPM (Av2)"] == selected_pm2].copy()


            st.markdown(f"#### 📋 Avaliações pendentes de AV2: **{selected_nome2}** ({selected_pm2})")


            cols_det2 = ["nrPM (Avaliado)", "Posto/Grad. (Avaliado)", "Nome (Avaliado)",


                         "Unidade RPM (Avaliado)", "Unidade Principal (Avaliado)",


                         "Status Avaliação", "Situação Comissão", "Data AV1", "Data AV2"]


            cols_ok2 = [c for c in cols_det2 if c in df_det2.columns]


            safe_df(df_det2[cols_ok2].reset_index(drop=True).style.map(color_status, subset=["Status Avaliação"]), height=180)


            


            # Botões de download lado a lado para esta lista filtrada específica


            dl1, dl2 = st.columns(2)


            with dl1:


                st.download_button(


                    "⬇️ Baixar esta lista (Excel .xlsx)",


                    df_to_xlsx(df_det2[cols_ok2]),


                    f"pendencias_AV2_{selected_pm2}_{now_br().strftime('%Y%m%d_%H%M')}.xlsx",


                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",


                    key="dl_av2_xlsx"


                )


            with dl2:


                st.download_button(


                    "⬇️ Baixar esta lista (CSV)",


                    df_det2[cols_ok2].to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig"),


                    f"pendencias_AV2_{selected_pm2}_{now_br().strftime('%Y%m%d_%H%M')}.csv",


                    mime="text/csv",


                    key="dl_av2_csv"


                )


            






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





    k1, k2, k3, k4 = st.columns([1,1,1,1])
    k1.metric("Homologadores com pendência", len(tb3))
    k2.metric("Total aguardando HOM", len(df_hom))
    k3.metric("CA / NP pendentes",
              f"{(df_hom['Situação Comissão']=='Comissão Atual').sum()} / "
              f"{(df_hom['Situação Comissão']=='Nota Provisória').sum()}")
    with k4:
        st.markdown("<br>", unsafe_allow_html=True)
        if not tb3.empty:
            st.download_button("⬇️ Homologadores (CSV)", tb3.to_csv(index=False,sep=";",encoding="utf-8-sig").encode("utf-8-sig"), f"homologadores_{now_br().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv", type="primary", use_container_width=True)





    if not tb3.empty:


        tb3_disp = tb3.reset_index(drop=True)
        tb3_disp.index = range(1, len(tb3_disp) + 1)
        tb3_disp = clean_none_values(tb3_disp)


        


        import inspect


        sig = inspect.signature(st.dataframe)


        has_select = "on_select" in sig.parameters


        


        selected_pm3 = None


        selected_nome3 = None


        


        if has_select:


            st.write("💡 *Dica: Clique em uma linha da tabela abaixo para abrir as avaliações deste homologador.*")


            event3 = st.dataframe(


                tb3_disp,


                use_container_width=True,


                height=280,


                on_select="rerun",


                selection_mode="single-row",


                key="select_tb3"


            )


            rows3 = event3.get("selection", {}).get("rows", [])


            if rows3:


                idx = rows3[0]


                selected_pm3 = tb3.iloc[idx]["Nº PM"]


                selected_nome3 = tb3.iloc[idx]["Nome"]


        else:


            st.dataframe(tb3_disp, use_container_width=True, height=280)


            sel_nome3 = st.selectbox("🔎 Selecione um Homologador para ver as avaliações:", ["-- Selecione --"] + list(tb3["Nome"].unique()), key="sel_tb3")


            if sel_nome3 != "-- Selecione --":


                row_sel3 = tb3[tb3["Nome"] == sel_nome3].iloc[0]


                selected_pm3 = row_sel3["Nº PM"]


                selected_nome3 = row_sel3["Nome"]


                


        if selected_pm3:


            df_det3 = df_hom[df_hom["nrPM (Hom)"] == selected_pm3].copy()


            st.markdown(f"#### 📋 Avaliações pendentes do Homologador: **{selected_nome3}** ({selected_pm3})")


            cols_det3 = ["nrPM (Avaliado)", "Posto/Grad. (Avaliado)", "Nome (Avaliado)",


                         "Unidade RPM (Avaliado)", "Unidade Principal (Avaliado)",


                         "Status Avaliação", "Situação Comissão", "Data AV1", "Data AV2", "Data HOM"]


            cols_ok3 = [c for c in cols_det3 if c in df_det3.columns]


            safe_df(df_det3[cols_ok3].reset_index(drop=True).style.map(color_status, subset=["Status Avaliação"]), height=180)


            


            # Botões de download lado a lado para esta lista filtrada específica


            dl1, dl2 = st.columns(2)


            with dl1:


                st.download_button(


                    "⬇️ Baixar esta lista (Excel .xlsx)",


                    df_to_xlsx(df_det3[cols_ok3]),


                    f"pendencias_HOM_{selected_pm3}_{now_br().strftime('%Y%m%d_%H%M')}.xlsx",


                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",


                    key="dl_hom_xlsx"


                )


            with dl2:


                st.download_button(


                    "⬇️ Baixar esta lista (CSV)",


                    df_det3[cols_ok3].to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig"),


                    f"pendencias_HOM_{selected_pm3}_{now_br().strftime('%Y%m%d_%H%M')}.csv",


                    mime="text/csv",


                    key="dl_hom_csv"


                )


            


        st.download_button(


            "⬇️ Homologadores pendentes (CSV)",


            tb3.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig"),


            f"hom_{now_br().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv")


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


                f"hom_detalhe_{now_br().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv")





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


    "Sit. Funcional",


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





    if df_global is not None:
        df_global = df_global[df_global["Sit. Funcional"].isin(SITUACOES_ALVO)]
    df_unit = df_unit[df_unit["Sit. Funcional"].isin(SITUACOES_ALVO)]

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


        pct   = f"{tot_n/total*100:.2f}%" if total > 0 else "0%"





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


        pct   = f"{tot_n/total*100:.2f}%" if total > 0 else "0%"





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


    df_pend = df_unit[df_unit["Status Avaliação"].isin(["Aberta", "Parcialmente Encerrada", "Homologação"]) & df_unit["Sit. Funcional"].isin(SITUACOES_ALVO)]


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
    zip_name = f"AADP_2026_{now_br().strftime('%Y%m%d_%H%M%S')}.zip"

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


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — GERAR RELATÓRIO EXCEL
# ══════════════════════════════════════════════════════════════════════════════
if active_page == "Gerar Relatório":
    st.markdown("### 📥 Gerar Relatório Excel")
    
    if active_role in ("P1", "SADM"):
        st.warning("⚠️ Você não possui permissão para acessar esta funcionalidade.")
    else:
        if "excel_modo_rel" not in st.session_state:
            st.session_state.excel_modo_rel = "Completo"
            
        st.markdown("<p style='font-size: 1.1rem; font-weight: bold; margin-bottom: 8px; color: #9b8a5c;'>Tipo de relatório:</p>", unsafe_allow_html=True)
        st.markdown("<div class='excel-scope-marker'></div>", unsafe_allow_html=True)
        
        col_ex1, col_ex2, col_ex3 = st.columns(3)
        with col_ex1:
            is_ex1 = (st.session_state.excel_modo_rel == "Completo")
            btn_ex1_type = "primary" if is_ex1 else "secondary"
            if st.button("🌐\nCompleto\n(Geral + RPMs)", key="btn_excel_scope_completo", use_container_width=True, type=btn_ex1_type):
                st.session_state.excel_modo_rel = "Completo"
                st.rerun()
        with col_ex2:
            is_ex2 = (st.session_state.excel_modo_rel == "Somente Geral")
            btn_ex2_type = "primary" if is_ex2 else "secondary"
            if st.button("📋\nGeral\n(Somente Geral)", key="btn_excel_scope_geral", use_container_width=True, type=btn_ex2_type):
                st.session_state.excel_modo_rel = "Somente Geral"
                st.rerun()
        with col_ex3:
            is_ex3 = (st.session_state.excel_modo_rel == "Unidades RPM específicas")
            btn_ex3_type = "primary" if is_ex3 else "secondary"
            if st.button("🎯\nEspecíficas\n(Filtrar Unidades)", key="btn_excel_scope_especifica", use_container_width=True, type=btn_ex3_type):
                st.session_state.excel_modo_rel = "Unidades RPM específicas"
                st.rerun()
                
        modo_rel = st.session_state.excel_modo_rel
    
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
                        log_action(st.session_state.user_pm, "EXPORTAR_EXCEL", f"Modo: {modo_code}, Unidades: {units_sel}")
                        st.download_button(
                            label=f"⬇️ Baixar {zip_name}",
                            data=zip_bytes,
                            file_name=zip_name,
                            mime="application/zip",
                            use_container_width=True,
                        )
                    except Exception as ex:
                        st.error(f"❌ Erro na geração: {ex}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — HOMOLOGAÇÃO RELATÓRIO WORD (.DOCX)
# ══════════════════════════════════════════════════════════════════════════════
if active_page == "Relatório Word":
    st.markdown("### 📄 Relatório Word (.docx)")
    df_word = df_full
                
    if df_word is not None:
        st.markdown("---")
        
        if active_role in ("SADM",):
            st.warning("⚠️ Você não possui permissão para acessar esta funcionalidade.")
        else:
            st.markdown("<h4 style='font-size: 1.35rem; font-weight: bold; margin-bottom: 12px; color: #9b8a5c;'>Configurações do Relatório</h4>", unsafe_allow_html=True)
            
            if "rel_scope" not in st.session_state:
                st.session_state.rel_scope = "Geral RPM" if active_role in ("ADMINISTRADOR", "GESTOR") else f"Geral Subordinadas da {active_rpm}"
                
            st.markdown("<p style='font-size: 1.1rem; font-weight: bold; margin-bottom: 8px; color: #9b8a5c;'>Escopo do Relatório:</p>", unsafe_allow_html=True)
            st.markdown("<div class='report-scope-marker'></div>", unsafe_allow_html=True)
            
            selected_rpms = []
            
            if active_role in ("ADMINISTRADOR", "GESTOR"):
                col_sc1, col_sc2, col_sc3 = st.columns(3)
                with col_sc1:
                    is_sc1 = (st.session_state.rel_scope == "Geral RPM")
                    btn_sc1_type = "primary" if is_sc1 else "secondary"
                    if st.button("🏢\nGeral RPM\n(UDI/UDG Principais)", key="btn_scope_geral_rpm", use_container_width=True, type=btn_sc1_type):
                        st.session_state.rel_scope = "Geral RPM"
                        st.rerun()
                with col_sc2:
                    is_sc2 = (st.session_state.rel_scope == "Geral Subordinadas")
                    btn_sc2_type = "primary" if is_sc2 else "secondary"
                    if st.button("🌐\nGeral Subordinadas\n(UDI/UDG + Subordinadas)", key="btn_scope_geral_sub", use_container_width=True, type=btn_sc2_type):
                        st.session_state.rel_scope = "Geral Subordinadas"
                        st.rerun()
                with col_sc3:
                    is_sc3 = (st.session_state.rel_scope == "Por RPM específica")
                    btn_sc3_type = "primary" if is_sc3 else "secondary"
                    if st.button("🎯\nPor RPM específica\n(Filtrar por Unidades)", key="btn_scope_especifica", use_container_width=True, type=btn_sc3_type):
                        st.session_state.rel_scope = "Por RPM específica"
                        st.rerun()
                        
                rel_scope = st.session_state.rel_scope
                if "específica" in rel_scope:
                    unique_rpms = sorted(df_word["Unidade RPM (Avaliado)"].dropna().unique().tolist(), key=rpm_sort_key)
                    selected_rpms = st.multiselect("Selecione as Unidades UDI/UDG para o relatório:", unique_rpms)
                    
            elif active_role == "P1":
                # For P1, we force general subordinates and hide the UI
                st.session_state.rel_scope = f"Geral Subordinadas da {active_rpm}"
                rel_scope = st.session_state.rel_scope

            # 🚀 Botão de Geração em Destaque no Início da Página
            if "específica" in rel_scope and not selected_rpms:
                st.warning("⚠️ Selecione ao menos uma unidade para gerar o relatório.")
            else:
                if st.button("🚀 Gerar e Baixar Relatório Word", key="btn_word_gen", type="primary", use_container_width=True):
                    with st.spinner("⏳ Gerando relatório executivo Word com gráficos... (Isso pode levar alguns instantes)"):
                        try:
                            from gerar_relatorio_word import generate_word_report
                            mode_code = "geral_rpm" if "Geral RPM" in rel_scope else "geral_subordinadas" if "Geral Subordinadas" in rel_scope else "especifica"
                            doc_bytes = generate_word_report(df_word, mode_code, selected_rpms, user_role=active_role)
                            st.success("✅ Relatório Word gerado com sucesso!")
                            log_action(active_pm, "EXPORTAR_WORD", f"Modo: {mode_code}, RPMs: {selected_rpms}")
                            doc_name = f"Relatorio_Executivo_AADP2026_{now_br().strftime('%Y%m%d_%H%M%S')}.docx"
                            st.download_button(
                                label=f"⬇️ Baixar {doc_name}",
                                data=doc_bytes,
                                file_name=doc_name,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                        except Exception as ex:
                            st.error(f"❌ Erro ao gerar o relatório: {ex}")
                            
            st.markdown("---")
            
            # ── PRÉ-VISUALIZAÇÃO DE GRÁFICOS DO WORD ──────────────────────────────────
            st.markdown("<h4 style='font-size: 1.25rem; font-weight: bold; margin-top: 15px; margin-bottom: 12px; color: #9b8a5c;'>📊 Pré-visualização dos Gráficos do Relatório</h4>", unsafe_allow_html=True)
            
            show_preview = True
            if "específica" in rel_scope and not selected_rpms:
                st.info("💡 Selecione ao menos uma UDI/UDG acima para visualizar a pré-visualização dos gráficos.")
                show_preview = False
                
            if show_preview:
                df_word_clean = df_word.copy()
                df_word_clean["Unidade Principal (Avaliado)"] = (
                    df_word_clean["Unidade Principal (Avaliado)"]
                    .astype(str)
                    .str.strip()
                    .replace({"nan": "", "-": ""})
                )
                mask_empty = df_word_clean["Unidade Principal (Avaliado)"] == ""
                df_word_clean.loc[mask_empty, "Unidade Principal (Avaliado)"] = df_word_clean.loc[mask_empty, "Unidade RPM (Avaliado)"]
                
                # Filtrar pelo escopo do relatório
                if "específica" in rel_scope:
                    df_escopo = df_word_clean[df_word_clean["Unidade RPM (Avaliado)"].isin(selected_rpms)].copy()
                    units_to_render = selected_rpms
                else:
                    df_escopo = df_word_clean
                    units_to_render = sorted(df_word_clean["Unidade RPM (Avaliado)"].dropna().unique().tolist(), key=rpm_sort_key)
                
                # Tipo de pré-visualização
                if active_role == "P1":
                    st.info("💡 **Nota:** O botão de download acima sempre gerará o relatório completo com **TODAS** as unidades subordinadas da sua RPM.")
                    sub_units_available = sorted(df_escopo["Unidade Principal (Avaliado)"].dropna().unique().tolist())
                    selected_subs = st.multiselect("Filtrar pré-visualização por Subunidade(s) específica(s) (Opcional):", sub_units_available, key="p1_prev_sub_filter")
                    prev_type = "P1_Custom"
                else:
                    prev_type = st.selectbox(
                        "Selecione o nível de detalhamento dos gráficos para visualização:",
                        ["🌍 Geral (Toda a PMMG + Gráficos de cada UDI/UDG)", "🏢 Por UDI/UDG específica (RPM + Gráficos das Subunidades)"],
                        key="word_prev_type"
                    )
                
                from gerar_relatorio_word import create_status_pie, create_comissao_bar, create_pending_units_bar
                import tempfile
                
                if prev_type == "🌍 Geral (Toda a PMMG + Gráficos de cada UDI/UDG)":
                    with st.spinner("⏳ Gerando prévia dos gráficos de toda a PMMG..."):
                        # 1. Visão Geral do Estado
                        st.markdown("<h5 style='font-size: 1.25rem; font-weight: bold; color: #9b8a5c; margin-top: 15px; margin-bottom: 10px;'>1. Visão Geral do Estado</h5>", unsafe_allow_html=True)
                        
                        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f_pie:
                            t_pie = f_pie.name
                        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f_bar:
                            t_bar = f_bar.name
                        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f_pend:
                            t_pend = f_pend.name
                            
                        try:
                            create_status_pie(df_escopo, t_pie)
                            create_comissao_bar(df_escopo, t_bar)
                            create_pending_units_bar(df_escopo, units_to_render, t_pend)
                            
                            c_sp1, c1, c2, c_sp2 = st.columns([1, 2, 2, 1])
                            with c1:
                                st.image(t_pie, caption="Status das Avaliações - Geral", use_container_width=True)
                            with c2:
                                st.image(t_bar, caption="Situação da Comissão - Geral", use_container_width=True)
                            
                            c_p_sp1, c_p, c_p_sp2 = st.columns([1, 2, 1])
                            with c_p:
                                st.image(t_pend, caption="Pendências Acumuladas por UDI/UDG", use_container_width=True)
                        finally:
                            for p in [t_pie, t_bar, t_pend]:
                                try:
                                    if os.path.exists(p): os.remove(p)
                                except Exception: pass
                                
                        # 2. Detalhamento de todas as UDI/UDG (sem subunidades)
                        st.markdown("<h5 style='font-size: 1.25rem; font-weight: bold; color: #9b8a5c; margin-top: 25px; margin-bottom: 10px;'>2. Detalhamento de todas as UDI/UDG</h5>", unsafe_allow_html=True)
                        for rpm in units_to_render:
                            df_rpm = df_escopo[df_escopo["Unidade RPM (Avaliado)"] == rpm]
                            if len(df_rpm) > 0:
                                st.markdown(
                                    f"""
                                    <div style='background-color: #3e3825; padding: 8px 12px; border-left: 5px solid #9b8a5c; border-radius: 4px; margin-top: 20px; margin-bottom: 10px;'>
                                        <span style='font-size: 1.15rem; font-weight: bold; color: #f5f5f5;'>🏢 Unidade Principal: {rpm}</span>
                                        <span style='font-size: 0.95rem; color: #cbbb8f; margin-left: 10px;'>({len(df_rpm):,} avaliações)</span>
                                    </div>
                                    """,
                                    unsafe_allow_html=True
                                )
                                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f_pie:
                                    t_pie = f_pie.name
                                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f_bar:
                                    t_bar = f_bar.name
                                try:
                                    create_status_pie(df_rpm, t_pie)
                                    create_comissao_bar(df_rpm, t_bar)
                                    c_sp1, c1, c2, c_sp2 = st.columns([1, 2, 2, 1])
                                    with c1:
                                        st.image(t_pie, use_container_width=True)
                                    with c2:
                                        st.image(t_bar, use_container_width=True)
                                finally:
                                    for p in [t_pie, t_bar]:
                                        try:
                                            if os.path.exists(p): os.remove(p)
                                        except Exception: pass
                                        
                elif prev_type == "🏢 Por UDI/UDG específica (RPM + Gráficos das Subunidades)":
                    # Por UDI/UDG específica (RPM + Gráficos das Subunidades)
                    rpms_available = sorted(df_escopo["Unidade RPM (Avaliado)"].dropna().unique().tolist(), key=rpm_sort_key)
                    if rpms_available:
                        selected_rpm = st.selectbox("Selecione a UDI/UDG (RPM) para detalhar:", rpms_available, key="prev_select_rpm")
                        
                        df_rpm = df_escopo[df_escopo["Unidade RPM (Avaliado)"] == selected_rpm]
                        
                        if not df_rpm.empty:
                            with st.spinner(f"⏳ Gerando gráficos de {selected_rpm} e suas subordinadas..."):
                                # 1. Gráficos da Unidade Principal
                                st.markdown(
                                    f"""
                                    <div style='background-color: #3e3825; padding: 10px 15px; border-left: 5px solid #9b8a5c; border-radius: 4px; margin-top: 20px; margin-bottom: 15px;'>
                                        <span style='font-size: 1.25rem; font-weight: bold; color: #f5f5f5;'>🏢 Unidade Principal: {selected_rpm}</span>
                                        <span style='font-size: 1.0rem; color: #cbbb8f; margin-left: 10px;'>({len(df_rpm):,} avaliações)</span>
                                    </div>
                                    """,
                                    unsafe_allow_html=True
                                )
                                
                                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f_pie:
                                    t_pie = f_pie.name
                                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f_bar:
                                    t_bar = f_bar.name
                                try:
                                    create_status_pie(df_rpm, t_pie)
                                    create_comissao_bar(df_rpm, t_bar)
                                    c_sp1, c1, c2, c_sp2 = st.columns([1, 2, 2, 1])
                                    with c1:
                                        st.image(t_pie, use_container_width=True)
                                    with c2:
                                        st.image(t_bar, use_container_width=True)
                                finally:
                                    for p in [t_pie, t_bar]:
                                        try:
                                            if os.path.exists(p): os.remove(p)
                                        except Exception: pass
                                        
                                # 2. Gráficos das Subunidades (Subordinadas)
                                st.markdown(f"<h5 style='font-size: 1.15rem; font-weight: bold; color: #9b8a5c; margin-top: 25px; margin-bottom: 10px;'>2.1 Unidades Subordinadas de {selected_rpm}</h5>", unsafe_allow_html=True)
                                unique_subs = sorted(df_rpm["Unidade Principal (Avaliado)"].dropna().unique().tolist())
                                for sub in unique_subs:
                                    df_sub = df_rpm[df_rpm["Unidade Principal (Avaliado)"] == sub]
                                    if len(df_sub) > 0:
                                        st.markdown(
                                            f"""
                                            <div style='background-color: #26231b; padding: 6px 10px; border-left: 3px solid #ff9f43; border-radius: 4px; margin-top: 15px; margin-bottom: 8px;'>
                                                <span style='font-size: 1.05rem; font-weight: bold; color: #f5f5f5;'>■ Subunidade: {sub}</span>
                                                <span style='font-size: 0.9rem; color: #ff9f43; margin-left: 10px;'>({len(df_sub):,} avaliações)</span>
                                            </div>
                                            """,
                                            unsafe_allow_html=True
                                        )
                                        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f_pie:
                                            t_pie = f_pie.name
                                        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f_bar:
                                            t_bar = f_bar.name
                                        try:
                                            create_status_pie(df_sub, t_pie)
                                            create_comissao_bar(df_sub, t_bar)
                                            c1, c2 = st.columns(2)
                                            with c1:
                                                st.image(t_pie, use_container_width=True)
                                            with c2:
                                                st.image(t_bar, use_container_width=True)
                                        finally:
                                            for p in [t_pie, t_bar]:
                                                try:
                                                    if os.path.exists(p): os.remove(p)
                                                except Exception: pass
                        else:
                            st.info("Sem dados disponíveis para a unidade selecionada.")
                    else:
                        st.info("Nenhuma UDI/UDG disponível no escopo atual.")
                
                elif prev_type == "P1_Custom":
                    df_rpm = df_escopo[df_escopo["Unidade RPM (Avaliado)"] == active_rpm]
                    if not df_rpm.empty:
                        with st.spinner(f"⏳ Gerando gráficos das subordinadas de {active_rpm}..."):
                            unique_subs = sorted(df_rpm["Unidade Principal (Avaliado)"].dropna().unique().tolist())
                            if selected_subs:
                                unique_subs = [s for s in unique_subs if s in selected_subs]
                                
                            st.markdown(f"<h5 style='font-size: 1.15rem; font-weight: bold; color: #9b8a5c; margin-top: 10px; margin-bottom: 10px;'>Gráficos das Unidades Subordinadas</h5>", unsafe_allow_html=True)
                            
                            for sub in unique_subs:
                                df_sub = df_rpm[df_rpm["Unidade Principal (Avaliado)"] == sub]
                                if len(df_sub) > 0:
                                    st.markdown(
                                        f"""
                                        <div style='background-color: #26231b; padding: 6px 10px; border-left: 3px solid #ff9f43; border-radius: 4px; margin-top: 15px; margin-bottom: 8px;'>
                                            <span style='font-size: 1.05rem; font-weight: bold; color: #f5f5f5;'>■ Subunidade: {sub}</span>
                                            <span style='font-size: 0.9rem; color: #ff9f43; margin-left: 10px;'>({len(df_sub):,} avaliações)</span>
                                        </div>
                                        """,
                                        unsafe_allow_html=True
                                    )
                                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f_pie:
                                        t_pie = f_pie.name
                                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f_bar:
                                        t_bar = f_bar.name
                                    try:
                                        create_status_pie(df_sub, t_pie)
                                        create_comissao_bar(df_sub, t_bar)
                                        c_sp1, c1, c2, c_sp2 = st.columns([1, 2, 2, 1])
                                        with c1:
                                            st.image(t_pie, use_container_width=True)
                                        with c2:
                                            st.image(t_bar, use_container_width=True)
                                    finally:
                                        for p in [t_pie, t_bar]:
                                            try:
                                                if os.path.exists(p): os.remove(p)
                                            except Exception: pass
                    else:
                        st.info("Sem dados disponíveis.")
            st.markdown("---")
            
            # O botão foi reposicionado no início da página
            pass


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — PAINEL ADMINISTRADOR
# ══════════════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — AUDITORIA DE NOTAS
# ══════════════════════════════════════════════════════════════════════════════
if active_page == "Auditoria de Notas" and sidebar_active_role.upper() in ("ADMINISTRADOR", "GESTOR", "P1", "SADM"):
    st.markdown("### 📊 Auditoria de Notas")
    _role_audit = sidebar_active_role
    active_rpm   = st.session_state.get("simulated_rpm", st.session_state.get("user_rpm", "")) if st.session_state.get("simulation_active", False) else st.session_state.get("user_rpm", "")
    _user_unit  = st.session_state.get("simulated_unit", st.session_state.get("user_unit", "")) if st.session_state.get("simulation_active", False) else st.session_state.get("user_unit", "")

    # Obter caminhos dos arquivos locais e Drive via configuração do ano ativo
    _ano_audit = str(st.session_state.get("selected_year", "2026"))
    _y_cfg = get_active_year_config(_ano_audit, cfg)
    drive_master_xlsx_id = _y_cfg.get("drive_master_xlsx_id", "")

    base_dir = os.path.dirname(os.path.abspath(__file__))
    cache_dir = os.path.join(tempfile.gettempdir(), f"aadp_drive_cache_{_ano_audit}")
    
    possible_master_paths = [
        os.path.join(base_dir, f"DADOS AADP {_ano_audit}", "Analise avaliacoes completa.xlsx"),
        os.path.join(_y_cfg.get("db_path", ""), "Analise avaliacoes completa.xlsx"),
        os.path.join(cache_dir, "Analise avaliacoes completa.xlsx"),
        os.path.join(str(DADOS_DIR), "Analise avaliacoes completa.xlsx"),
        os.path.join(base_dir, "Analise avaliacoes completa.xlsx")
    ]
    master_xlsx_path = next((p for p in possible_master_paths if os.path.exists(p) and os.path.getsize(p) > 0), None)
    if not master_xlsx_path:
        master_xlsx_path = os.path.join(cache_dir, "Analise avaliacoes completa.xlsx")

    # Se o arquivo não existe localmente e não há ID do Google Drive configurado
    if not os.path.exists(master_xlsx_path) and not drive_master_xlsx_id:
        st.error(f"❌ ID da Planilha Mestre no Google Drive não configurado para o ano {_ano_audit}!")
        st.warning(f"⚠️ Configure a chave `drive_master_xlsx_id_{_ano_audit}` (ou `drive_master_xlsx_id`) nas configurações (st.secrets ou config_aadp.json) para habilitar o download automático da auditoria online.")
        st.stop()
            
    with st.spinner("Carregando dados da Planilha Mestre de Auditoria..."):
        df_audit, err = load_audit_excel(master_xlsx_path, drive_master_xlsx_id, ano=_ano_audit)
        if not err and not df_audit.empty:
            df_audit['Tipo AADP'] = df_audit['Sit. Funcional'].apply(_c_aadp_type)
            df_audit = df_audit[df_audit['Tipo AADP'] == global_aadp]
            df_audit.drop(columns=['Tipo AADP'], inplace=True)
            
    if err:
        st.error(f"Erro ao carregar auditoria: {err}")
        st.stop()
        
    # ── Escopo por perfil ─────────────────────────────────────────────────────
    # ADMINISTRADOR / GESTOR → acesso integral (visão de toda a PMMG)
    # P1                     → filtrado pelo RPM do usuário (UDI / UDG)
    # SADM                   → filtrado pela Unidade Principal do usuário
    
    # Garantir que as colunas SIGEF não sejam perdidas e que colunas não se dupliquem
    base_cols = [
        "NR PM", "Posto/Graduação", "Nome Completo", "Quadro",
        "Sit. Funcional",
        "Nome RPM", "Nome Unidade Principal",
        "Conceito", "Reg. Adicional", "Ano Base", "Ord. Almanaque",
        "Qtd Avaliações", "Todas Avaliações Foram Encerradas?", "Nota Final - Média Aritmética",
        "Nota SIRH", "Auditoria"
    ]
    actual_base = [c for c in base_cols if c in df_audit.columns]
    other_cols = [c for c in df_audit.columns if c not in actual_base]
    df_audit_disp = df_audit[actual_base + other_cols].copy()

    if _role_audit == "P1":
        # P1 enxerga apenas os avaliados da sua UDI/UDG (RPM)
        if active_rpm and "Nome RPM" in df_audit_disp.columns:
            df_audit_disp = df_audit_disp[
                df_audit_disp["Nome RPM"].astype(str).str.upper() == str(active_rpm).upper()
            ]
            st.info(f"🔒 Exibindo apenas registros da sua UDI/UDG: **{active_rpm}**")
        else:
            st.warning("⚠️ RPM do usuário não identificado. Contate o administrador.")
            st.stop()

    elif _role_audit == "SADM":
        # SADM enxerga apenas os avaliados da sua Unidade Principal
        if _user_unit and "Nome Unidade Principal" in df_audit_disp.columns:
            df_audit_disp = df_audit_disp[
                df_audit_disp["Nome Unidade Principal"].astype(str).str.upper() == str(_user_unit).upper()
            ]
            st.info(f"🔒 Exibindo apenas registros da sua Unidade Principal: **{_user_unit}**")
        else:
            st.warning("⚠️ Unidade Principal do usuário não identificada. Contate o administrador.")
            st.stop()

    else:
        # ADMINISTRADOR / GESTOR: aplicar filtros opcionais da barra lateral
        if rpm_filter:
            df_audit_disp = df_audit_disp[df_audit_disp["Nome RPM"].isin(rpm_filter)]
        if unid_filter:
            df_audit_disp = df_audit_disp[df_audit_disp["Nome Unidade Principal"].isin(unid_filter)]

    # ── CARDS DE RESUMO DA AUDITORIA ──────────────────────────────────────────
    # 1. Total militares com avaliações abertas e sem nota final encerrada
    c1_mask = (df_audit_disp['Todas Avaliações Foram Encerradas?'].astype(str).str.upper() == 'NAO')
    card1_count = len(df_audit_disp[c1_mask])
    
    # 2. Total militares com todas avaliações encerradas e com nota final
    def is_numeric_grade(val):
        try:
            if val is None or str(val).strip() in ("", "-", "None", "nan"):
                return False
            float(str(val).replace(",", "."))
            return True
        except ValueError:
            return False

    c2_mask = (df_audit_disp['Todas Avaliações Foram Encerradas?'].astype(str).str.upper() == 'SIM') &                (df_audit_disp['Nota Final - Média Aritmética'].apply(is_numeric_grade))
    card2_count = len(df_audit_disp[c2_mask])
    
    # 3. Média da Nota da Unidade (média aritmética de todas as notas finais já encerradas do banco geral)
    df_evals_audit = df_full.copy()
    if _role_audit == "P1":
        if active_rpm:
            df_evals_audit = df_evals_audit[df_evals_audit["Unidade RPM (Avaliado)"].astype(str).str.upper() == str(active_rpm).upper()]
    elif _role_audit == "SADM":
        if _user_unit:
            df_evals_audit = df_evals_audit[df_evals_audit["Unidade Principal (Avaliado)"].astype(str).str.upper() == str(_user_unit).upper()]
    else:
        if rpm_filter:
            df_evals_audit = df_evals_audit[df_evals_audit["Unidade RPM (Avaliado)"].isin(rpm_filter)]
        if unid_filter:
            df_evals_audit = df_evals_audit[df_evals_audit["Unidade Principal (Avaliado)"].isin(unid_filter)]

    if len(df_evals_audit) == 0:
        card3_val = "0,00"
    else:
        # Apenas notas de avaliações encerradas
        df_encerradas = df_evals_audit[df_evals_audit["Status Avaliação"].astype(str).str.upper().str.strip() == "ENCERRADA"]
        grades = pd.to_numeric(df_encerradas["Nota Geral"].astype(str).str.replace(",", "."), errors="coerce").dropna()
        if len(grades) > 0:
            card3_avg = grades.mean()
            import math
            card3_avg_rounded = math.floor(card3_avg * 100 + 0.5) / 100.0
            card3_val = f"{card3_avg_rounded:.2f}".replace(".", ",")
        else:
            card3_val = "0,00"



    # Renderizar os 3 cards lado a lado usando CSS kpi-card
    col_k1, col_k2, col_k3 = st.columns(3)
    with col_k1:
        st.markdown('<div class="kpi-card kpi-aberta">'
                    '<div class="label">MILITARES COM PENDÊNCIAS / AVALIAÇÕES ABERTAS</div>'
                    f'<div class="value">{fmt_num(card1_count)}</div>'
                    '<div class="sub">Militares com pendências de encerramento</div>'
                    '</div>', unsafe_allow_html=True)
    with col_k2:
        st.markdown('<div class="kpi-card kpi-enc">'
                    '<div class="label">MILITARES 100% ENCERRADOS COM NOTA</div>'
                    f'<div class="value">{fmt_num(card2_count)}</div>'
                    '<div class="sub">Militares com nota final encerrada</div>'
                    '</div>', unsafe_allow_html=True)
    with col_k3:
        st.markdown('<div class="kpi-card kpi-ca">'
                    '<div class="label">MÉDIA DAS NOTAS FINAIS</div>'
                    f'<div class="value">{card3_val}</div>'
                    '<div class="sub">Média aritmética das notas encerradas</div>'
                    '</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(f"##### Conteúdo da Planilha Mestre Consolidada ({fmt_num(len(df_audit_disp))} registros)")
    
    # Exibir a planilha de auditoria diretamente!
    safe_df(df_audit_disp, height=540,
            show_download=True, download_type="excel", download_name=f"Auditoria_Notas_Consolidado_{now_br().strftime('%Y%m%d_%H%M')}", download_label="Baixar Resultados Filtrados")


if active_page == "Dados Consolidados" and sidebar_active_role.upper() in ("ADMINISTRADOR", "GESTOR", "P1", "SADM"):
    st.markdown("### 📊 Dados Consolidados")
    st.markdown("---")
    
    st.markdown("""
        <style>
            div[data-testid="stExpander"] div.kpi-card {
                min-height: 145px !important;
                height: 145px !important;
                display: flex !important;
                flex-direction: column !important;
                justify-content: space-between !important;
                padding: 12px 16px !important;
            }
            div[data-testid="stExpander"] div.kpi-card .value {
                margin: auto 0 !important;
            }
            /* Style the expander summary itself to be a premium liquid crystal button */
            div[data-testid="stExpander"] details summary {
                background: linear-gradient(135deg, rgba(128, 128, 128, 0.05) 0%, rgba(20, 20, 20, 0.4) 100%) !important;
                backdrop-filter: blur(10px) !important;
                -webkit-backdrop-filter: blur(10px) !important;
                border: 1px solid rgba(128, 128, 128, 0.25) !important;
                border-top: 1px solid rgba(255, 255, 255, 0.15) !important;
                border-radius: 12px !important;
                padding: 12px 18px !important;
                box-shadow: inset 0 1px 1px rgba(255, 255, 255, 0.05), 0 4px 15px rgba(0, 0, 0, 0.3) !important;
                transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
            }
            div[data-testid="stExpander"] details summary:hover {
                background: linear-gradient(135deg, rgba(155, 138, 92, 0.15) 0%, rgba(30, 30, 30, 0.5) 100%) !important;
                border-color: rgba(155, 138, 92, 0.45) !important;
                box-shadow: inset 0 1px 2px rgba(255, 255, 255, 0.1), 0 0 15px rgba(155, 138, 92, 0.25) !important;
                transform: translateY(-1px) !important;
            }
            /* Highlight the unit name (which we'll make bold in markdown) */
            div[data-testid="stExpander"] details summary strong {
                color: #e5dccb !important;
                font-size: 1.05rem !important;
                font-weight: 800 !important;
                text-shadow: 0 1px 2px rgba(0, 0, 0, 0.8) !important;
                background: rgba(155, 138, 92, 0.20) !important;
                border: 1px solid rgba(155, 138, 92, 0.4) !important;
                padding: 4px 10px !important;
                border-radius: 8px !important;
                box-shadow: inset 0 1px 1px rgba(255, 255, 255, 0.1) !important;
                display: inline-block !important;
                margin-right: 15px !important;
            }
            /* Style code blocks inside summary as crystal liquid sub-buttons */
            div[data-testid="stExpander"] details summary code {
                background: rgba(128, 128, 128, 0.08) !important;
                backdrop-filter: blur(5px) !important;
                -webkit-backdrop-filter: blur(5px) !important;
                border: 1px solid rgba(128, 128, 128, 0.25) !important;
                border-top: 1px solid rgba(255, 255, 255, 0.15) !important;
                color: #a0a0a0 !important;
                font-family: inherit !important;
                font-size: 0.85rem !important;
                font-weight: 700 !important;
                padding: 4px 12px !important;
                border-radius: 30px !important;
                box-shadow: inset 0 1px 0px rgba(255, 255, 255, 0.05), 0 2px 5px rgba(0,0,0,0.2) !important;
                margin-left: 8px !important;
                display: inline-block !important;
                transition: all 0.25s ease !important;
            }
            /* Style the inside tags when hover */
            div[data-testid="stExpander"] details summary:hover code {
                color: #e5dccb !important;
                border-color: rgba(155, 138, 92, 0.3) !important;
                background: rgba(155, 138, 92, 0.12) !important;
            }
        </style>
    """, unsafe_allow_html=True)
    
    drive_master_xlsx_id = cfg.get("drive_master_xlsx_id", "")
    fonte = cfg.get("fonte_dados", "📁 Pasta local / Servidor")
    
    if fonte == "☁️ Google Drive":
        master_xlsx_path = os.path.join(str(DADOS_DIR), "Analise avaliacoes completa.xlsx")
    else:
        master_xlsx_path = os.path.join(str(Path(DADOS_DIR).parent), "Analise avaliacoes completa.xlsx")
        
    df_audit = None
            
    # Obter variáveis de perfil do usuário (considerando simulação)
    _role_consol = sidebar_active_role.upper()
    active_rpm = st.session_state.get("simulated_rpm", st.session_state.get("user_rpm", "")) if st.session_state.get("simulation_active", False) else st.session_state.get("user_rpm", "")
    _user_unit = st.session_state.get("simulated_unit", st.session_state.get("user_unit", "")) if st.session_state.get("simulation_active", False) else st.session_state.get("user_unit", "")

    # Aplicar filtragem de escopo de dados conforme o perfil
    df_evals_source = df_full.copy()
    df_audit_source = None

    if _role_consol == "P1":
        df_evals_source = df_evals_source[df_evals_source["Unidade RPM (Avaliado)"] == active_rpm]
    elif _role_consol == "SADM":
        df_evals_source = df_evals_source[df_evals_source["Unidade Principal (Avaliado)"] == _user_unit]

    def compute_metrics(df_sub_evals, df_sub_audit=None):
        def get_numeric_value(val):
            try:
                if val is None or str(val).strip() in ("", "-", "None", "nan"):
                    return None
                return float(str(val).replace(",", "."))
            except ValueError:
                return None

        if len(df_sub_evals) == 0:
            mil_sim = 0
            mil_nao = 0
            mean_val = 0.0
        else:
            is_enc = (df_sub_evals["Status Avaliação"].astype(str).str.upper().str.strip() == "ENCERRADA")
            
            pm_col = df_sub_evals["nrPM (Avaliado)"]
            all_enc = is_enc.groupby(pm_col).all()

            mil_sim = int(all_enc.sum())
            mil_nao = int((~all_enc).sum())
            
            # Média simples da coluna Nota Geral (apenas encerradas)
            df_enc = df_sub_evals[is_enc]
            grades = pd.to_numeric(df_enc["Nota Geral"].astype(str).str.replace(",", "."), errors="coerce").dropna()
            mean_val = float(grades.mean()) if len(grades) > 0 else 0.0




        n_total = len(df_sub_evals)
        n_enc = (df_sub_evals["Status Avaliação"] == "Encerrada").sum()
        n_aberta = (df_sub_evals["Status Avaliação"] == "Aberta").sum()
        n_parc = (df_sub_evals["Status Avaliação"] == "Parcialmente Encerrada").sum()
        n_hom = (df_sub_evals["Status Avaliação"] == "Homologação").sum()

        return {
            "Avaliações Realizadas": n_total,
            "Comissão Atual": (df_sub_evals["Situação Comissão"] == "Comissão Atual").sum(),
            "Nota Provisória": (df_sub_evals["Situação Comissão"] == "Nota Provisória").sum(),
            "Encerradas": n_enc,
            "Abertas": n_aberta,
            "Parcialmente Encerradas": n_parc,
            "Homologação": n_hom,
            "AV1 Pendente": n_aberta,
            "AV2 Pendente": n_parc,
            "HOM Pendente": n_hom,
            "Militares Encerrados": mil_sim,
            "Militares Pendentes": mil_nao,
            "Média Notas": mean_val
        }
        
    all_rpms = sorted(df_evals_source["Unidade RPM (Avaliado)"].dropna().unique(), key=rpm_sort_key)
    
    # ── PAINEL DE EXPORTAÇÃO CONSOLIDADA POR PERFIL ──────────────────────────
    col_title_consol, col_btn_consol = st.columns([3, 1])
    with col_title_consol:
        st.markdown("<h4 style='font-size: 1.2rem; color: #9b8a5c; margin-bottom: 8px;'>📥 Exportar Relatório Consolidado</h4>", unsafe_allow_html=True)
    btn_container_consol = col_btn_consol.empty()
    
    enable_export = True
    selected_rpms_rel = all_rpms
    selected_subs_rel = None
    
    if _role_consol in ("ADMINISTRADOR", "GESTOR"):
        escopo_rel = st.radio("Selecione o escopo da exportação:", ["🌐 Geral (Todas as RPMs)", "🏢 Por RPM específica"], horizontal=True, key="escopo_rel_consolidado")
        if "específica" in escopo_rel:
            selected_rpms_rel = st.multiselect("Selecione as Unidades RPM/UDG para exportar:", all_rpms, default=[])
            if not selected_rpms_rel:
                st.warning("⚠️ Selecione ao menos uma unidade para habilitar a exportação.")
                enable_export = False
                
    elif _role_consol == "P1":
        escopo_rel = st.radio("Selecione o escopo da exportação:", [f"🌐 Geral (Todas as Subordinadas da {active_rpm})", "🏢 Por Unidade subordinada específica"], horizontal=True, key="escopo_rel_consolidado")
        
        all_subs_in_rpm = sorted(df_evals_source["Unidade Principal (Avaliado)"].dropna().unique())
        selected_subs_rel = all_subs_in_rpm
        if "específica" in escopo_rel:
            selected_subs_rel = st.multiselect("Selecione as Unidades Subordinadas para exportar:", all_subs_in_rpm, default=[])
            if not selected_subs_rel:
                st.warning("⚠️ Selecione ao menos uma unidade subordinada para habilitar a exportação.")
                enable_export = False
                
    elif _role_consol == "SADM":
        selected_subs_rel = [_user_unit]
        st.info(f"ℹ️ O relatório consolidado conterá exclusivamente os dados da sua Unidade: **{_user_unit}**.")
        
    if enable_export:
        # Gerar planilha consolidada em memória
        def export_consolidated_xlsx(df_evals, df_audit, rpms_to_export, role=_role_consol, subs_to_export=selected_subs_rel):
            import io
            import openpyxl
            from openpyxl.styles import Font, Alignment, PatternFill
            
            wb = openpyxl.Workbook()
            default_sheet = wb.active
            wb.remove(default_sheet)
            
            def get_numeric_value(val):
                try:
                    if val is None or str(val).strip() in ("", "-", "None", "nan"):
                        return None
                    return float(str(val).replace(",", "."))
                except ValueError:
                    return None

            def get_eval_grade(row):
                h = get_numeric_value(row.get("Nota Homologação"))
                return h if h is not None else g

            def get_militares_counts(df_sub):
                if len(df_sub) == 0:
                    return 0, 0, 0.0
                is_enc = (df_sub["Status Avaliação"].astype(str).str.upper().str.strip() == "ENCERRADA")
                
                pm_col = df_sub["nrPM (Avaliado)"]
                all_enc = is_enc.groupby(pm_col).all()

                mil_sim = int(all_enc.sum())
                mil_nao = int((~all_enc).sum())
                
                # Média simples da coluna Nota Geral (apenas encerradas)
                df_enc = df_sub[is_enc]
                grades = pd.to_numeric(df_enc["Nota Geral"].astype(str).str.replace(",", "."), errors="coerce").dropna()
                mean_val = float(grades.mean()) if len(grades) > 0 else 0.0
                return mil_sim, mil_nao, mean_val


                    
            headers = [
                "Unidade Principal (RPM/UDG)", "Total Avaliações", "Comissão Atual", "Nota Provisória",
                "Encerradas", "Abertas", "Parcialmente Encerradas", "Homologação",
                "AV1 Pendentes", "AV2 Pendentes", "HOM Pendentes",
                "Militares Encerrados", "Militares Pendentes", "Média Notas"
            ]
            
            if role == "SADM":
                ws1 = wb.create_sheet("Consolidado Unidade")
                headers_sadm = ["Unidade Principal", "Total Avaliações", "Comissão Atual", "Nota Provisória",
                                "Encerradas", "Abertas", "Parcialmente Encerradas", "Homologação",
                                "AV1 Pendentes", "AV2 Pendentes", "HOM Pendentes",
                                "Militares Encerrados", "Militares Pendentes", "Média Notas"]
                ws1.append(headers_sadm)
                sheets_list = [ws1]
            else:
                ws1 = wb.create_sheet("Consolidado RPM")
                ws1.append(headers)
                
                ws2 = wb.create_sheet("Detalhamento Subordinadas")
                headers_sub = ["RPM/UDG", "Unidade Subordinada"] + headers[1:]
                ws2.append(headers_sub)
                sheets_list = [ws1, ws2]
            
            font_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            fill_header = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
            for ws in sheets_list:
                ws.row_dimensions[1].height = 28
                for cell in ws[1]:
                    cell.font = font_bold
                    cell.fill = fill_header
                    cell.alignment = align_center
                    
            if role == "SADM":
                df_sub_evals = df_evals[df_evals["Unidade Principal (Avaliado)"] == _user_unit]
                mil_sim, mil_nao, mean_val = get_militares_counts(df_sub_evals)
                
                sub_row = [
                    _user_unit,
                    len(df_sub_evals),
                    (df_sub_evals["Situação Comissão"] == "Comissão Atual").sum(),
                    (df_sub_evals["Situação Comissão"] == "Nota Provisória").sum(),
                    (df_sub_evals["Status Avaliação"] == "Encerrada").sum(),
                    (df_sub_evals["Status Avaliação"] == "Aberta").sum(),
                    (df_sub_evals["Status Avaliação"] == "Parcialmente Encerrada").sum(),
                    (df_sub_evals["Status Avaliação"] == "Homologação").sum(),
                    (df_sub_evals["Status Avaliação"] == "Aberta").sum(),
                    (df_sub_evals["Status Avaliação"] == "Parcialmente Encerrada").sum(),
                    (df_sub_evals["Status Avaliação"] == "Homologação").sum(),
                    mil_sim,
                    mil_nao,
                    round(mean_val, 2)
                ]
                ws1.append(sub_row)
            else:
                for rpm in rpms_to_export:
                    df_rpm_evals = df_evals[df_evals["Unidade RPM (Avaliado)"] == rpm]
                    
                    if len(df_rpm_evals) == 0:
                        continue
                        
                    mil_sim, mil_nao, mean_val = get_militares_counts(df_rpm_evals)
                    
                    rpm_row = [
                        rpm,
                        len(df_rpm_evals),
                        (df_rpm_evals["Situação Comissão"] == "Comissão Atual").sum(),
                        (df_rpm_evals["Situação Comissão"] == "Nota Provisória").sum(),
                        (df_rpm_evals["Status Avaliação"] == "Encerrada").sum(),
                        (df_rpm_evals["Status Avaliação"] == "Aberta").sum(),
                        (df_rpm_evals["Status Avaliação"] == "Parcialmente Encerrada").sum(),
                        (df_rpm_evals["Status Avaliação"] == "Homologação").sum(),
                        (df_rpm_evals["Status Avaliação"] == "Aberta").sum(),
                        (df_rpm_evals["Status Avaliação"] == "Parcialmente Encerrada").sum(),
                        (df_rpm_evals["Status Avaliação"] == "Homologação").sum(),
                        mil_sim,
                        mil_nao,
                        round(mean_val, 2)
                    ]
                    ws1.append(rpm_row)
                    
                    unique_subs = sorted(df_rpm_evals["Unidade Principal (Avaliado)"].dropna().unique())
                    for sub in unique_subs:
                        if subs_to_export is not None and sub not in subs_to_export:
                            continue
                        sub_evals = df_rpm_evals[df_rpm_evals["Unidade Principal (Avaliado)"] == sub]
                        sub_mil_sim, sub_mil_nao, sub_mean_val = get_militares_counts(sub_evals)
                        
                        sub_row = [
                            rpm,
                            sub,
                            len(sub_evals),
                            (sub_evals["Situação Comissão"] == "Comissão Atual").sum(),
                            (sub_evals["Situação Comissão"] == "Nota Provisória").sum(),
                            (sub_evals["Status Avaliação"] == "Encerrada").sum(),
                            (sub_evals["Status Avaliação"] == "Aberta").sum(),
                            (sub_evals["Status Avaliação"] == "Parcialmente Encerrada").sum(),
                            (sub_evals["Status Avaliação"] == "Homologação").sum(),
                            (sub_evals["Status Avaliação"] == "Aberta").sum(),
                            (sub_evals["Status Avaliação"] == "Parcialmente Encerrada").sum(),
                            (sub_evals["Status Avaliação"] == "Homologação").sum(),
                            sub_mil_sim,
                            sub_mil_nao,
                            round(sub_mean_val, 2)
                        ]
                        ws2.append(sub_row)
                        
            for ws in sheets_list:
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = openpyxl.utils.get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
                    
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return buf.read()
            
        xlsx_consolidated = export_consolidated_xlsx(df_evals_source, df_audit_source, selected_rpms_rel)
        with btn_container_consol:
            st.download_button(
                label="📥 Baixar Planilha Consolidada (Excel .xlsx)",
                data=xlsx_consolidated,
                file_name=f"Relatorio_Consolidado_AADP_{now_br().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                key="btn_download_consolidado"
            )
        
    st.markdown("---")
    
    # ── GRÁFICO GERAL DE RPMs ─────────────────────────────────────────────────
    chart_data = []
    all_grades = pd.to_numeric(df_evals_source["Nota Geral"].astype(str).str.replace(",", "."), errors="coerce").dropna()
    general_avg = float(all_grades.mean()) if len(all_grades) > 0 else 0.0

    for rpm in all_rpms:
        df_rpm_evals = df_evals_source[df_evals_source["Unidade RPM (Avaliado)"] == rpm]
        if len(df_rpm_evals) > 0:
            rpm_grades = pd.to_numeric(df_rpm_evals["Nota Geral"].astype(str).str.replace(",", "."), errors="coerce").dropna()
            rpm_avg = float(rpm_grades.mean()) if len(rpm_grades) > 0 else 0.0
            chart_data.append({"RPM": rpm, "Média": rpm_avg})
            
    df_chart = pd.DataFrame(chart_data)
    if not df_chart.empty:
        # Calcular y_min dinâmico para zoom elegante entre 9 e 10
        min_val = df_chart["Média"].min()
        import math
        y_min = max(0.0, math.floor(min_val * 10) / 10.0 - 0.1)
        if y_min > 8.5:
            y_min = 8.5  # dá uma margem elegante abaixo de 9.0

        fig_gen = px.bar(
            df_chart, 
            x="RPM", 
            y="Média", 
            text=[f"{v:.2f}".replace(".", ",") for v in df_chart["Média"]],
            labels={"Média": "Média das Notas", "RPM": "RPM/UDG"},
            color_discrete_sequence=["#9b8a5c"]
        )
        # Linha tracejada da Média Geral
        fig_gen.add_hline(
            y=general_avg, 
            line_dash="dash", 
            line_color="#ff6b6b"
        )
        # Rótulo de texto fora do gráfico (margem direita)
        fig_gen.add_annotation(
            xref="paper",
            yref="y",
            x=1.01,
            y=general_avg,
            text=f"Média Geral<br><b>{general_avg:.2f}</b>".replace(".", ","),
            showarrow=False,
            font=dict(color="#ff6b6b", size=9, family="sans-serif"),
            xanchor="left",
            yanchor="middle",
            align="left"
        )
        fig_gen.update_traces(
            textposition='outside',
            textfont=dict(color="#cbbb8f", size=10, family="sans-serif"),
            marker_line_color="#cbbb8f",
            marker_line_width=1.5,
            opacity=0.85,
            cliponaxis=False
        )
        fig_gen.update_layout(
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            yaxis=dict(
                range=[y_min, 10.05],
                gridcolor="rgba(255, 255, 255, 0.05)",
                zeroline=False,
                tickfont=dict(color="#a0a0a0", size=10),
                title=None
            ),
            xaxis=dict(
                gridcolor="rgba(255, 255, 255, 0.05)",
                tickfont=dict(color="#a0a0a0", size=10),
                title=None
            ),
            bargap=0.45,
            margin=dict(l=30, r=100, t=30, b=50),  # Margem direita expandida para o rótulo
            height=250
        )
        st.plotly_chart(fig_gen, use_container_width=True)

    st.markdown("Clique em uma Unidade Principal (RPM/UDG) para visualizar seu resumo consolidado e expandir suas Unidades Subordinadas.")
    
    for rpm in all_rpms:
        df_rpm_evals = df_evals_source[df_evals_source["Unidade RPM (Avaliado)"] == rpm]
        df_rpm_audit = None
        
        rpm_metrics = compute_metrics(df_rpm_evals, df_rpm_audit)
        
        total_evals = rpm_metrics["Avaliações Realizadas"]
        enc_evals = rpm_metrics["Encerradas"]
        avg_grade = rpm_metrics["Média Notas"]
        
        avg_grade_str = f"{avg_grade:.2f}".replace(".", ",")
        exp_header = f"🏢 **{rpm}** &nbsp;&nbsp;&nbsp;&nbsp; `Total: {fmt_num(total_evals)}` &nbsp;&nbsp; `Encerradas: {fmt_num(enc_evals)}` &nbsp;&nbsp; `Média: {avg_grade_str}`"
        
        with st.expander(exp_header):
            # Linha 1: 4 Cards principais
            k_c1, k_c2, k_c3, k_c4 = st.columns(4)
            with k_c1:
                st.markdown(f'<div class="kpi-card kpi-total">'
                            f'<div class="label">Total Avaliações</div>'
                            f'<div class="value">{fmt_num(total_evals)}</div>'
                            f'<div class="sub">Comissão: {fmt_num(rpm_metrics["Comissão Atual"])} CA | {fmt_num(rpm_metrics["Nota Provisória"])} NP</div>'
                            f'</div>', unsafe_allow_html=True)
            with k_c2:
                st.markdown(f'<div class="kpi-card kpi-aberta">'
                            f'<div class="label">Pendências Funcionais</div>'
                            f'<div class="value" style="font-size: 1.05rem; line-height: 1.35; font-weight: 700; margin-top: 6px; margin-bottom: 6px;">'
                            f'<span style="color: #ff6b6b !important;">AVALIADOR 1: {fmt_num(rpm_metrics["AV1 Pendente"])}</span><br>'
                            f'<span style="color: #ff9f43 !important;">AVALIADOR 2: {fmt_num(rpm_metrics["AV2 Pendente"])}</span><br>'
                            f'<span style="color: #ffd257 !important;">HOMOLOGADOR: {fmt_num(rpm_metrics["HOM Pendente"])}</span>'
                            f'</div>'
                            f'</div>', unsafe_allow_html=True)
            with k_c3:
                st.markdown(f'<div class="kpi-card kpi-parc">'
                            f'<div class="label">Militares Pendentes</div>'
                            f'<div class="value">{fmt_num(rpm_metrics["Militares Pendentes"])}</div>'
                            f'<div class="sub">Militar com avaliações não encerradas</div>'
                            f'</div>', unsafe_allow_html=True)
            with k_c4:
                avg_grade_str = f"{avg_grade:.2f}".replace(".", ",")
                st.markdown(f'<div class="kpi-card kpi-ca">'
                            f'<div class="label">Média das Notas</div>'
                            f'<div class="value">{avg_grade_str}</div>'
                            f'<div class="sub">Média aritmética final consolidada</div>'
                            f'</div>', unsafe_allow_html=True)
                            
            st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)
            
            # Linha 2: 4 Quadrantes de Status
            q_c1, q_c2, q_c3, q_c4 = st.columns(4)
            with q_c1:
                st.markdown(f'<div class="kpi-card kpi-enc">'
                            f'<div class="label">🟢 Encerradas</div>'
                            f'<div class="value">{fmt_num(enc_evals)}</div>'
                            f'<div class="sub">Avaliações finalizadas</div>'
                            f'</div>', unsafe_allow_html=True)
            with q_c2:
                st.markdown(f'<div class="kpi-card kpi-aberta">'
                            f'<div class="label">🔴 Abertas</div>'
                            f'<div class="value">{fmt_num(rpm_metrics["Abertas"])}</div>'
                            f'<div class="sub">Sem nota ou AV1 pendente</div>'
                            f'</div>', unsafe_allow_html=True)
            with q_c3:
                st.markdown(f'<div class="kpi-card kpi-parc">'
                            f'<div class="label">🟠 Parcialmente Enc.</div>'
                            f'<div class="value">{fmt_num(rpm_metrics["Parcialmente Encerradas"])}</div>'
                            f'<div class="sub">Aguardando AV2</div>'
                            f'</div>', unsafe_allow_html=True)
            with q_c4:
                st.markdown(f'<div class="kpi-card kpi-hom">'
                            f'<div class="label">🟡 Homologação</div>'
                            f'<div class="value">{fmt_num(rpm_metrics["Homologação"])}</div>'
                            f'<div class="sub">Aguardando Homologador</div>'
                            f'</div>', unsafe_allow_html=True)
                            
            st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)
            st.markdown("##### 📁 Detalhamento de Unidades Subordinadas")
            
            sub_rows = []
            unique_subs = sorted(df_rpm_evals["Unidade Principal (Avaliado)"].dropna().unique())
            # SADM enxerga somente a sua respectiva unidade
            if _role_consol == "SADM":
                unique_subs = [s for s in unique_subs if s == _user_unit]
                
            for sub in unique_subs:
                sub_evals = df_rpm_evals[df_rpm_evals["Unidade Principal (Avaliado)"] == sub]
                sub_audit = None
                
                metrics = compute_metrics(sub_evals, sub_audit)
                row = {"Unidade Subordinada": sub}
                row.update(metrics)
                sub_rows.append(row)
                
            if sub_rows:
                df_sub_table = pd.DataFrame(sub_rows)
                
                # Renderizar o gráfico das subunidades com zoom e design premium
                st.markdown("###### 📊 Média de Notas por Subunidade")
                
                # Calcular y_min dinâmico para zoom elegante das subunidades
                sub_min_val = df_sub_table["Média Notas"].min()
                import math
                y_min_sub = max(0.0, math.floor(sub_min_val * 10) / 10.0 - 0.1)
                if y_min_sub > 8.5:
                    y_min_sub = 8.5  # dá uma margem elegante abaixo de 9.0

                fig_sub = px.bar(
                    df_sub_table,
                    x="Unidade Subordinada",
                    y="Média Notas",
                    text=[f"{v:.2f}".replace(".", ",") for v in df_sub_table["Média Notas"]],
                    labels={"Média Notas": "Média das Notas", "Unidade Subordinada": "Subunidade"},
                    color_discrete_sequence=["#9b8a5c"]
                )
                # Linha tracejada da Média da RPM
                fig_sub.add_hline(
                    y=avg_grade,
                    line_dash="dash",
                    line_color="#ff9f43"
                )
                # Rótulo de texto fora do gráfico (margem direita)
                fig_sub.add_annotation(
                    xref="paper",
                    yref="y",
                    x=1.01,
                    y=avg_grade,
                    text=f"Média {rpm}<br><b>{avg_grade:.2f}</b>".replace(".", ","),
                    showarrow=False,
                    font=dict(color="#ff9f43", size=9, family="sans-serif"),
                    xanchor="left",
                    yanchor="middle",
                    align="left"
                )
                fig_sub.update_traces(
                    textposition='outside',
                    textfont=dict(color="#cbbb8f", size=10, family="sans-serif"),
                    marker_line_color="#cbbb8f",
                    marker_line_width=1.5,
                    opacity=0.85,
                    cliponaxis=False
                )
                fig_sub.update_layout(
                    template="plotly_dark",
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    yaxis=dict(
                        range=[y_min_sub, 10.05],
                        gridcolor="rgba(255, 255, 255, 0.05)",
                        zeroline=False,
                        tickfont=dict(color="#a0a0a0", size=10),
                        title=None
                    ),
                    xaxis=dict(
                        gridcolor="rgba(255, 255, 255, 0.05)",
                        tickfont=dict(color="#a0a0a0", size=10),
                        title=None
                    ),
                    bargap=0.45,
                    margin=dict(l=30, r=100, t=30, b=50),  # Margem direita expandida para o rótulo
                    height=220
                )
                st.plotly_chart(fig_sub, use_container_width=True)
                
                # Reordenar colunas para colocar a média das notas logo após o nome da unidade subordinada
                cols_ordered = [
                    "Unidade Subordinada", "Média Notas", "Avaliações Realizadas",
                    "Comissão Atual", "Nota Provisória", "Encerradas", "Abertas",
                    "Parcialmente Encerradas", "Homologação", "AV1 Pendente",
                    "AV2 Pendente", "HOM Pendente", "Militares Encerrados", "Militares Pendentes"
                ]
                cols_ordered = [c for c in cols_ordered if c in df_sub_table.columns]
                df_sub_table = df_sub_table[cols_ordered]
                
                df_sub_table_disp = df_sub_table.rename(columns={
                    "Média Notas": "Média das<br>Notas",
                    "Avaliações Realizadas": "Total<br>Avaliações",
                    "Comissão Atual": "Comissão<br>Atual",
                    "Nota Provisória": "Nota<br>Provisória",
                    "Encerradas": "Encerradas",
                    "Abertas": "Abertas",
                    "Parcialmente Encerradas": "Parcialmente<br>Encerradas",
                    "Homologação": "Homologação",
                    "AV1 Pendente": "Avaliador 1<br>Pendente",
                    "AV2 Pendente": "Avaliador 2<br>Pendente",
                    "HOM Pendente": "Homologador<br>Pendente",
                    "Militares Encerrados": "Militares<br>Encerrados",
                    "Militares Pendentes": "Militares<br>Pendentes"
                })
                
                styler = df_sub_table_disp.style.format({
                    "Média das<br>Notas": lambda x: f"{x:.2f}".replace(".", ",")
                })
                
                try:
                    styler = styler.hide(axis="index")
                except Exception:
                    try:
                        styler = styler.hide_index()
                    except Exception:
                        pass
                
                styler = styler.set_table_attributes('class="consolidated-table"')
                
                cols_gold = [c for c in ["Média das<br>Notas"] if c in df_sub_table_disp.columns]
                cols_grey = [c for c in ["Total<br>Avaliações", "Comissão<br>Atual", "Nota<br>Provisória", "Avaliador 1<br>Pendente", "Avaliador 2<br>Pendente", "Homologador<br>Pendente"] if c in df_sub_table_disp.columns]
                cols_green = [c for c in ["Encerradas", "Militares<br>Encerrados"] if c in df_sub_table_disp.columns]
                cols_red = [c for c in ["Abertas", "Militares<br>Pendentes"] if c in df_sub_table_disp.columns]
                cols_orange = [c for c in ["Parcialmente<br>Encerradas"] if c in df_sub_table_disp.columns]
                cols_yellow = [c for c in ["Homologação"] if c in df_sub_table_disp.columns]
                
                if cols_gold: styler = styler.map(lambda x: 'color: #9b8a5c; font-weight: bold; text-align: center;', subset=cols_gold)
                if cols_grey: styler = styler.map(lambda x: 'color: #a0a0a0; text-align: center;', subset=cols_grey)
                if cols_green: styler = styler.map(lambda x: 'color: #7bed9f; font-weight: bold; text-align: center;', subset=cols_green)
                if cols_red: styler = styler.map(lambda x: 'color: #ff6b6b; font-weight: bold; text-align: center;', subset=cols_red)
                if cols_orange: styler = styler.map(lambda x: 'color: #ff9f43; font-weight: bold; text-align: center;', subset=cols_orange)
                if cols_yellow: styler = styler.map(lambda x: 'color: #ffd257; font-weight: bold; text-align: center;', subset=cols_yellow)
                
                html_table = styler.to_html(escape=False)
                
                html_content = (
                    "<style>\n"
                    ".consolidated-table-container {\n"
                    "    width: 100%;\n"
                    "    overflow-x: auto;\n"
                    "    margin: 15px 0;\n"
                    "}\n"
                    ".consolidated-table {\n"
                    "    width: 100%;\n"
                    "    border-collapse: collapse;\n"
                    "    font-size: 0.82rem;\n"
                    "    font-family: inherit;\n"
                    "    background: rgba(30, 30, 30, 0.4);\n"
                    "    backdrop-filter: blur(8px);\n"
                    "    -webkit-backdrop-filter: blur(8px);\n"
                    "    border-radius: 8px;\n"
                    "    overflow: hidden;\n"
                    "}\n"
                    ".consolidated-table th {\n"
                    "    background-color: rgba(155, 138, 92, 0.15) !important;\n"
                    "    color: #e5dccb !important;\n"
                    "    font-weight: 700 !important;\n"
                    "    padding: 8px 6px !important;\n"
                    "    text-align: center !important;\n"
                    "    border: 1px solid rgba(128, 128, 128, 0.2) !important;\n"
                    "    white-space: normal !important;\n"
                    "    word-break: normal !important;\n"
                    "    vertical-align: middle !important;\n"
                    "    line-height: 1.2 !important;\n"
                    "}\n"
                    ".consolidated-table td {\n"
                    "    padding: 8px 6px !important;\n"
                    "    text-align: center !important;\n"
                    "    border: 1px solid rgba(128, 128, 128, 0.15) !important;\n"
                    "    vertical-align: middle !important;\n"
                    "}\n"
                    ".consolidated-table tr:nth-child(even) {\n"
                    "    background-color: rgba(255, 255, 255, 0.01);\n"
                    "}\n"
                    ".consolidated-table tr:hover {\n"
                    "    background-color: rgba(155, 138, 92, 0.05);\n"
                    "}\n"
                    "</style>\n"
                    f"<div class=\"consolidated-table-container\">\n{html_table}\n</div>"
                )
                st.markdown(html_content, unsafe_allow_html=True)
            else:
                st.info("Nenhuma unidade subordinada encontrada.")



if active_page == "Painel Administrador" and st.session_state.user_role == "ADMINISTRADOR":
    st.markdown("### ⚙️ Painel Administrador")
    
    tab_pending, tab_active, tab_logs = st.tabs([
        "⏳ Solicitações de Cadastro Pendentes",
        "👥 Gerenciar/Alterar Usuários Cadastrados",
        "📜 Auditoria"
    ])
    
    # ── 1) Solicitações de Cadastro Pendentes ────────────────────────────────
    with tab_pending:
        col_t1, col_t2 = st.columns([3, 1])
        with col_t1:
            st.markdown('#### ⏳ Solicitações de Cadastro Pendentes')
        with col_t2:
            if st.button('🔄 Atualizar Lista', use_container_width=True, key='refresh_pendentes'):
                refresh_db_cache()
                st.rerun()
        pend_list = db_get_pending_users()
        
        if not pend_list:
            st.info("Não há solicitações pendentes no momento.")
        else:
            for pm, name, rank, rpm, unit, function, created_at in pend_list:
                pm_str = str(pm)
                with st.container():
                    st.markdown(f"**{rank} {name}** (PM: `{pm_str}`)")
                    st.markdown(f"RPM: `{rpm}` | Unidade: `{unit}` | Função: `{function}` | Data: `{created_at}`")
                    
                    c_role = st.selectbox("Selecione o perfil de acesso:", ["P1", "SADM", "Gestor", "ADMINISTRADOR"], key=f"role_{pm_str}")
                    
                    col_ap, col_rec, _ = st.columns([1, 1, 4])
                    if col_ap.button("✅ Autorizar Acesso", key=f"ap_{pm_str}", type="primary"):
                        target_rpm = "Gestor" if c_role in ("Gestor", "ADMINISTRADOR") else rpm
                        db_approve_user(pm_str, c_role, target_rpm)
                        log_action("ADM", "AUTORIZAR_ACESSO", f"Usuario {pm_str} ({name}) aprovado como {c_role}")
                        st.success(f"Acesso de {name} autorizado com sucesso!")
                        st.rerun()
                    if col_rec.button("❌ Recusar", key=f"rec_{pm_str}", type="secondary"):
                        db_reject_user(pm_str)
                        log_action("ADM", "RECUSAR_CADASTRO", f"Cadastro do usuario {pm_str} ({name}) recusado")
                        st.warning(f"Cadastro de {name} recusado!")
                        st.rerun()
                    st.markdown("---")
                    
    # ── 2) Gerenciar/Alterar Usuários Cadastrados ────────────────────────────
    with tab_active:
        st.markdown("#### 👥 Gerenciar / Alterar Usuários Cadastrados")
        
        # --- Sincronização manual com SIGEF ---
        if st.button("🔄 Sincronizar Cadastros com SIGEF (Atualiza Órgãos/Unidades)", use_container_width=True, key="btn_sync_sigef"):
            with st.spinner("⏳ Sincronizando dados com a base SIGEF..."):
                sync_users_with_sigef()
            st.success("✅ Cadastros sincronizados com a base SIGEF com sucesso!")
            st.rerun()
            
        st.markdown("---")
        
        # --- SIMULADOR DE VISÃO DE TELA ---
        st.markdown("##### 🕵️ Simulador de Visão de Tela")
        sim_users = db_get_simulator_users()
        
        if st.session_state.get("simulation_active", False):
            st.warning(f"🕵️ **Simulação Ativa**: Você está visualizando o sistema como **{st.session_state.simulated_name}** ({st.session_state.simulated_role}).")
            if st.button("❌ Parar Simulação / Voltar ao normal", type="secondary", use_container_width=True, key="btn_stop_sim"):
                st.session_state.simulation_active = False
                st.session_state.simulated_pm = ""
                st.session_state.simulated_name = ""
                st.session_state.simulated_role = ""
                st.session_state.simulated_rpm = ""
                st.session_state.simulated_unit = ""
                st.session_state.active_page = "Painel Administrador"
                log_action("ADM", "ENCERRAR_SIMULACAO", "Simulacao desativada")
                st.rerun()
        else:
            sim_options = []
            for u in sim_users:
                sim_options.append(f"{u[2]} {u[1]} ({u[3]} - {u[4]}) [PM: {u[0]}]")
            
            selected_sim_user = st.selectbox(
                "Selecione um usuário cadastrado para simular a visão dele:",
                sim_options,
                index=None,
                placeholder="Escolha um usuário...",
                key="sim_user_select"
            )
            
            if selected_sim_user:
                pm_match = re.search(r'\[PM:\s*(\w+)\]', selected_sim_user)
                if pm_match:
                    target_pm = pm_match.group(1)
                    if st.button("🎭 Simular visão deste usuário", type="primary", use_container_width=True, key="btn_start_sim"):
                        for u in sim_users:
                            if str(u[0]).strip() == str(target_pm).strip():
                                st.session_state.simulation_active = True
                                st.session_state.simulated_pm = u[0]
                                st.session_state.simulated_name = f"{u[2]} {u[1]}"
                                st.session_state.simulated_role = u[3]
                                st.session_state.simulated_rpm = u[4]
                                st.session_state.simulated_unit = u[5]
                                st.session_state.active_page = "Análise Gráfica"
                                log_action("ADM", "INICIAR_SIMULACAO", f"Simulando usuario {u[0]}")
                                st.success(f"Iniciando simulação de visão de: {u[2]} {u[1]}")
                                st.rerun()
                                
        st.markdown("---")
        
        # --- LISTA E ALTERAÇÃO DE USUÁRIOS ATIVOS ---
        st.markdown("##### 👥 Relação de Usuários Cadastrados")
        
        # Filtro por Nº PM
        filter_pm = st.text_input("🔍 Consultar por Nº PM (deixe em branco para ver todos):", "", key="active_users_filter_pm").strip()
        
        active_list = db_get_active_users()
        if not active_list:
            st.info("Nenhum usuário ativo cadastrado.")
        else:
            df_act = pd.DataFrame(active_list, columns=["Nº PM", "Posto/Grad.", "Nome", "RPM", "Unidade Principal", "Setor", "Perfil", "Data Cadastro"])
            
            # Reordenar para Nº PM / Posto/Grad. / Nome primeiro
            col_order = ["Nº PM", "Posto/Grad.", "Nome", "RPM", "Unidade Principal", "Setor", "Perfil", "Data Cadastro"]
            df_act = df_act[[c for c in col_order if c in df_act.columns]]
            
            # Filtragem se houver termo digitado
            if filter_pm:
                df_act = df_act[df_act["Nº PM"].astype(str).str.contains(filter_pm)]
                
            if df_act.empty:
                st.info("Nenhum usuário cadastrado encontrado com o Nº PM informado.")
            else:
                st.dataframe(clean_none_values(df_act), use_container_width=True, hide_index=True)
                
                st.markdown("##### ⚙️ Gerenciar / Alterar Cadastro:")
                
                user_options = [f"{row['Posto/Grad.']} {row['Nome']} (PM: {row['Nº PM']})" for _, row in df_act.iterrows()]
                selected_user_label = st.selectbox("Escolha o usuário para gerenciar:", user_options, key="manage_user_select")
                
                if selected_user_label:
                    m_pm = selected_user_label.split(" (PM: ")[1].rstrip(")")
                    matching_rows = df_act[df_act["Nº PM"].astype(str).str.strip() == str(m_pm).strip()]
                    if not matching_rows.empty:
                        row_sel = matching_rows.iloc[0]
                        
                        st.write(f"**Dados Atuais:** Perfil: `{row_sel['Perfil']}` | RPM: `{row_sel['RPM']}`")
                        
                        col_role, col_rpm = st.columns(2)
                        with col_role:
                            new_role = st.selectbox("Alterar Perfil de Acesso:", ["P1", "SADM", "Gestor", "ADMINISTRADOR"], 
                                                    index=["P1", "SADM", "Gestor", "ADMINISTRADOR"].index(row_sel["Perfil"]) if row_sel["Perfil"] in ["P1", "SADM", "Gestor", "ADMINISTRADOR"] else 0,
                                                    key=f"edit_role_{m_pm}")
                        with col_rpm:
                            rpm_choices = ["Gestor"] + [f"{i} RPM" for i in range(1, 20)] + [
                                "AM-ALMG", "AM-TJMG", "APM", "AUD SET", "CME", "COMAVE", "CPE", 
                                "CPM", "DAL", "DCO", "DEE", "DF", "DINT", "DOP", "DPS", "DRH", "DTS", 
                                "EMPM/SCG", "GCG", "GMG"
                            ]
                            try:
                                cur_index = rpm_choices.index(row_sel["RPM"])
                            except ValueError:
                                cur_index = 0
                                
                            new_rpm = st.selectbox("Alterar RPM / Diretoria / UDG:", rpm_choices, index=cur_index, key=f"edit_rpm_{m_pm}")
                        
                        col_save, col_rev = st.columns(2)
                        with col_save:
                            if st.button("💾 Salvar Alterações", type="primary", use_container_width=True, key=f"save_edit_{m_pm}"):
                                db_update_user_role_rpm(m_pm, new_role, new_rpm)
                                log_action("ADM", "ALTERAR_CADASTRO", f"Usuario {m_pm} alterado para Perfil: {new_role}, RPM: {new_rpm}")
                                st.success("✅ Alterações salvas com sucesso!")
                                st.rerun()
                        with col_rev:
                            if st.button("🚫 Revogar Acesso", type="secondary", use_container_width=True, key=f"revoke_edit_{m_pm}"):
                                db_revoke_user(m_pm)
                                log_action("ADM", "REVOGAR_ACESSO", f"Acesso do usuario {m_pm} revogado")
                                st.warning("🚫 Acesso revogado com sucesso!")
                                st.rerun()
                    else:
                        st.error("Erro ao selecionar usuário. O cache pode estar desatualizado.")
                        st.stop()
                        
    # ── 3) Auditoria ─────────────────────────────────────────────────────────
    with tab_logs:
        st.markdown("#### 📜 Auditoria de Atividades / Logs")
    
        # Buscar lista de usuários únicos que possuem ações nos logs
        logged_pms = db_get_logs_pms()
    
        # Mapear PM para Nome de forma amigável
        user_map = {"Todos": "Todos os Usuários"}
        for pm_id in logged_pms:
            u_row = db_get_user_info(pm_id)
            if u_row:
                user_map[pm_id] = f"{u_row[0]} {u_row[1]} (PM: {pm_id})"
            else:
                user_map[pm_id] = f"PM: {pm_id}"
            
        # Componente de filtro de militares
        sel_log_user = st.selectbox(
            "Filtrar logs por militar:",
            list(user_map.keys()),
            format_func=lambda x: user_map[x],
            key="filter_log_user"
        )
    
        # Filtro de período por data
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            start_date = st.date_input("Data de início:", value=now_br().date() - timedelta(days=7), key="log_start_date")
        with col_d2:
            end_date = st.date_input("Data de fim:", value=now_br().date(), key="log_end_date")
        
        start_str = start_date.strftime("%Y-%m-%d")
        end_str = end_date.strftime("%Y-%m-%d")
    
        df_logs = db_get_logs_df(sel_log_user, start_str, end_str)
    
        if df_logs.empty:
            st.info("Nenhum log registrado para a seleção.")
        else:
            df_logs_disp = df_logs.copy()
            df_logs_disp.index = range(1, len(df_logs_disp) + 1)
            st.dataframe(clean_none_values(df_logs_disp), use_container_width=True)
        
        dl_log_xlsx = df_to_xlsx(df_logs)
        dl_log_csv = df_logs.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
    
        col_l1, col_l2 = st.columns(2)
        with col_l1:
            st.download_button(
                "⬇️ Baixar Logs Filtrados (Excel .xlsx)",
                dl_log_xlsx,
                f"logs_{sel_log_user}_{now_br().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_logs_xlsx"
            )
        with col_l2:
            st.download_button(
                "⬇️ Baixar Logs Filtrados (CSV)",
                dl_log_csv,
                f"logs_{sel_log_user}_{now_br().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                key="dl_logs_csv"
            )

    # ── 4) Comissões ─────────────────────────────────────────────────────────
if active_page == "Comissões" and sidebar_active_role.upper() in ("ADMINISTRADOR", "GESTOR", "P1", "SADM"):
    st.markdown("### ⚖️ Análise de Comissões")
    
    @st.cache_data(show_spinner=False)
    def load_comissoes_tab_data(_db_path, _drive_com_id, _drive_si_id, _ano="2026"):
        import pandas as pd
        import os
        import tempfile
        
        cache_dir = os.path.join(tempfile.gettempdir(), f"aadp_drive_cache_{_ano}")
        if _drive_com_id and _drive_si_id:
            sigef_path = os.path.join(cache_dir, "SIGEF.csv")
            comissao_path = os.path.join(cache_dir, "comissao.csv")
            if not os.path.exists(comissao_path) or os.path.getsize(comissao_path) == 0:
                _baixar_drive(_drive_com_id, comissao_path)
        else:
            sigef_path = os.path.join(_db_path, "SIGEF.csv")
            comissao_path = os.path.join(_db_path, "avaliacoes.csv")
            
        try:
            df_sigef = pd.read_csv(sigef_path, sep=';', encoding='cp1252', dtype=str, on_bad_lines='skip', index_col=False)
        except Exception:
            df_sigef = pd.DataFrame()
            
        try:
            df_com = pd.read_csv(comissao_path, sep=';', encoding='cp1252', dtype=str, on_bad_lines='warn', index_col=False)
            df_com = df_com.rename(columns={
                "Unidade RPM (Avaliado)": "Unidade RPM Atual (Avaliado)",
                "Unidade Principal (Avaliado)": "Unidade Principal Atual (Avaliado)",
                "Local/Unidade (Avaliado)": "Local/Unidade Atual (Avaliado)"
            })
        except Exception:
            df_com = pd.DataFrame()
            
        return df_sigef, df_com

    @st.cache_data(show_spinner=False)
    def load_cdp_status_set(_db_path, _drive_geral_id, _drive_metas_id, _ano="2026"):
        import csv, os
        base_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
        cache_dir = os.path.join(base_dir, ".cache_aadp", str(_ano))
        os.makedirs(cache_dir, exist_ok=True)

        def _c_pm_clean(val):
            if not val or str(val).strip() in ("", "-", "nan", "none", "None", "<NA>"):
                return ""
            try:
                return str(int(float(str(val).strip())))
            except Exception:
                s = str(val).strip().lstrip("0")
                return s if s else "0"

        possible_geral = [
            os.path.join(base_dir, f"DADOS AADP {_ano}", "geral.csv"),
            os.path.join(cache_dir, "geral.csv"),
            os.path.join(_db_path or "", "geral.csv"),
            os.path.join(base_dir, "dados", "geral.csv"),
            os.path.join(base_dir, "geral.csv")
        ]
        geral_path = next((p for p in possible_geral if os.path.exists(p) and os.path.getsize(p) > 0), None)
        if not geral_path and _drive_geral_id:
            geral_dest = os.path.join(cache_dir, "geral.csv")
            try:
                _baixar_drive(_drive_geral_id, geral_dest)
                if os.path.exists(geral_dest) and os.path.getsize(geral_dest) > 0:
                    geral_path = geral_dest
            except Exception:
                pass

        possible_metas = [
            os.path.join(base_dir, f"DADOS AADP {_ano}", f"Metas e acompanhamentos {_ano} Completo.csv"),
            os.path.join(base_dir, f"DADOS AADP {_ano}", "metas.csv"),
            os.path.join(cache_dir, f"Metas e acompanhamentos {_ano} Completo.csv"),
            os.path.join(cache_dir, "metas.csv"),
            os.path.join(_db_path or "", f"Metas e acompanhamentos {_ano} Completo.csv"),
            os.path.join(_db_path or "", "metas.csv"),
            os.path.join(base_dir, "dados", "metas.csv"),
            os.path.join(base_dir, "metas.csv")
        ]
        metas_path = next((p for p in possible_metas if os.path.exists(p) and os.path.getsize(p) > 0), None)
        if not metas_path and _drive_metas_id:
            metas_dest = os.path.join(cache_dir, "metas.csv")
            try:
                _baixar_drive(_drive_metas_id, metas_dest)
                if os.path.exists(metas_dest) and os.path.getsize(metas_dest) > 0:
                    metas_path = metas_dest
            except Exception:
                pass

        cdp_pms = set()
        if geral_path and os.path.exists(geral_path):
            with open(geral_path, "r", encoding="cp1252", errors="replace") as f_ge:
                r_ge = csv.reader(f_ge, delimiter=";")
                next(r_ge, [])
                for row in r_ge:
                    if len(row) > 5 and row[5].strip() not in ("", "-", "nan", "None", "<NA>"):
                        p = _c_pm_clean(row[1])
                        if p:
                            cdp_pms.add(p)

        if metas_path and os.path.exists(metas_path):
            with open(metas_path, "r", encoding="cp1252", errors="replace") as f_m:
                r_m = csv.reader(f_m, delimiter=";")
                next(r_m, [])
                for row in r_m:
                    if len(row) > 26 and row[26].strip() not in ("", "-", "nan", "None", "<NA>"):
                        p = _c_pm_clean(row[0])
                        if p:
                            cdp_pms.add(p)

        return cdp_pms

    with st.spinner("Carregando bases de dados (SIGEF e Comissões)..."):
        # Resolve os argumentos usando cfg_to_use diretamente, para evitar falha caso a função não receba
        cfg_to_use = load_config()
        _d_path = cfg_to_use.get("db_path", "")
        _d_com_id = cfg_to_use.get("drive_com_id", "")
        _d_si_id = cfg_to_use.get("drive_si_id", "")
        _act_y = str(st.session_state.get("selected_year", "2026"))
        _y_cfg = get_active_year_config(_act_y, cfg_to_use)
        df_sigef, df_com = load_comissoes_tab_data(_y_cfg["db_path"], _y_cfg["drive_com_id"], _y_cfg["drive_si_id"], _ano=_act_y)
        cdp_pms_set = load_cdp_status_set(_y_cfg["db_path"], _y_cfg["drive_geral_id"], _y_cfg["drive_metas_id"], _ano=_act_y)

    if df_sigef.empty:
        st.error("Erro: SIGEF.csv não encontrado ou ilegível.")
    elif df_com.empty or 'Posto/Graduação (Avaliador2)' not in df_com.columns:
        st.error("Erro: Planilha de Comissões não pôde ser carregada ou possui formato inválido.")
        st.warning("Se você estiver na nuvem, o Google Drive pode ter bloqueado temporariamente o link por excesso de downloads (Cota de Tráfego excedida). Aguarde alguns minutos ou verifique se o arquivo está como 'Qualquer pessoa com o link'.")
        st.stop()
    else:
        def _c_get_pm(val):
            import re
            if pd.isna(val): return ""
            s = str(val).strip()
            if '-' in s: s = s.split('-')[0].strip()
            s = s.lstrip('0')
            if s.endswith('.0'): s = s[:-2]
            s = re.sub(r'[^0-9]', '', s)
            return s if s not in ("", "nan", "None") else ""

        def _c_classify_com(row):
            av1 = _c_get_pm(row.get('nrPM (Avaliador1)', ''))
            av2 = _c_get_pm(row.get('nrPM (Avaliador2)', ''))
            hom = _c_get_pm(row.get('nrPM (Homologador)', ''))
            
            missing = []
            if not av1: missing.append("AV1")
            if not av2: missing.append("AV2")
            if not hom: missing.append("HOM")
            
            if missing:
                if len(missing) == 3: 
                    return "INCOMPLETA (FALTA TODOS)"
                return f"INCOMPLETA (FALTA {', '.join(missing)})"
                
            pms = {av1, av2, hom}
            if len(pms) == 1:
                return "COMPLETA (COMISSÃO ÚNICA)"
            elif len(pms) == 2:
                return "COMPLETA (2 MEMBROS)"
            elif len(pms) == 3:
                return "COMPLETA (3 MEMBROS)"
                
            return "COMPLETA (PADRÃO INCORRETO)"

        def _c_aadp_category(sit):
            sit = str(sit).strip().upper()
            if sit in ["ATIV. DIRECAO GERAL", "ATIVIDADE MEIO", "ATIV. FIM NA SEDE", "ATIV. FIM DESTACADO", "QUADRO ESPECIALISTA", "DISP MED DEFINITIVA"]:
                return "AADP Regular"
            return "AADP SF. Restrito"

        df_sigef['NUMERO_CLEAN'] = df_sigef['NUMERO'].apply(_c_get_pm)
        df_sigef = df_sigef[df_sigef['NUMERO_CLEAN'] != ""]
        

        if not df_com.empty:
            df_com['nrPM_Avaliado_CLEAN'] = df_com['nrPM (Avaliado)'].apply(_c_get_pm)
            df_com = df_com[df_com['nrPM_Avaliado_CLEAN'] != ""]
            
            def _count_valid_evals(row):
                c = 0
                if str(row.get('nrPM (Avaliador1)', '')).strip() not in ('', 'nan', 'None'): c += 1
                if str(row.get('nrPM (Avaliador2)', '')).strip() not in ('', 'nan', 'None'): c += 1
                if str(row.get('nrPM (Homologador)', '')).strip() not in ('', 'nan', 'None'): c += 1
                return c
            
            df_com['__qtd_membros'] = df_com.apply(_count_valid_evals, axis=1)
            df_com['__original_order'] = range(len(df_com))
            df_com = df_com.sort_values(by=['nrPM_Avaliado_CLEAN', '__qtd_membros', '__original_order'], ascending=[True, True, True])
            df_com = df_com.drop_duplicates(subset=['nrPM_Avaliado_CLEAN'], keep='last')
            df_com = df_com.drop(columns=['__qtd_membros', '__original_order'])
            df_merge = pd.merge(df_sigef, df_com, left_on='NUMERO_CLEAN', right_on='nrPM_Avaliado_CLEAN', how='left')
        else:
            df_merge = df_sigef.copy()
            for col in ['nrPM (Avaliador1)', 'nrPM (Avaliador2)', 'nrPM (Homologador)']:
                df_merge[col] = ""
            for col in ['Unidade RPM Atual (Avaliado)', 'Unidade Principal Atual (Avaliado)', 'Posto/Graduação (Avaliado)']:
                df_merge[col] = None

        df_merge['Status da Comissão'] = df_merge.apply(_c_classify_com, axis=1)
        df_merge['Tipo AADP'] = df_merge['SIT. FUNCIONAL'].apply(_c_aadp_type)
        df_merge['Categoria AADP'] = df_merge['SIT. FUNCIONAL'].apply(_c_aadp_category)

        df_merge['RPM Final'] = df_merge['Unidade RPM Atual (Avaliado)'].fillna(df_merge['NOME RPM']).str.strip()
        df_merge['RPM Final'] = df_merge['RPM Final'].replace({'NÃO': 'DINT', 'AUDI SET': 'AUD SET'})
        df_merge['Unidade Principal Final'] = df_merge['Unidade Principal Atual (Avaliado)'].fillna(df_merge['NOME UNIDADE PRINCIPAL']).str.strip()
        df_merge['Posto/Graduação Final'] = df_merge['Posto/Graduação (Avaliado)'].fillna(df_merge['POSTO/GRADUACAO']).str.strip()
        df_merge = df_merge[df_merge['SIT. FUNCIONAL'] != 'JUIZ/TJM']
        df_merge['Situação AADP'] = df_merge['Status da Comissão'].apply(lambda x: 'Completa' if str(x).upper().startswith('COMPLETA') else 'Incompleta')
        df_merge['Status CDP'] = df_merge['NUMERO_CLEAN'].apply(lambda x: 'CDP Cadastrado' if x in cdp_pms_set else 'Sem CDP')
        
        # Alerta de Praça em Função de Avaliador 2 e Homologador
        pracas_regex = r"(SOLDADO|CABO|3 SARGENTO|2 SARGENTO|1 SARGENTO|SUBTENENTE)"
        av2_praca = df_merge['Posto/Graduação (Avaliador2)'].str.upper().str.contains(pracas_regex, na=False)
        hom_praca = df_merge['Posto/Graduação (Homologador)'].str.upper().str.contains(pracas_regex, na=False)
        df_merge['Alerta Praça AV2/HOM'] = av2_praca | hom_praca

        # Alerta Hierarquia (AV1 > AV2 > HOM)
        HIERARQUIA = {
            "SOLDADO DE 2 CLASSE": 1, "SOLDADO DE 1 CLASSE": 2, "CABO": 3,
            "3 SARGENTO": 4, "2 SARGENTO": 5, "1 SARGENTO": 6, "SUBTENENTE": 7,
            "CADETE": 8, "ALUNO": 8, "ASPIRANTE A OFICIAL": 9,
            "2 TENENTE": 10, "1 TENENTE": 11, "CAPITAO": 12, "CAPITÃO": 12,
            "MAJOR": 13, "TENENTE CORONEL": 14, "CORONEL": 15
        }
        h1 = df_merge['Posto/Graduação (Avaliador1)'].str.upper().str.strip().map(HIERARQUIA).fillna(-1)
        h2 = df_merge['Posto/Graduação (Avaliador2)'].str.upper().str.strip().map(HIERARQUIA).fillna(-1)
        hh = df_merge['Posto/Graduação (Homologador)'].str.upper().str.strip().map(HIERARQUIA).fillna(-1)

        c1 = (h1 > -1) & (h2 > -1) & (h1 > h2)
        c2 = (h2 > -1) & (hh > -1) & (h2 > hh)
        c3 = (h2 == -1) & (h1 > -1) & (hh > -1) & (h1 > hh)
        df_merge['Alerta Hierarquia'] = c1 | c2 | c3

        df_merge['Alerta Geral'] = 'Sem Alerta'
        df_merge.loc[df_merge['Alerta Hierarquia'], 'Alerta Geral'] = '⚠️ ALERTA HIERARQUIA'
        df_merge.loc[df_merge['Alerta Praça AV2/HOM'], 'Alerta Geral'] = '🚨 ALERTA PRAÇA'

        # Configuração da UI de Filtros e Aplicação Simultânea (Cascata)
        df_f = df_merge[df_merge['Tipo AADP'] == global_aadp].copy()
        
        with st.expander("🔍 Filtros de Comissões", expanded=True):
            with st.form("form_filtros_comissoes"):
                r1c1, r1c2, r1c3 = st.columns(3)
                with r1c1:
                    import re
                    def sort_rpm(rpm_name):
                        m = re.match(r'^(\d+)\s*RPM', str(rpm_name))
                        if m:
                            return (0, int(m.group(1)), str(rpm_name))
                        return (1, 0, str(rpm_name))

                    curr_rpm = st.session_state.get('rpm_key', [])
                    rpm_list = sorted(list(set(df_f['RPM Final'].dropna().unique()) | set(curr_rpm)), key=sort_rpm)
                    if sidebar_active_role in ("P1", "SADM") and active_rpm:
                        rpm_filter = [active_rpm]
                        st.info(f"RPM Fixa: {active_rpm}")
                    else:
                        rpm_filter = st.multiselect("Unidade de Direção (RPM)", rpm_list, key='rpm_key')

                with r1c2:
                    curr_unid = st.session_state.get('unid_key', [])
                    unidades = sorted(list(set(df_f['Unidade Principal Final'].dropna().unique()) | set(curr_unid)))
                    if sidebar_active_role == "SADM" and active_unit:
                        unid_filter = [active_unit]
                        st.info(f"Unidade Fixa: {active_unit}")
                    else:
                        unid_filter = st.multiselect("Unidade Principal", unidades, key='unid_key')

                with r1c3:
                    curr_posto = st.session_state.get('posto_key', [])
                    postos = sorted(list(set(df_f['Posto/Graduação Final'].dropna().unique()) | set(curr_posto)), key=lambda x: (-HIERARQUIA.get(str(x).strip().upper(), 0), str(x)))
                    posto_filter = st.multiselect("Posto/Graduação Avaliado", postos, key='posto_key')

                r2c1, r2c2, r2c3 = st.columns(3)
                sub_ativa_filter = []
                sit_especifica = []

                with r2c1:
                    if global_aadp == "Ativa":
                        sub_ativa_filter = st.multiselect("Categoria AADP", ["AADP Regular", "AADP SF. Restrito"], default=["AADP Regular", "AADP SF. Restrito"], help="AADP REGULAR: Militares cuja situação funcional não possui restrição.\nAADP SF. RESTRITO: Militares cuja situação funcional gera alerta/restrição para avaliação regular.")

                with r2c2:
                    curr_sit = st.session_state.get('sit_aadp_key', [])
                    situacao_aadp_list = sorted(list(set(df_f['Situação AADP'].unique()) | set(curr_sit)))
                    situacao_aadp_filter = st.multiselect("Situação AADP", situacao_aadp_list, key='sit_aadp_key')

                with r2c3:
                    curr_tipo = st.session_state.get('tipo_aadp_key', [])
                    tipo_aadp_list = sorted(list(set(df_f['Status da Comissão'].unique()) | set(curr_tipo)))
                    tipo_aadp_filter = st.multiselect("Tipo AADP", tipo_aadp_list, key='tipo_aadp_key')

                r3c1, r3c2, r3c3 = st.columns(3)
                with r3c1:
                    curr_rest = st.session_state.get('sit_rest_key', [])
                    valid_rests = set(df_f[df_f['Categoria AADP'] == 'AADP SF. Restrito']['SIT. FUNCIONAL'].dropna().unique())
                    sits_restritivas = sorted(list(valid_rests | set(curr_rest)))
                    sit_especifica = st.multiselect("Situação Funcional (Restritiva)", sits_restritivas, key='sit_rest_key')
                    
                with r3c2:
                    OPCOES_ALERTA = ["Todas as Comissões", "Com Qualquer Alerta", "🚨 ALERTA PRAÇA", "⚠️ ALERTA HIERARQUIA"]
                    alerta_filter = st.selectbox("Alertas de Comissão", OPCOES_ALERTA, key='alerta_geral', help="ALERTA PRAÇA: Comissões que possuem praças na função de Avaliador 2 ou Homologador de forma incorreta.\nALERTA HIERARQUIA: Membro da comissão exercendo função hierarquicamente de superior de forma indevida.")

                with r3c3:
                    OPCOES_CDP_COMPL = ["Todas as Comissões", "Completas COM CDP Cadastrado", "Completas SEM CDP Cadastrado"]
                    cdp_compl_filter = st.selectbox(
                        "Status CDP (Comissões Completas)",
                        OPCOES_CDP_COMPL,
                        key='cdp_compl_filtro',
                        help="Filtra especificamente as comissões completas que já possuem CDP cadastrado ou que ainda não possuem CDP cadastrado."
                    )

                submit_filtros = st.form_submit_button("Aplicar Filtros 🚀", use_container_width=True)

        # Aplicando filtros após formulário
        if rpm_filter: df_f = df_f[df_f['RPM Final'].isin(rpm_filter)]
        if unid_filter: df_f = df_f[df_f['Unidade Principal Final'].isin(unid_filter)]
        if posto_filter: df_f = df_f[df_f['Posto/Graduação Final'].isin(posto_filter)]
        
        if sub_ativa_filter:
            df_f = df_f[df_f['Categoria AADP'].isin(sub_ativa_filter)]
        elif global_aadp == "Ativa":
            df_f = df_f.iloc[0:0]
            
        if situacao_aadp_filter: df_f = df_f[df_f['Situação AADP'].isin(situacao_aadp_filter)]
        if tipo_aadp_filter: df_f = df_f[df_f['Status da Comissão'].isin(tipo_aadp_filter)]
        
        if "AADP SF. Restrito" in sub_ativa_filter and sit_especifica:
            mask_restritiva = (df_f['Categoria AADP'] == 'AADP SF. Restrito') & (df_f['SIT. FUNCIONAL'].isin(sit_especifica))
            mask_not_restritiva = (df_f['Categoria AADP'] != 'AADP SF. Restrito')
            df_f = df_f[mask_restritiva | mask_not_restritiva]

        if alerta_filter == "Com Qualquer Alerta":
            df_f = df_f[df_f['Alerta Geral'] != 'Sem Alerta']
        elif alerta_filter in ["🚨 ALERTA PRAÇA", "⚠️ ALERTA HIERARQUIA"]:
            df_f = df_f[df_f['Alerta Geral'] == alerta_filter]

        if cdp_compl_filter == "Completas COM CDP Cadastrado":
            df_f = df_f[df_f['Status da Comissão'].str.upper().str.startswith("COMPLETA") & (df_f['Status CDP'] == 'CDP Cadastrado')]
        elif cdp_compl_filter == "Completas SEM CDP Cadastrado":
            df_f = df_f[df_f['Status da Comissão'].str.upper().str.startswith("COMPLETA") & (df_f['Status CDP'] == 'Sem CDP')]

        # KPIs
        total_mil = len(df_f)
        mask_compl_kpi = df_f['Status da Comissão'].str.upper().str.startswith("COMPLETA")
        com_compl = len(df_f[mask_compl_kpi])
        com_compl_com_cdp = len(df_f[mask_compl_kpi & (df_f['Status CDP'] == 'CDP Cadastrado')])
        com_compl_sem_cdp = len(df_f[mask_compl_kpi & (df_f['Status CDP'] == 'Sem CDP')])
        com_pend = total_mil - com_compl

        pct_com_cdp = (com_compl_com_cdp / com_compl * 100) if com_compl > 0 else 0
        pct_sem_cdp = (com_compl_sem_cdp / com_compl * 100) if com_compl > 0 else 0

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Total Filtrados", f"{total_mil:,}".replace(',', '.'))
        k2.metric("Comissões Completas", f"{com_compl:,}".replace(',', '.'))
        k3.metric("🟢 Completas COM CDP", f"{com_compl_com_cdp:,}".replace(',', '.'), f"{pct_com_cdp:.1f}% das completas")
        k4.metric("🔴 Completas SEM CDP", f"{com_compl_sem_cdp:,}".replace(',', '.'), f"{pct_sem_cdp:.1f}% das completas")
        k5.metric("⚠️ Incompletas/Ausentes", f"{com_pend:,}".replace(',', '.'))

        st.markdown("---")

        # Tabela Detalhada
        st.markdown("---")
        st.markdown("### 📋 Tabela Detalhada de Comissões")
        cols_disp = [
            'NUMERO', 'Posto/Graduação Final', 'NOME SERVIDOR', 'RPM Final', 'Unidade Principal Final', 
            'SIT. FUNCIONAL', 'Tipo AADP', 'Categoria AADP', 'Status da Comissão', 'Status CDP', 'Alerta Geral',
            'nrPM (Avaliador1)', 'Posto/Graduação (Avaliador1)', 'Nome Completo (Avaliador1)',
            'nrPM (Avaliador2)', 'Posto/Graduação (Avaliador2)', 'Nome Completo (Avaliador2)',
            'nrPM (Homologador)', 'Posto/Graduação (Homologador)', 'Nome Completo (Homologador)'
        ]
        cols_presentes = [c for c in cols_disp if c in df_f.columns]
        df_export = df_f[cols_presentes].copy()
        
        csv_data = df_export.to_csv(index=False, sep=';').encode('utf-8-sig')
        st.download_button(
        "Baixar Auditoria de Comissoes (CSV Rapido)",
        csv_data,
        f"Auditoria_Comissoes_{now_br().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
            type="primary"
        )
        
        def highlight_alert(row):
            alerta = row.get('Alerta Geral', '')
            if alerta == '🚨 ALERTA PRAÇA':
                return ['background-color: rgba(255, 0, 0, 0.3)'] * len(row)
            elif alerta == '⚠️ ALERTA HIERARQUIA':
                return ['background-color: rgba(255, 255, 0, 0.3); color: black'] * len(row)
            return [''] * len(row)

        if not df_export.empty:
            if len(df_export) > 1500:
                st.warning('Aviso: Exibindo apenas as primeiras 1.500 linhas na tela. Baixe o CSV acima para visualizar todos.')
            styled_df = df_export.head(1500).style.apply(highlight_alert, axis=1)
            st.dataframe(styled_df, use_container_width=True, hide_index=True)
        else:
            st.dataframe(df_export, use_container_width=True, hide_index=True)

        # Novo Ranking de Membros
        st.markdown("---")
        st.markdown("### 🏆 Ranking de Membros de Comissão")
        
        # Preparar dados dos membros
        df_av1 = df_f[['nrPM (Avaliador1)', 'Nome Completo (Avaliador1)', 'Posto/Graduação (Avaliador1)']].copy()
        df_av1.columns = ['nrPM', 'Nome Completo', 'Posto/Graduação']
        df_av1['Função'] = 'Avaliador 1'
        
        df_av2 = df_f[['nrPM (Avaliador2)', 'Nome Completo (Avaliador2)', 'Posto/Graduação (Avaliador2)']].copy()
        df_av2.columns = ['nrPM', 'Nome Completo', 'Posto/Graduação']
        df_av2['Função'] = 'Avaliador 2'
        
        df_hom = df_f[['nrPM (Homologador)', 'Nome Completo (Homologador)', 'Posto/Graduação (Homologador)']].copy()
        df_hom.columns = ['nrPM', 'Nome Completo', 'Posto/Graduação']
        df_hom['Função'] = 'Homologador'
        
        df_members = pd.concat([df_av1, df_av2, df_hom], ignore_index=True)
        # Limpar vazios
        df_members = df_members.dropna(subset=['nrPM'])
        df_members = df_members[df_members['nrPM'].str.strip() != '']
        df_members = df_members[df_members['nrPM'].str.strip() != '-']
        
        df_members['nrPM_CLEAN'] = df_members['nrPM'].apply(_c_get_pm)
        if 'NUMERO_CLEAN' in df_sigef.columns and 'NOME RPM' in df_sigef.columns and 'NOME UNIDADE PRINCIPAL' in df_sigef.columns:
            df_sigef_u = df_sigef[['NUMERO_CLEAN', 'NOME RPM', 'NOME UNIDADE PRINCIPAL']].drop_duplicates(subset=['NUMERO_CLEAN'])
            df_members = pd.merge(df_members, df_sigef_u, left_on='nrPM_CLEAN', right_on='NUMERO_CLEAN', how='left')
            df_members.rename(columns={'NOME RPM': 'RPM', 'NOME UNIDADE PRINCIPAL': 'Unidade Principal'}, inplace=True)
        else:
            df_members['RPM'] = '-'
            df_members['Unidade Principal'] = '-'
            
        df_members['RPM'] = df_members['RPM'].fillna('-')
        df_members['Unidade Principal'] = df_members['Unidade Principal'].fillna('-')
        
        r1c1, r1c2 = st.columns(2)
        with r1c1:
            postos_membros = sorted(list(set(df_members['Posto/Graduação'].dropna().unique())), key=lambda x: (-HIERARQUIA.get(str(x).strip().upper(), 0), str(x)))
            filtro_posto_membro = st.multiselect("Posto/Graduação do Membro", postos_membros, key='filtro_posto_ranking')

        with r1c2:
            filtro_funcao = st.multiselect("Função Exercida", ['Avaliador 1', 'Avaliador 2', 'Homologador'], key='filtro_funcao_ranking')

        if filtro_posto_membro:
            df_members = df_members[df_members['Posto/Graduação'].isin(filtro_posto_membro)]

        if filtro_funcao:
            df_members = df_members[df_members['Função'].isin(filtro_funcao)]
            
        if not df_members.empty:
            counts = df_members.groupby(['nrPM', 'Posto/Graduação', 'Nome Completo', 'RPM', 'Unidade Principal', 'Função']).size().unstack(fill_value=0).reset_index()
            
            for c in ['Avaliador 1', 'Avaliador 2', 'Homologador']:
                if c not in counts.columns:
                    counts[c] = 0
                    
            counts['Total'] = counts['Avaliador 1'] + counts['Avaliador 2'] + counts['Homologador']
            counts = counts[['nrPM', 'Posto/Graduação', 'Nome Completo', 'RPM', 'Unidade Principal', 'Avaliador 1', 'Avaliador 2', 'Homologador', 'Total']]
            counts = counts.sort_values('Total', ascending=False)
            
            st.dataframe(counts, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum membro encontrado com os filtros atuais.")
            
        st.markdown("---")
        st.markdown("### 📊 Gráficos de Distribuição")
        
        c_graf1, c_graf2 = st.columns(2)
        
        with c_graf1:
            st.markdown("**Distribuição por Status da Comissão**")
            import plotly.express as px
            rotacao_grafico = st.slider("Ajuste de Rotação (°)", min_value=0, max_value=360, value=115, step=5, key="rot_pizza")
            sc_df = df_f['Status da Comissão'].value_counts().reset_index()
            sc_df.columns = ['Status', 'Quantidade']
            if not sc_df.empty:
                fig_pizza = px.pie(sc_df, values='Quantidade', names='Status', hole=0.3,
                                   color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_pizza.update_traces(textposition='outside', textinfo='label+value+percent', rotation=rotacao_grafico)
                fig_pizza.update_layout(showlegend=False, margin=dict(r=350, l=200, t=150, b=100), height=650)
                st.plotly_chart(fig_pizza, use_container_width=True)
            else:
                st.info("Sem dados.")
                
        with c_graf2:
            st.markdown("**Resumo por Situação Funcional**")
            sf_df = df_f['SIT. FUNCIONAL'].value_counts().reset_index()
            sf_df.columns = ['Situação Funcional', 'Quantidade']
            if not sf_df.empty:
                fig_bar = px.bar(sf_df, x='Situação Funcional', y='Quantidade', text='Quantidade',
                                 color='Situação Funcional', color_discrete_sequence=px.colors.qualitative.Set2)
                fig_bar.update_traces(textposition='outside')
                if len(sf_df) <= 6:
                    fig_bar.update_traces(width=0.05 * len(sf_df))
                fig_bar.update_layout(showlegend=False, xaxis_title="", yaxis_title="Quantidade")
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("Sem dados.")




# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO: CONTROLE DO CDP (Metas e Acompanhamentos)
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner="⏳ Carregando dados do Controle do CDP...")
def load_controle_cdp_data(_db_path: str = "", _drive_metas_id: str = "", _drive_geral_id: str = "", _drive_si_id: str = "", _ano="2026"):
    """Processa a planilha de Metas e acompanhamentos do ano cruzando com geral.csv e SIGEF.csv."""
    import tempfile
    from datetime import datetime, date
    import math

    cache_dir = os.path.join(tempfile.gettempdir(), f"aadp_drive_cache_{_ano}")
    os.makedirs(cache_dir, exist_ok=True)
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # Localizar Metas e acompanhamentos 2026 Completo.csv
    possible_metas = [
        os.path.join(base_dir, f"DADOS AADP {_ano}", f"Metas e acompanhamentos {_ano} Completo.csv"),
        os.path.join(base_dir, f"DADOS AADP {_ano}", "Metas e acompanhamentos 2026 Completo.csv"),
        os.path.join(cache_dir, f"Metas e acompanhamentos {_ano} Completo.csv"),
        os.path.join(cache_dir, "Metas e acompanhamentos 2026 Completo.csv"),
        os.path.join(_db_path or "", f"Metas e acompanhamentos {_ano} Completo.csv"),
        os.path.join(_db_path or "", "Metas e acompanhamentos 2026 Completo.csv"),
        os.path.join(base_dir, "dados", _ano, f"Metas e acompanhamentos {_ano} Completo.csv"),
        os.path.join(base_dir, "dados", f"Metas e acompanhamentos {_ano} Completo.csv"),
        os.path.join(base_dir, f"Metas e acompanhamentos {_ano} Completo.csv"),
        os.path.join(base_dir, "dados", "Metas e acompanhamentos 2026 Completo.csv"),
        os.path.join(base_dir, "Metas e acompanhamentos 2026 Completo.csv")
    ]
    metas_path = next((p for p in possible_metas if os.path.exists(p) and os.path.getsize(p) > 0), None)

    if not metas_path and _drive_metas_id:
        metas_dest = os.path.join(cache_dir, f"Metas e acompanhamentos {_ano} Completo.csv")
        try:
            _baixar_drive(_drive_metas_id, metas_dest)
            if os.path.exists(metas_dest) and os.path.getsize(metas_dest) > 0:
                metas_path = metas_dest
        except Exception:
            pass

    if not metas_path:
        return pd.DataFrame(), pd.DataFrame(), f"Planilha 'Metas e acompanhamentos {_ano} Completo.csv' não encontrada."

    # Localizar SIGEF.csv
    possible_sigef = [
        os.path.join(base_dir, f"DADOS AADP {_ano}", "SIGEF.csv"),
        os.path.join(cache_dir, "SIGEF.csv"),
        os.path.join(_db_path or "", "SIGEF.csv"),
        os.path.join(base_dir, "dados", "SIGEF.csv"),
        os.path.join(base_dir, "SIGEF.csv")
    ]
    sigef_path = next((p for p in possible_sigef if os.path.exists(p) and os.path.getsize(p) > 0), None)
    if not sigef_path and _drive_si_id:
        sigef_dest = os.path.join(cache_dir, "SIGEF.csv")
        try:
            _baixar_drive(_drive_si_id, sigef_dest)
            if os.path.exists(sigef_dest) and os.path.getsize(sigef_dest) > 0:
                sigef_path = sigef_dest
        except Exception:
            pass

    # Localizar geral.csv
    possible_geral = [
        os.path.join(base_dir, f"DADOS AADP {_ano}", "geral.csv"),
        os.path.join(cache_dir, "geral.csv"),
        os.path.join(_db_path or "", "geral.csv"),
        os.path.join(base_dir, "dados", "geral.csv"),
        os.path.join(base_dir, "geral.csv")
    ]
    geral_path = next((p for p in possible_geral if os.path.exists(p) and os.path.getsize(p) > 0), None)
    if not geral_path and _drive_geral_id:
        geral_dest = os.path.join(cache_dir, "geral.csv")
        try:
            _baixar_drive(_drive_geral_id, geral_dest)
            if os.path.exists(geral_dest) and os.path.getsize(geral_dest) > 0:
                geral_path = geral_dest
        except Exception:
            pass

    def _parse_d(d_str):
        if not d_str or str(d_str).strip() in ("", "-", "nan", "none", "None", "<NA>"):
            return None
        d_str = str(d_str).strip()
        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(d_str, fmt).date()
            except Exception:
                pass
        return None

    def _norm_pm(val):
        if not val or str(val).strip() in ("", "-", "nan", "none", "None", "<NA>"):
            return ""
        try:
            return str(int(float(str(val).strip())))
        except Exception:
            s = str(val).strip().lstrip("0")
            return s if s else "0"

    # 1. Carregar lotação do SIGEF
    sigef_map = {}
    if sigef_path:
        with open(sigef_path, encoding="cp1252", errors="replace") as f_si:
            for row in csv.reader(f_si, delimiter=";"):
                if len(row) > 9:
                    sigef_map[_norm_pm(row[0])] = row[9].strip()

    # 2. Carregar Data do CDP do geral.csv
    cdp_map = {}
    pm_max_cdp = {}
    if geral_path:
        with open(geral_path, "r", encoding="cp1252", errors="replace") as f_ge:
            r_ge = csv.reader(f_ge, delimiter=";")
            header_ge = next(r_ge, [])
            for row in r_ge:
                if len(row) > 37:
                    pm_g = _norm_pm(row[1])
                    av1_g = _norm_pm(row[28])
                    av2_g = _norm_pm(row[37])
                    dt_g = _parse_d(row[5])
                    if dt_g:
                        cdp_map[(pm_g, av1_g, av2_g)] = dt_g
                        cdp_map[(pm_g, av1_g)] = dt_g
                        cdp_map[pm_g] = dt_g
                        if pm_g not in pm_max_cdp or dt_g > pm_max_cdp[pm_g]:
                            pm_max_cdp[pm_g] = dt_g

    meta_offsets = [
        (1, 24, 20), (2, 90, 20), (3, 156, 20), (4, 222, 20),
        (5, 288, 20), (6, 354, 20), (7, 420, 20), (8, 486, 20),
        (9, 552, 17), (10, 609, 9), (11, 642, 9), (12, 675, 9)
    ]

    # Agrupar linhas brutas de metas por Instância: (pm_aval, pm_av1, pm_av2)
    # Se o avaliado tiver o mesmo AV1 e AV2 em vários registros de metas, trata-se de UMA ÚNICA INSTÂNCIA.
    raw_instances = {}
    with open(metas_path, "r", encoding="cp1252", errors="replace") as f_m:
        r_m = csv.reader(f_m, delimiter=";")
        next(r_m, [])
        for row in r_m:
            if not row or len(row) < 24:
                continue
            pm_aval = _norm_pm(row[0])
            pm_av1 = _norm_pm(row[8])
            pm_av2 = _norm_pm(row[16])
            inst_key = (pm_aval, pm_av1, pm_av2)
            if inst_key not in raw_instances:
                raw_instances[inst_key] = {
                    "row_info": row[:24],
                    "rows": []
                }
            raw_instances[inst_key]["rows"].append(row)

    data_atual = now_br().date()
    pms_instances = {}
    metas_detail_rows = []

    for (pm_aval, pm_av1, pm_av2), inst_data in raw_instances.items():
        row0 = inst_data["row_info"]
        nome_aval = row0[1].strip()
        posto_aval = row0[2].strip()
        rpm_aval = row0[3].strip()
        unid_aval = row0[4].strip()
        local_aval = row0[5].strip()
        quadro_aval = row0[6].strip()
        sit_aval = row0[7].strip()

        nome_av1 = row0[9].strip()
        posto_av1 = row0[10].strip()
        rpm_av1 = row0[11].strip()
        unid_av1 = row0[12].strip()
        local_av1 = row0[13].strip()

        nome_av2 = row0[17].strip()
        posto_av2 = row0[18].strip()
        rpm_av2 = row0[19].strip()
        unid_av2 = row0[20].strip()
        local_av2 = row0[21].strip()

        seen_metas = set()
        metas_dates = []
        first_meta_date = None
        total_acomps_av1 = 0
        total_outros_acomps = 0
        all_inst_intervals = []
        all_inst_dates = []
        inst_metas_indices = []

        dt_cdp = cdp_map.get((pm_aval, pm_av1, pm_av2)) or cdp_map.get((pm_aval, pm_av1)) or cdp_map.get(pm_aval)
        if dt_cdp:
            all_inst_dates.append(dt_cdp)

        for r in inst_data["rows"]:
            for m_num, start_idx, max_acomps in meta_offsets:
                if start_idx + 5 >= len(r):
                    continue
                d_meta = _parse_d(r[start_idx + 2])
                if not d_meta:
                    continue
                desc_meta = r[start_idx + 1].strip()
                meta_key = (m_num, desc_meta, d_meta)
                if meta_key in seen_metas:
                    continue
                seen_metas.add(meta_key)

                if first_meta_date is None or d_meta < first_meta_date:
                    first_meta_date = d_meta

                metas_dates.append(f"M{m_num}: {d_meta.strftime('%d/%m/%Y')}")
                all_inst_dates.append(d_meta)

                prazo_ini = _parse_d(r[start_idx + 3].strip())
                prazo_fim = _parse_d(r[start_idx + 4].strip())

                acomps_av1 = []
                outros_acomps = 0
                for k in range(1, max_acomps + 1):
                    idx_autor = start_idx + 6 + 3 * (k - 1)
                    idx_desc = idx_autor + 1
                    idx_data = idx_autor + 2
                    if idx_data >= len(r):
                        break
                    autor_k = _norm_pm(r[idx_autor])
                    data_k = _parse_d(r[idx_data])
                    if not data_k:
                        continue
                    if autor_k == pm_av1:
                        acomps_av1.append((data_k, r[idx_desc].strip()))
                        all_inst_dates.append(data_k)
                    else:
                        outros_acomps += 1

                total_acomps_av1 += len(acomps_av1)
                total_outros_acomps += outros_acomps

                acomps_av1.sort(key=lambda x: x[0])
                meta_intervals = []
                prev_d = d_meta
                acomps_eventos = []
                for num_ac_idx, (d_acomp, desc_ac) in enumerate(acomps_av1, start=1):
                    diff = (d_acomp - prev_d).days
                    meta_intervals.append(diff)
                    all_inst_intervals.append(diff)
                    acomps_eventos.append({
                        "num": num_ac_idx,
                        "data_str": d_acomp.strftime("%d/%m/%Y"),
                        "data_iso": d_acomp.strftime("%Y-%m-%d"),
                        "dias": diff,
                        "desc": desc_ac
                    })
                    prev_d = d_acomp

                avg_meta_interval = round(sum(meta_intervals) / len(meta_intervals), 1) if meta_intervals else None
                interv_strs = [f"{diff}d" for diff in meta_intervals]
                dias_meta_acomp1 = meta_intervals[0] if len(meta_intervals) > 0 else None

                m_idx = len(metas_detail_rows)
                inst_metas_indices.append(m_idx)
                metas_detail_rows.append({
                    "nrPM (Avaliado)": pm_aval,
                    "Nome (Avaliado)": nome_aval,
                    "Posto/Grad. (Avaliado)": posto_aval,
                    "Unidade RPM": rpm_aval,
                    "Unidade Principal": unid_aval,
                    "Situação Comissão": "Comissão Atual",  # será atualizado após ordenação
                    "nrPM (AV1)": pm_av1,
                    "Nome (AV1)": nome_av1,
                    "Meta": f"Meta {m_num}",
                    "Descrição da Meta": desc_meta,
                    "Data Cadastro Meta": d_meta.strftime("%d/%m/%Y"),
                    "Prazo Inicial": prazo_ini.strftime("%d/%m/%Y") if prazo_ini else "-",
                    "Prazo Final": prazo_fim.strftime("%d/%m/%Y") if prazo_fim else "-",
                    "Qtd Acomps AV1": len(acomps_av1),
                    "Qtd Outros Acomps": outros_acomps,
                    "Dias (Meta -> 1º Acomp)": dias_meta_acomp1 if dias_meta_acomp1 is not None else "-",
                    "Intervalos Acomps": " -> ".join(interv_strs) if interv_strs else "-",
                    "Média Dias Lançamentos": avg_meta_interval if avg_meta_interval is not None else "-",
                    "Data Último Acomp": acomps_av1[-1][0].strftime("%d/%m/%Y") if acomps_av1 else "-",
                    "Data Cadastro Meta ISO": d_meta.strftime("%Y-%m-%d"),
                    "Acomps Eventos": acomps_eventos
                })

        if dt_cdp:
            status_cdp = "CDP Cadastrado"
            dt_cdp_str = dt_cdp.strftime("%d/%m/%Y")
        elif seen_metas and first_meta_date:
            status_cdp = "CDP Cadastrado"
            dt_cdp = first_meta_date
            dt_cdp_str = first_meta_date.strftime("%d/%m/%Y")
            if dt_cdp not in all_inst_dates:
                all_inst_dates.insert(0, dt_cdp)
        else:
            status_cdp = "CDP NÃO Cadastrado"
            dt_cdp = None
            dt_cdp_str = "-"

        avg_inst_interval = round(sum(all_inst_intervals) / len(all_inst_intervals), 1) if all_inst_intervals else None
        if all_inst_dates:
            last_dt = max(all_inst_dates)
            dias_desde_ultimo = max(0, (data_atual - last_dt).days)
            last_dt_str = last_dt.strftime("%d/%m/%Y")
        else:
            last_dt_str = "-"
            dias_desde_ultimo = "-"

        ref_date = dt_cdp or (max(all_inst_dates) if all_inst_dates else date.min)

        inst_row = {
            "nrPM (Avaliado)": pm_aval,
            "Nome (Avaliado)": nome_aval,
            "Posto/Grad. (Avaliado)": posto_aval,
            "Unidade RPM (Avaliado)": rpm_aval,
            "Unidade Principal (Avaliado)": unid_aval,
            "Situação Comissão": "Comissão Atual",  # será atualizado
            "Status CDP": status_cdp,
            "Data do CDP": dt_cdp_str,
            "Qtd Metas": len(seen_metas),
            "Datas das Metas": "; ".join(metas_dates) if metas_dates else "-",
            "Qtd Acomps AV1": total_acomps_av1,
            "Qtd Outros Acomps": total_outros_acomps,
            "Média Dias entre Lançamentos": avg_inst_interval if avg_inst_interval is not None else "-",
            "Último Lançamento": last_dt_str,
            "Dias desde Último Lançamento": dias_desde_ultimo,
            "nrPM (AV1)": pm_av1,
            "Nome (AV1)": nome_av1,
            "Posto (AV1)": posto_av1,
            "Unidade Principal (AV1)": unid_av1,
            "nrPM (AV2)": pm_av2,
            "Nome (AV2)": nome_av2,
            "Posto (AV2)": posto_av2,
            "Unidade Principal (AV2)": unid_av2,
        }

        if pm_aval not in pms_instances:
            pms_instances[pm_aval] = []
        pms_instances[pm_aval].append({
            "inst_key": (pm_aval, pm_av1, pm_av2),
            "ref_date": ref_date,
            "inst_dict": inst_row,
            "metas_indices": inst_metas_indices
        })

    # Atribuição da Situação da Comissão:
    # Para o mesmo avaliado, a instância mais recente é a "Comissão Atual", e as anteriores são "Nota Provisória"
    instances_rows = []
    for pm_aval, inst_list in pms_instances.items():
        if len(inst_list) == 1:
            sc = "Comissão Atual"
            inst_list[0]["inst_dict"]["Situação Comissão"] = sc
            for m_idx in inst_list[0]["metas_indices"]:
                metas_detail_rows[m_idx]["Situação Comissão"] = sc
            instances_rows.append(inst_list[0]["inst_dict"])
        else:
            inst_list.sort(key=lambda x: x["ref_date"], reverse=True)
            inst_list[0]["inst_dict"]["Situação Comissão"] = "Comissão Atual"
            for m_idx in inst_list[0]["metas_indices"]:
                metas_detail_rows[m_idx]["Situação Comissão"] = "Comissão Atual"
            instances_rows.append(inst_list[0]["inst_dict"])

            for older_item in inst_list[1:]:
                older_item["inst_dict"]["Situação Comissão"] = "Nota Provisória"
                for m_idx in older_item["metas_indices"]:
                    metas_detail_rows[m_idx]["Situação Comissão"] = "Nota Provisória"
                instances_rows.append(older_item["inst_dict"])

    df_inst = pd.DataFrame(instances_rows)
    df_metas = pd.DataFrame(metas_detail_rows)
    return df_inst, df_metas, None


if active_page == "Controle do CDP" and sidebar_active_role.upper() in ("ADMINISTRADOR", "GESTOR", "P1", "SADM"):
    st.markdown("### 🎯 Controle do CDP — Acompanhamento de Metas e Prazos")
    st.caption("Auditoria passo a passo das comissões: registro do CDP, pactuação de metas e acompanhamentos tempestivos realizados pelo Avaliador 1 (AV1).")

    _role_cdp = sidebar_active_role.upper()
    active_rpm = st.session_state.get("simulated_rpm", st.session_state.get("user_rpm", "")) if st.session_state.get("simulation_active", False) else st.session_state.get("user_rpm", "")
    active_unit = st.session_state.get("simulated_unit", st.session_state.get("user_unit", "")) if st.session_state.get("simulation_active", False) else st.session_state.get("user_unit", "")

    cfg_cdp = load_config()
    _active_y = str(st.session_state.get("selected_year", "2026"))
    _y_cfg = get_active_year_config(_active_y, cfg_cdp)
    db_p = _y_cfg["db_path"]
    d_metas_id = _y_cfg["drive_metas_id"]
    d_geral_id = _y_cfg["drive_geral_id"]
    d_si_id = _y_cfg["drive_si_id"]

    df_inst, df_metas, err_cdp = load_controle_cdp_data(db_p, d_metas_id, d_geral_id, d_si_id, _ano=_active_y)

    if err_cdp:
        st.error(f"❌ {err_cdp}")
    elif df_inst.empty:
        st.warning("Nenhum dado encontrado para análise de Controle do CDP.")
    else:
        # ── Restrição de Escopo por Perfil ─────────────────────────────────────
        # ADMINISTRADOR / GESTOR → Acesso integral a todas as unidades da PMMG
        # P1                     → Somente avaliados pertencentes à sua UDI/UDG (RPM)
        # SADM                   → Somente avaliados pertencentes à sua Unidade Principal
        if _role_cdp == "P1":
            if active_rpm and "Unidade RPM (Avaliado)" in df_inst.columns:
                df_inst = df_inst[df_inst["Unidade RPM (Avaliado)"].astype(str).str.upper() == str(active_rpm).upper()]
                st.info(f"🔒 Exibindo apenas avaliados da sua UDI/UDG: **{active_rpm}**")
            else:
                st.warning("⚠️ RPM do usuário não identificado. Contate o administrador.")
                st.stop()
        elif _role_cdp == "SADM":
            if active_unit and "Unidade Principal (Avaliado)" in df_inst.columns:
                df_inst = df_inst[df_inst["Unidade Principal (Avaliado)"].astype(str).str.upper() == str(active_unit).upper()]
                st.info(f"🔒 Exibindo apenas avaliados da sua Unidade Principal: **{active_unit}**")
            else:
                st.warning("⚠️ Unidade Principal do usuário não identificada. Contate o administrador.")
                st.stop()

        if df_inst.empty:
            st.warning("Nenhum militar avaliado encontrado para a sua unidade.")
            st.stop()

        # Garantir que df_metas contenha apenas metas dos avaliados permitidos da unidade
        if not df_metas.empty:
            allowed_pms = set(df_inst["nrPM (Avaliado)"])
            df_metas = df_metas[df_metas["nrPM (Avaliado)"].isin(allowed_pms)].copy()

        # ── FILTROS ─────────────────────────────────────────────────────────────
        with st.expander("🔍 Filtros de Consulta", expanded=True):
            f_col1, f_col2, f_col3, f_col4 = st.columns([1.5, 1.5, 1, 1])
            with f_col1:
                busca_aval = st.text_input("Buscar por Nº PM ou Nome (Avaliado):", "", key="cdp_busca_aval").strip().lower()
            with f_col2:
                busca_av1 = st.text_input("Buscar por Nº PM ou Nome (AV1):", "", key="cdp_busca_av1").strip().lower()
            with f_col3:
                rpms_opts = ["Todas"] + sorted([str(x) for x in df_inst["Unidade RPM (Avaliado)"].unique() if x and str(x) != "-"])
                if _role_cdp in ("P1", "SADM"):
                    sel_rpm = active_rpm if active_rpm in rpms_opts else (rpms_opts[1] if len(rpms_opts) > 1 else "Todas")
                    st.selectbox("Unidade RPM:", [sel_rpm], key="cdp_rpm", disabled=True)
                else:
                    sel_rpm = st.selectbox("Unidade RPM:", rpms_opts, key="cdp_rpm")
            with f_col4:
                sc_opts = ["Todas", "Comissão Atual", "Nota Provisória"]
                sel_sc = st.selectbox("Situação da Comissão:", sc_opts, key="cdp_sc")

            f_col5, f_col6, f_col7, f_col8 = st.columns(4)
            with f_col5:
                unids_opts = ["Todas"] + sorted([str(x) for x in df_inst["Unidade Principal (Avaliado)"].unique() if x and str(x) != "-"])
                if _role_cdp == "SADM":
                    sel_unid = active_unit if active_unit in unids_opts else (unids_opts[1] if len(unids_opts) > 1 else "Todas")
                    st.selectbox("Unidade Principal:", [sel_unid], key="cdp_unid", disabled=True)
                else:
                    sel_unid = st.selectbox("Unidade Principal:", unids_opts, key="cdp_unid")
            with f_col6:
                status_cdp_opts = ["Todos", "CDP Cadastrado", "CDP NÃO Cadastrado"]
                sel_cdp_status = st.selectbox("Status do CDP:", status_cdp_opts, key="cdp_st")
            with f_col7:
                metas_opts = ["Todas", "Sem Metas (0)", "1 Meta", "2 Metas", "3 ou mais Metas"]
                sel_metas = st.selectbox("Qtd de Metas:", metas_opts, key="cdp_metas")
            with f_col8:
                acomp_opts = ["Todos", "Com Acompanhamento AV1", "Sem Acompanhamento AV1"]
                sel_acomp = st.selectbox("Acompanhamentos AV1:", acomp_opts, key="cdp_acomp")

        # Aplicar filtros no DataFrame de Instâncias
        df_filtered_inst = df_inst.copy()

        if busca_aval:
            mask_aval = (
                df_filtered_inst["nrPM (Avaliado)"].str.lower().str.contains(busca_aval, na=False) |
                df_filtered_inst["Nome (Avaliado)"].str.lower().str.contains(busca_aval, na=False)
            )
            df_filtered_inst = df_filtered_inst[mask_aval]

        if busca_av1:
            mask_av1 = (
                df_filtered_inst["nrPM (AV1)"].str.lower().str.contains(busca_av1, na=False) |
                df_filtered_inst["Nome (AV1)"].str.lower().str.contains(busca_av1, na=False)
            )
            df_filtered_inst = df_filtered_inst[mask_av1]

        if sel_rpm != "Todas":
            df_filtered_inst = df_filtered_inst[df_filtered_inst["Unidade RPM (Avaliado)"] == sel_rpm]

        if sel_unid != "Todas":
            df_filtered_inst = df_filtered_inst[df_filtered_inst["Unidade Principal (Avaliado)"] == sel_unid]

        if sel_sc != "Todas":
            df_filtered_inst = df_filtered_inst[df_filtered_inst["Situação Comissão"] == sel_sc]

        if sel_cdp_status != "Todos":
            df_filtered_inst = df_filtered_inst[df_filtered_inst["Status CDP"] == sel_cdp_status]

        if sel_metas == "Sem Metas (0)":
            df_filtered_inst = df_filtered_inst[df_filtered_inst["Qtd Metas"] == 0]
        elif sel_metas == "1 Meta":
            df_filtered_inst = df_filtered_inst[df_filtered_inst["Qtd Metas"] == 1]
        elif sel_metas == "2 Metas":
            df_filtered_inst = df_filtered_inst[df_filtered_inst["Qtd Metas"] == 2]
        elif sel_metas == "3 ou mais Metas":
            df_filtered_inst = df_filtered_inst[df_filtered_inst["Qtd Metas"] >= 3]

        if sel_acomp == "Com Acompanhamento AV1":
            df_filtered_inst = df_filtered_inst[df_filtered_inst["Qtd Acomps AV1"] > 0]
        elif sel_acomp == "Sem Acompanhamento AV1":
            df_filtered_inst = df_filtered_inst[df_filtered_inst["Qtd Acomps AV1"] == 0]

        # Filtrar o DataFrame de Metas de acordo com os PMs filtrados
        valid_pms = set(df_filtered_inst["nrPM (Avaliado)"])
        df_filtered_metas = df_metas[df_metas["nrPM (Avaliado)"].isin(valid_pms)].copy() if not df_metas.empty else pd.DataFrame()

        # ── CARDS DE RESUMO EXECUTIVO ──────────────────────────────────────────
        total_inst = len(df_filtered_inst)
        cadastrados = (df_filtered_inst["Status CDP"] == "CDP Cadastrado").sum()
        pct_cdp = (cadastrados / total_inst * 100) if total_inst > 0 else 0
        total_metas_count = df_filtered_inst["Qtd Metas"].sum()
        total_acomps_av1_count = df_filtered_inst["Qtd Acomps AV1"].sum()

        valid_intervals = [float(x) for x in df_filtered_inst["Média Dias entre Lançamentos"] if x != "-" and str(x).replace(".", "").isdigit()]
        media_geral_dias = round(sum(valid_intervals) / len(valid_intervals), 1) if valid_intervals else "-"

        c_m1, c_m2, c_m3, c_m4, c_m5, c_m6 = st.columns(6)
        c_m1.metric("📋 Instâncias", fmt_num(total_inst))
        c_m2.metric("🟢 CDP Cadastrado", f"{pct_cdp:.1f}%", f"{fmt_num(cadastrados)} cadastrados")
        c_m3.metric("🎯 Metas Pactuadas", fmt_num(total_metas_count))
        c_m4.metric("👥 Acomps AV1", fmt_num(total_acomps_av1_count))
        c_m5.metric("⏱️ Média Intervalo", f"{media_geral_dias} dias" if media_geral_dias != "-" else "-")
        c_m6.metric("⚠️ Sem Acomp. AV1", fmt_num((df_filtered_inst["Qtd Acomps AV1"] == 0).sum()))

        st.markdown("---")

        # ── ABAS DE NAVEGAÇÃO ──────────────────────────────────────────────────
        tab_inst, tab_metas_det, tab_timeline = st.tabs([
            "📋 Visão por Instância (Avaliado / Comissão)",
            "🎯 Detalhamento por Meta e Prazos",
            "🔍 Linha do Tempo e Consulta Individual"
        ])

        with tab_inst:
            st.markdown(f"#### Relação de Instâncias de Avaliação ({fmt_num(len(df_filtered_inst))})")

            # Botão de Exportação Excel (baixa a base completa filtrada)
            col_exp1, col_exp2 = st.columns([1.5, 3.5])
            with col_exp1:
                import io
                buf_inst = io.BytesIO()
                with pd.ExcelWriter(buf_inst, engine="openpyxl") as writer:
                    df_filtered_inst.to_excel(writer, index=False, sheet_name="Controle CDP Instancias")
                st.download_button(
                    label=f"📥 Baixar Excel Completo ({fmt_num(len(df_filtered_inst))})",
                    data=buf_inst.getvalue(),
                    file_name="Controle_CDP_Instancias.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_cdp_inst"
                )

            total_i = len(df_filtered_inst)
            if total_i > 1000:
                st.caption(f"ℹ️ Exibindo os primeiros **1.000** registros de **{fmt_num(total_i)}** para melhor performance no navegador. Use os filtros acima para refinar a consulta.")
                st.dataframe(df_filtered_inst.head(1000), use_container_width=True, hide_index=True)
            else:
                st.dataframe(df_filtered_inst, use_container_width=True, hide_index=True)

        with tab_metas_det:
            st.markdown(f"#### Relação Analítica de Metas Pactuadas ({fmt_num(len(df_filtered_metas))})")
            st.caption("Tempo decorrido entre o lançamento da meta e os acompanhamentos sucessivos realizados exclusivamente pelo Avaliador 1 (AV1).")

            if not df_filtered_metas.empty:
                # Remove colunas técnicas internas para exibição e exportação limpa
                df_metas_export = df_filtered_metas.drop(columns=["Acomps Eventos", "Data Cadastro Meta ISO"], errors="ignore")

                col_exp_m1, col_exp_m2 = st.columns([1.5, 3.5])
                with col_exp_m1:
                    import io
                    buf_metas = io.BytesIO()
                    with pd.ExcelWriter(buf_metas, engine="openpyxl") as writer:
                        df_metas_export.to_excel(writer, index=False, sheet_name="Detalhamento Metas AV1")
                    st.download_button(
                        label=f"📥 Baixar Excel Completo ({fmt_num(len(df_metas_export))})",
                        data=buf_metas.getvalue(),
                        file_name="Controle_CDP_Metas_Detalhado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_cdp_metas"
                    )

                total_m = len(df_metas_export)
                if total_m > 1000:
                    st.caption(f"ℹ️ Exibindo os primeiros **1.000** registros de **{fmt_num(total_m)}** para melhor performance no navegador. Use os filtros acima para refinar a consulta.")
                    st.dataframe(df_metas_export.head(1000), use_container_width=True, hide_index=True)
                else:
                    st.dataframe(df_metas_export, use_container_width=True, hide_index=True)
            else:
                st.info("Nenhuma meta encontrada para os filtros selecionados.")

        with tab_timeline:
            st.markdown("#### Consulta Individual e Linha do Tempo da Avaliação")
            st.caption("Consulte a linha cronológica de qualquer militar para inspecionar todas as metas pactuadas e seus respectivos acompanhamentos.")

            col_s1, col_s2 = st.columns([2, 2])
            with col_s1:
                busca_indiv = st.text_input("🔍 Digite o Nº PM ou Nome do Militar:", value=busca_aval if busca_aval else "", key="cdp_busca_indiv").strip().lower()

            # Encontrar militares que combinam com a busca
            if busca_indiv:
                candidatos = df_inst[
                    df_inst["nrPM (Avaliado)"].str.lower().str.contains(busca_indiv, na=False) |
                    df_inst["Nome (Avaliado)"].str.lower().str.contains(busca_indiv, na=False)
                ]
            else:
                candidatos = df_filtered_inst

            pms_candidatos = sorted(list(candidatos["nrPM (Avaliado)"].unique()))

            sel_pm_busca = None
            if len(pms_candidatos) == 1:
                sel_pm_busca = pms_candidatos[0]
                with col_s2:
                    st.success(f"Militar selecionado: **{sel_pm_busca} - {candidatos[candidatos['nrPM (Avaliado)'] == sel_pm_busca]['Nome (Avaliado)'].iloc[0]}**")
            elif 1 < len(pms_candidatos) <= 100:
                with col_s2:
                    sel_pm_busca = st.selectbox(
                        f"Selecione entre os {len(pms_candidatos)} militares encontrados:",
                        pms_candidatos,
                        format_func=lambda x: f"{x} - {candidatos[candidatos['nrPM (Avaliado)'] == x]['Nome (Avaliado)'].iloc[0]}",
                        key="cdp_sel_pm"
                    )
            elif len(pms_candidatos) > 100:
                with col_s2:
                    st.info(f"Foram encontrados **{fmt_num(len(pms_candidatos))}** militares. Digite o número PM ou nome mais detalhado para selecionar.")

            if sel_pm_busca:
                # Obter instâncias do militar selecionado
                mil_insts = df_inst[df_inst["nrPM (Avaliado)"] == sel_pm_busca]
                metas_todas_mil = df_metas[df_metas["nrPM (Avaliado)"] == sel_pm_busca] if not df_metas.empty else pd.DataFrame()

                # ── GRÁFICO DE TIMELINE COM DEGRAUS DE METAS E ACOMPANHAMENTOS ─────────
                if not metas_todas_mil.empty:
                    st.markdown("##### 📈 Gráfico de Linha do Tempo e Degraus de Evolução")
                    st.caption("Evolução passo a passo em degraus: da Pactuação inicial aos acompanhamentos sucessivos do AV1 ao longo do tempo.")
                    import plotly.graph_objects as go
                    from plotly.subplots import make_subplots
                    from datetime import datetime

                    # Função local para corrigir artefatos de codificação (ex: funÃ§Ã£o -> função)
                    def _fix_enc(txt):
                        if not txt or not isinstance(txt, str):
                            return txt
                        if "Ã" in txt or "Â" in txt:
                            try:
                                return txt.encode("cp1252").decode("utf-8")
                            except Exception:
                                pass
                        return txt

                    # Deduplicar metas repetidas do mesmo militar
                    metas_plot = metas_todas_mil.drop_duplicates(
                        subset=["Meta", "Descrição da Meta", "Data Cadastro Meta", "Situação Comissão", "nrPM (AV1)"]
                    ).copy()
                    metas_plot["Descrição da Meta"] = metas_plot["Descrição da Meta"].apply(_fix_enc)

                    total_metas_mil = len(metas_plot)

                    # Se o militar possuir mais de 1 meta, disponibilizar seletor para ver todas empilhadas ou focar em uma
                    if total_metas_mil > 1:
                        opcoes_metas = ["Todas as Metas (Degraus Empilhados)"] + [
                            f"{r['Meta']} ({r['Situação Comissão']}): {_fix_enc(r['Descrição da Meta'])[:40]}"
                            for _, r in metas_plot.iterrows()
                        ]
                        sel_meta_view = st.radio(
                            "🎯 Visualização das Metas:",
                            opcoes_metas,
                            horizontal=True,
                            key=f"cdp_sel_meta_{sel_pm_busca}"
                        )
                        if sel_meta_view != "Todas as Metas (Degraus Empilhados)":
                            idx_escolhida = opcoes_metas.index(sel_meta_view) - 1
                            metas_to_render = metas_plot.iloc[[idx_escolhida]]
                        else:
                            metas_to_render = metas_plot
                    else:
                        metas_to_render = metas_plot

                    n_render = len(metas_to_render)
                    hoje_iso = now_br().strftime("%Y-%m-%d")
                    hoje_str = now_br().strftime("%d/%m/%Y")
                    ciclo_fim_iso = f"{_active_y}-06-30"
                    ciclo_fim_str = f"30/06/{_active_y}"

                    if n_render == 1:
                        # Gráfico individual detalhado com degraus verticais
                        m_row = metas_to_render.iloc[0]
                        label_meta = f"{m_row['Meta']} ({m_row['Situação Comissão']}) - {_fix_enc(m_row['Descrição da Meta'])[:50]}"
                        d_meta_iso = m_row.get("Data Cadastro Meta ISO", "")
                        if not d_meta_iso:
                            try:
                                d_meta_iso = datetime.strptime(m_row["Data Cadastro Meta"], "%d/%m/%Y").strftime("%Y-%m-%d")
                            except Exception:
                                d_meta_iso = None

                        acomps_evs = m_row.get("Acomps Eventos", [])
                        max_acomp_num = max([ev["num"] for ev in acomps_evs], default=0)
                        n_steps = max(3, max_acomp_num)

                        tick_vals = list(range(0, n_steps + 1))
                        tick_texts = ["🎯 Pactuação"] + [f"👥 {k}º Acomp" for k in range(1, n_steps + 1)]

                        fig_timeline = go.Figure()

                        # Construir linha em degrau (staircase)
                        x_line = []
                        y_line = []
                        if d_meta_iso:
                            x_line.append(d_meta_iso)
                            y_line.append(0)

                        for ev in acomps_evs:
                            x_line.append(ev["data_iso"])
                            y_line.append(ev["num"])

                        if len(x_line) > 1:
                            fig_timeline.add_trace(go.Scatter(
                                x=x_line,
                                y=y_line,
                                mode='lines',
                                line_shape='hv',
                                line=dict(color='#3b82f6', width=3),
                                hoverinfo='skip',
                                showlegend=False
                            ))

                        # Ponto da Pactuação (Degrau 0)
                        if d_meta_iso:
                            fig_timeline.add_trace(go.Scatter(
                                x=[d_meta_iso],
                                y=[0],
                                mode='markers+text',
                                name='🎯 Pactuação da Meta',
                                marker=dict(size=14, color='#2563eb', symbol='diamond', line=dict(color='white', width=2)),
                                text=[f"  <b>Pactuação</b> ({m_row['Data Cadastro Meta']})"],
                                textposition='middle right',
                                hovertemplate=(
                                    f"<b>{label_meta}</b><br>"
                                    f"Degrau: 🎯 Pactuação da Meta<br>"
                                    f"Data do Cadastro: {m_row['Data Cadastro Meta']}<br>"
                                    f"Prazos Pactuados: {m_row['Prazo Inicial']} a {m_row['Prazo Final']}<extra></extra>"
                                ),
                                showlegend=True
                            ))

                        # Pontos dos Acompanhamentos do AV1 (Degraus 1, 2, 3...)
                        if acomps_evs:
                            x_ac = [ev["data_iso"] for ev in acomps_evs]
                            y_ac = [ev["num"] for ev in acomps_evs]
                            txt_ac = [f"  <b>{ev['num']}º Acomp</b> ({ev['dias']}d - {ev['data_str']})" for ev in acomps_evs]
                            custom_ac = [[ev['num'], ev['data_str'], ev['dias'], _fix_enc(ev['desc'])] for ev in acomps_evs]

                            fig_timeline.add_trace(go.Scatter(
                                x=x_ac,
                                y=y_ac,
                                mode='markers+text',
                                name='👥 Acompanhamento AV1',
                                marker=dict(size=13, color='#10b981', symbol='circle', line=dict(color='white', width=2)),
                                text=txt_ac,
                                textposition='middle right',
                                hovertemplate=(
                                    f"<b>{label_meta}</b><br>"
                                    "Degrau: 👥 %{customdata[0]}º Acompanhamento (AV1)<br>"
                                    "Data: %{customdata[1]}<br>"
                                    "Tempo desde passo anterior: %{customdata[2]} dias<br>"
                                    "Descrição: %{customdata[3]}<extra></extra>"
                                ),
                                customdata=custom_ac,
                                showlegend=True
                            ))

                        # Linha de encerramento do ciclo
                        fig_timeline.add_vline(
                            x=ciclo_fim_iso,
                            line_width=1.5,
                            line_dash="dash",
                            line_color="#f59e0b",
                            annotation_text=f"Fim do Ciclo ({ciclo_fim_str})",
                            annotation_position="top right"
                        )
                        # Linha de hoje
                        fig_timeline.add_vline(
                            x=hoje_iso,
                            line_width=1.5,
                            line_dash="dot",
                            line_color="#ef4444",
                            annotation_text=f"Hoje ({hoje_str})",
                            annotation_position="bottom right"
                        )

                        altura_graf = max(420, (n_steps + 1) * 75)
                        fig_timeline.update_layout(
                            title=f"Linha Cronológica e Degraus — {label_meta} (Militar {sel_pm_busca})",
                            xaxis_title="Linha do Tempo (Datas)",
                            yaxis=dict(
                                tickvals=tick_vals,
                                ticktext=tick_texts,
                                title="Degraus de Evolução",
                                range=[-0.5, n_steps + 0.6]
                            ),
                            height=altura_graf,
                            hovermode="closest",
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                            margin=dict(l=60, r=220, t=80, b=50)
                        )
                        st.plotly_chart(fig_timeline, use_container_width=True)

                    else:
                        # Múltiplas metas em subplots empilhados compartilhando a linha do tempo (shared_xaxes)
                        sub_titles = [
                            f"🎯 {r['Meta']} ({r['Situação Comissão']}): {_fix_enc(r['Descrição da Meta'])[:50]}"
                            for _, r in metas_to_render.iterrows()
                        ]
                        fig_timeline = make_subplots(
                            rows=n_render,
                            cols=1,
                            shared_xaxes=True,
                            vertical_spacing=0.12,
                            subplot_titles=sub_titles
                        )

                        for r_idx, (_, m_row) in enumerate(metas_to_render.iterrows(), start=1):
                            label_meta = f"{m_row['Meta']} ({m_row['Situação Comissão']}) - {_fix_enc(m_row['Descrição da Meta'])[:40]}"
                            d_meta_iso = m_row.get("Data Cadastro Meta ISO", "")
                            if not d_meta_iso:
                                try:
                                    d_meta_iso = datetime.strptime(m_row["Data Cadastro Meta"], "%d/%m/%Y").strftime("%Y-%m-%d")
                                except Exception:
                                    d_meta_iso = None

                            acomps_evs = m_row.get("Acomps Eventos", [])
                            max_acomp_num = max([ev["num"] for ev in acomps_evs], default=0)
                            n_steps = max(3, max_acomp_num)

                            tick_vals = list(range(0, n_steps + 1))
                            tick_texts = ["🎯 Pactuação"] + [f"👥 {k}º Acomp" for k in range(1, n_steps + 1)]

                            # Linha do degrau
                            x_line = []
                            y_line = []
                            if d_meta_iso:
                                x_line.append(d_meta_iso)
                                y_line.append(0)

                            for ev in acomps_evs:
                                x_line.append(ev["data_iso"])
                                y_line.append(ev["num"])

                            if len(x_line) > 1:
                                fig_timeline.add_trace(go.Scatter(
                                    x=x_line,
                                    y=y_line,
                                    mode='lines',
                                    line_shape='hv',
                                    line=dict(color='#3b82f6', width=3),
                                    hoverinfo='skip',
                                    showlegend=False
                                ), row=r_idx, col=1)

                            # Ponto Pactuação
                            if d_meta_iso:
                                fig_timeline.add_trace(go.Scatter(
                                    x=[d_meta_iso],
                                    y=[0],
                                    mode='markers+text',
                                    name='🎯 Pactuação da Meta' if r_idx == 1 else '',
                                    marker=dict(size=14, color='#2563eb', symbol='diamond', line=dict(color='white', width=2)),
                                    text=[f"  <b>Pactuação</b> ({m_row['Data Cadastro Meta']})"],
                                    textposition='middle right',
                                    hovertemplate=(
                                        f"<b>{label_meta}</b><br>"
                                        f"Degrau: 🎯 Pactuação da Meta<br>"
                                        f"Data: {m_row['Data Cadastro Meta']}<br>"
                                        f"Prazos: {m_row['Prazo Inicial']} a {m_row['Prazo Final']}<extra></extra>"
                                    ),
                                    showlegend=(r_idx == 1)
                                ), row=r_idx, col=1)

                            # Pontos Acompanhamentos
                            if acomps_evs:
                                x_ac = [ev["data_iso"] for ev in acomps_evs]
                                y_ac = [ev["num"] for ev in acomps_evs]
                                txt_ac = [f"  <b>{ev['num']}º Acomp</b> ({ev['dias']}d - {ev['data_str']})" for ev in acomps_evs]
                                custom_ac = [[ev['num'], ev['data_str'], ev['dias'], _fix_enc(ev['desc'])] for ev in acomps_evs]

                                fig_timeline.add_trace(go.Scatter(
                                    x=x_ac,
                                    y=y_ac,
                                    mode='markers+text',
                                    name='👥 Acompanhamento AV1' if r_idx == 1 else '',
                                    marker=dict(size=13, color='#10b981', symbol='circle', line=dict(color='white', width=2)),
                                    text=txt_ac,
                                    textposition='middle right',
                                    hovertemplate=(
                                        f"<b>{label_meta}</b><br>"
                                        "Degrau: 👥 %{customdata[0]}º Acompanhamento (AV1)<br>"
                                        "Data: %{customdata[1]}<br>"
                                        "Tempo desde passo anterior: %{customdata[2]} dias<br>"
                                        "Descrição: %{customdata[3]}<extra></extra>"
                                    ),
                                    customdata=custom_ac,
                                    showlegend=(r_idx == 1)
                                ), row=r_idx, col=1)

                            # Configurar eixo Y da linha do subplot
                            fig_timeline.update_yaxes(
                                tickvals=tick_vals,
                                ticktext=tick_texts,
                                range=[-0.5, n_steps + 0.6],
                                row=r_idx,
                                col=1
                            )

                            # Adicionar linhas de referência
                            fig_timeline.add_vline(
                                x=ciclo_fim_iso,
                                line_width=1.5,
                                line_dash="dash",
                                line_color="#f59e0b",
                                annotation_text=f"Fim Ciclo ({ciclo_fim_str})" if r_idx == 1 else None,
                                annotation_position="top right",
                                row=r_idx, col=1
                            )
                            fig_timeline.add_vline(
                                x=hoje_iso,
                                line_width=1.5,
                                line_dash="dot",
                                line_color="#ef4444",
                                annotation_text=f"Hoje ({hoje_str})" if r_idx == 1 else None,
                                annotation_position="bottom right",
                                row=r_idx, col=1
                            )

                        altura_graf = max(450, n_render * 280)
                        fig_timeline.update_layout(
                            title=f"Linha Cronológica e Degraus de Todas as Metas do Militar {sel_pm_busca}",
                            xaxis_title="Linha do Tempo (Datas)",
                            height=altura_graf,
                            hovermode="closest",
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                            margin=dict(l=60, r=220, t=80, b=50)
                        )
                        st.plotly_chart(fig_timeline, use_container_width=True)

                    st.markdown("---")

                # Detalhes das instâncias
                for idx_inst, inst_row in mil_insts.iterrows():
                    with st.container():
                        st.markdown(f"##### 🎖️ Avaliado: {inst_row['Posto/Grad. (Avaliado)']} {inst_row['Nome (Avaliado)']} ({inst_row['nrPM (Avaliado)']})")
                        col_i1, col_i2, col_i3, col_i4 = st.columns(4)
                        col_i1.info(f"**Situação Comissão:** {inst_row['Situação Comissão']}")
                        col_i2.info(f"**Status CDP:** {inst_row['Status CDP']} - **Data:** {inst_row['Data do CDP']}")
                        col_i3.info(f"**Avaliador 1 (AV1):** {inst_row['Posto (AV1)']} {inst_row['Nome (AV1)']} ({inst_row['nrPM (AV1)']})")
                        col_i4.info(f"**Último Lançamento:** {inst_row['Último Lançamento']} ({inst_row['Dias desde Último Lançamento']} dias)")

                        # Metas dessa instância
                        metas_mil = df_metas[
                            (df_metas["nrPM (Avaliado)"] == sel_pm_busca) &
                            (df_metas["nrPM (AV1)"] == inst_row["nrPM (AV1)"])
                        ]

                        if not metas_mil.empty:
                            st.markdown("**Detalhamento de Prazos das Metas desta Instância:**")
                            for _, m_row in metas_mil.iterrows():
                                with st.expander(f"🎯 {m_row['Meta']} — {m_row['Descrição da Meta']} (Cadastrada em {m_row['Data Cadastro Meta']})", expanded=True):
                                    st.write(f"**Prazos Pactuados:** {m_row['Prazo Inicial']} a {m_row['Prazo Final']}")
                                    st.write(f"**Acompanhamentos AV1:** {m_row['Qtd Acomps AV1']} | **Outros Acompanhamentos:** {m_row['Qtd Outros Acomps']}")
                                    if m_row["Qtd Acomps AV1"] > 0:
                                        st.success(f"⏱️ **Tempo até 1º Acompanhamento:** {m_row['Dias (Meta -> 1º Acomp)']} dias | **Intervalos entre acompanhamentos:** {m_row['Intervalos Acomps']} | **Média da Meta:** {m_row['Média Dias Lançamentos']} dias")
                                    else:
                                        st.warning("⚠️ Nenhum acompanhamento realizado pelo AV1 até o momento.")
                        else:
                            st.warning("Nenhuma meta cadastrada para esta instância.")
                        st.markdown("---")
            else:
                st.info("Nenhum militar localizado com os filtros atuais.")


st.markdown("---")
st.markdown(f"<center><small>AADP 2026 · Polícia Militar de Minas Gerais · "
            f"Resolução 5458/2025 · {now_br().strftime('%d/%m/%Y')}<br>"
            f"DIRETORIA DE RECURSO HUMANOS - DRH6</small></center>",
            unsafe_allow_html=True)

print(f"[AADP PROFILE] Script finished in {time.time() - _prof_start:.4f} seconds", flush=True)