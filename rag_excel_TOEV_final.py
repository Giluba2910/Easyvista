import os
import re
import tempfile
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="SAP Transport Orders & EasyVista Tool", layout="wide")
st.title("SAP Transport Order & EasyVista Tool")
st.caption("Filtrage multi-zones + recherche textuelle sur TO Short Description + EV Title")

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_FILE = BASE_DIR / "data" / "SAP_TOEV.XLSX"

if not DEFAULT_FILE.exists():
    st.error(f"Fichier Excel introuvable: {DEFAULT_FILE}")
    st.stop()

@st.cache_data(show_spinner=False)
def load_sheet(file_source: str):
    df = pd.read_excel(file_source, sheet_name="TO&EV")
    df.columns = [str(c).strip() for c in df.columns]
    date_cols = [c for c in df.columns if "date" in c.lower()]
    for c in date_cols:
        df[c] = pd.to_datetime(df[c], errors="coerce").dt.strftime("%d/%m/%y")
    df = df.fillna("")
    return df

def normalize_text(x):
    return re.sub(r"\s+", " ", str(x).lower()).strip()

TRANSLATION_MAP = {
    "paiement": "payment",
    "terme": "term",
    "conditions": "condition",
    "facture": "invoice",
    "client": "customer",
    "prix": "price",
    "commande": "order",
    "livraison": "delivery",
    "demande": "request",
    "statut": "status",
    "projet": "project",
    "priorite": "priority",
    "élevée": "high",
    "elevee": "high",
    "propriétaire": "owner",
    "proprietaire": "owner",
    "équipe": "team",
    "equipe": "team",
    "texte": "text",
    "description": "description",
    "fermer": "closed",
    "ouvert": "open",
}

def normalize_query_terms(question: str):
    q = normalize_text(question)
    tokens = [t for t in re.split(r"\W+", q) if len(t) > 1]
    expanded = []
    for tok in tokens:
        expanded.append(tok)
        if tok in TRANSLATION_MAP:
            expanded.append(TRANSLATION_MAP[tok])
        for fr, en in TRANSLATION_MAP.items():
            if tok == en:
                expanded.append(fr)
    return list(dict.fromkeys(expanded))

def build_filters(df: pd.DataFrame):
    filter_cols = [
        "Request/Task",
        "Object Type Text",
        "Object Name",
        "Ticket Support Nr",
        "EV.Request Type",
        "EV.Request Type Last Level",
        "EV.Project",
        "EV.Lead Stream",
        "EV.Ticket Owner",
        "EV.Action Status",
        "EV.Cluster",
        "EV.Country",
        "EV.Priority High Level",
    ]
    filters = {}
    st.sidebar.header("Filtres")
    for col in filter_cols:
        if col in df.columns:
            vals = sorted([v for v in df[col].astype(str).unique().tolist() if v != ""])
            selected = st.sidebar.multiselect(col, vals, default=[])
            if selected:
                filters[col] = selected
    return filters

def apply_filters(df: pd.DataFrame, filters: dict):
    out = df.copy()
    for col, vals in filters.items():
        out = out[out[col].astype(str).isin(vals)]
    return out

def tokenize_query(question: str):
    q = normalize_text(question)
    q = q.replace("(", " ( ").replace(")", " ) ")
    return [t for t in re.split(r"\s+", q) if t]

def term_mask(df: pd.DataFrame, term: str):
    cols = [c for c in ["Short Description", "EV.Title"] if c in df.columns]
    if not cols or term == "":
        return pd.Series([True] * len(df), index=df.index)
    if term in {"and", "or", "not", "(", ")"}:
        raise ValueError("Reserved token")
    expanded = normalize_query_terms(term)
    pattern = re.escape(term)
    masks = []
    for col in cols:
        series = df[col].astype(str).str.lower()
        m = series.str.contains(pattern, na=False)
        for tok in expanded:
            m = m | series.str.contains(re.escape(tok), na=False)
        masks.append(m)
    out = masks[0]
    for m in masks[1:]:
        out = out | m
    return out

def query_mask(df: pd.DataFrame, question: str):
    tokens = tokenize_query(question)
    if not tokens:
        return pd.Series([True] * len(df), index=df.index)

    result = None
    op = None
    negate_next = False
    stack = []

    def apply_op(left, right, current_op):
        if current_op == "and":
            return left & right
        return left | right

    i = 0
    while i < len(tokens):
        tok = tokens[i]
        if tok == "(":
            stack.append((result, op, negate_next))
            result, op, negate_next = None, None, False
            i += 1
            continue
        if tok == ")":
            if stack:
                prev_result, prev_op, prev_neg = stack.pop()
                if prev_result is None:
                    result = ~result if prev_neg else result
                else:
                    if prev_neg:
                        result = ~result
                    result = apply_op(prev_result, result, prev_op or "and")
                negate_next = False
                op = None
            i += 1
            continue
        if tok in {"and", "or"}:
            op = tok
            i += 1
            continue
        if tok == "not":
            negate_next = True
            i += 1
            continue

        m = term_mask(df, tok)
        if negate_next:
            m = ~m
            negate_next = False
        if result is None:
            result = m
        else:
            result = apply_op(result, m, op or "and")
        op = None
        i += 1

    if result is None:
        return pd.Series([True] * len(df), index=df.index)
    return result

def get_unique_requests(df: pd.DataFrame):
    if "Request/Task" not in df.columns:
        return []
    return df["Request/Task"].astype(str).dropna().drop_duplicates().tolist()

def export_excel(df: pd.DataFrame):
    out_path = Path(tempfile.gettempdir()) / f"sap_to_ev_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="results")
    return out_path

source_mode = st.radio("Source", ["Fichier préchargé", "Téléverser un fichier"], horizontal=True)
file_source = str(DEFAULT_FILE)
if source_mode == "Téléverser un fichier":
    uploaded = st.file_uploader("Choisir un .xlsx", type=["xlsx"])
    if uploaded is not None:
        temp_path = Path(tempfile.gettempdir()) / uploaded.name
        with open(temp_path, "wb") as f:
            f.write(uploaded.getbuffer())
        file_source = str(temp_path)
else:
    st.info("Téléversez un fichier ou utilisez le fichier préchargé.")

if "history" not in st.session_state:
    st.session_state.history = []
if "last_export" not in st.session_state:
    st.session_state.last_export = None

if os.path.exists(file_source):
    try:
        df = load_sheet(file_source)
        st.success(f"Feuille TO&EV chargée: {len(df)} lignes, {len(df.columns)} colonnes.")

        filters = build_filters(df)
        filtered = apply_filters(df, filters)
        st.sidebar.write(f"Lignes après filtres: {len(filtered)}")

        q_col, r_col = st.columns([4, 1])
        with q_col:
            question = st.text_input("Question métier ( Use logical operators:AND / OR / NOT)")
        with r_col:
            reset_clicked = st.button("Réinitialiser")
        if reset_clicked:
            st.session_state.history = []
            st.session_state.last_export = None
            st.rerun()

        if st.button("Rechercher", type="primary"):
            if question.strip() == "":
                result_df = filtered.copy()
            else:
                q_mask = query_mask(filtered, question)
                result_df = filtered[q_mask].copy()

            hidden_cols = [c for c in ["Internal comment", "Approved by", "Changed on"] if c in result_df.columns]
            display_df = result_df.drop(columns=hidden_cols, errors="ignore")

            unique_requests = get_unique_requests(result_df)
            st.session_state.history.append({
                "question": question if question.strip() else "(vide - toutes les entrées)",
                "result_df": result_df,
                "display_df": display_df,
                "unique_requests": unique_requests,
            })
            st.session_state.last_export = export_excel(display_df)

        if st.session_state.history:
            st.subheader("Historique")
            for i, item in enumerate(reversed(st.session_state.history), 1):
                st.markdown(f"### Q{i}: {item['question']}")
                if item["unique_requests"]:
                    st.markdown(f"**Request/Task trouvés:** {len(item['unique_requests'])}")
                    st.write(item["unique_requests"])
                else:
                    st.write("Aucun Request/Task trouvé.")
                st.dataframe(item["display_df"], use_container_width=True)

        if st.session_state.last_export:
            with open(st.session_state.last_export, "rb") as f:
                st.download_button(
                    "Télécharger la table filtrée en .xlsx",
                    data=f,
                    file_name=Path(st.session_state.last_export).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        with st.sidebar:
            st.header("Session")
            st.write(f"Questions en mémoire: {len(st.session_state.history)}")
            if st.button("Réinitialiser tout"):
                st.session_state.history = []
                st.session_state.last_export = None
                st.rerun()

    except Exception as e:
        st.error(f"Erreur: {e}")
else:
    st.warning("Fichier Excel introuvable.")