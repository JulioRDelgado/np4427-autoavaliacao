# app.py
# Autoavaliação de Maturidade NP 4427 — Streamlit (Google Sheets por URL, com fallback de URL/Upload)
# Autor: STEM Data Analytics & AI CoE (Data&AI4All)

import io
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import requests

# =========================
# 🔧 CONFIGURAÇÃO
# =========================
st.set_page_config(page_title="Autoavaliação NP 4427 — STEM CoE", page_icon="✅", layout="wide")

APP_TITLE = "Autoavaliação de Maturidade — NP 4427"
APP_SUBTITLE = "STEM Data Analytics & AI CoE · Data&AI4All"

# TEU Google Sheets (partilhado como Anyone with the link → Viewer)
# Transformado para exportação XLSX direta:
EXCEL_URL_DEFAULT = (
    "https://docs.google.com/spreadsheets/d/"
    "1eWjwt2qXE_g3ZJIzc2PD7V16VzHuIGb__eqgcrQ04q4"
    "/export?format=xlsx"
)

# =========================
# 🎨 CABEÇALHO
# =========================
st.markdown(
    f"""
    <div style="padding:8px 0 0 0">
        <h1 style="margin:0;color:#0C4A6E;">{APP_TITLE}</h1>
        <p style="margin:0;color:#334155;font-size:16px;">{APP_SUBTITLE}</p>
    </div>
    """,
    unsafe_allow_html=True
)
st.markdown("---")

# =========================
# 📥 CARREGAR MODELO (URL + fallback)
# =========================
@st.cache_data(show_spinner=True, ttl=600)
def load_excel_from_url(url: str) -> pd.ExcelFile:
    r = requests.get(url, timeout=30)
    r.raise_for_status()  # dispara erro se 403/404
    return pd.ExcelFile(io.BytesIO(r.content))

st.caption("A app lê o modelo diretamente do Google Sheets. Se necessário, podes indicar outro URL ou carregar um .xlsx.")
c1, c2 = st.columns([3, 2])
with c1:
    excel_url = st.text_input("URL do modelo (opcional, substitui o padrão)", value=EXCEL_URL_DEFAULT)
with c2:
    uploaded = st.file_uploader("Ou carregar .xlsx", type=["xlsx"])

with st.spinner("A carregar modelo NP 4427..."):
    try:
        if uploaded is not None:
            xls = pd.ExcelFile(uploaded.read())
        else:
            xls = load_excel_from_url(excel_url.strip() or EXCEL_URL_DEFAULT)
    except Exception as e:
        st.error(
            "❌ Falhou o carregamento do modelo. "
            "Verifica se o Google Sheets está partilhado como **Anyone with the link → Viewer** "
            "ou usa o carregador .xlsx. Detalhe: " + str(e)
        )
        st.stop()

SHEET_NAME = "Checklist & Autoavaliação"
if SHEET_NAME not in xls.sheet_names:
    st.error(f"A folha **'{SHEET_NAME}'** não existe. Folhas disponíveis: {xls.sheet_names}")
    st.stop()

df = pd.read_excel(xls, sheet_name=SHEET_NAME)
# sanity check a colunas essenciais
required = ["Pilar / Dimensão", "Código", "Requisito (NP 4427)",
            "Descrição / Pergunta de Avaliação", "Peso (%)"]
miss = [c for c in required if c not in df.columns]
if miss:
    st.error(f"Colunas obrigatórias em falta no modelo: {miss}")
    st.stop()

df["Peso (%)"] = pd.to_numeric(df["Peso (%)"], errors="coerce").fillna(0)
pillars = df["Pilar / Dimensão"].dropna().unique().tolist()

def interpret_level(x: float) -> str:
    if x < 2: return "Inicial"
    if x < 3: return "Básico"
    if x < 4: return "Padronizado"
    if x <= 4.5: return "Gerido"
    return "Otimizado"

# =========================
# 🧑‍🎓 IDENTIFICAÇÃO
# =========================
st.subheader("Identificação do Participante")
colA, colB, colC = st.columns([3,3,2])
with colA:
    nome = st.text_input("Nome", "")
with colB:
    email = st.text_input("Email (opcional)", "")
with colC:
    turma = st.text_input("Turma/Grupo (opcional)", "")
st.markdown("---")

# =========================
# 📝 FORMULÁRIO
# =========================
st.subheader("Preenchimento do Checklist (Nível 1–5)")
st.caption("Selecione o nível para cada requisito (1=Inicial · 5=Otimizado)")
respostas = {}
with st.form("avaliacao_form"):
    for pillar in pillars:
        bloco = df[df["Pilar / Dimensão"] == pillar].reset_index(drop=True)
        with st.expander(f"📌 {pillar}", expanded=False):
            for _, row in bloco.iterrows():
                codigo = str(row["Código"])
                requisito = str(row["Requisito (NP 4427)"])
                descricao = str(row["Descrição / Pergunta de Avaliação"]).strip()
                peso = float(row["Peso (%)"])
                key = f"{pillar}::{codigo}"

                cL, cR = st.columns([5,1])
                with cL:
                    st.markdown(f"**{codigo} – {requisito}**  \n<small>{descricao}</small>", unsafe_allow_html=True)
                with cR:
                    respostas[key] = st.select_slider("Nível", options=[1,2,3,4,5],
                                                      value=3, label_visibility="collapsed", key=key)
                st.caption(f"Peso: **{peso:.0f}%**")
                st.divider()
    submitted = st.form_submit_button("Calcular Maturidade", use_container_width=True)

if not submitted:
    st.info("Preenche o formulário e clica em **Calcular Maturidade**.")
    st.stop()

# =========================
# 📊 CÁLCULOS
# =========================
df_calc = df.copy()
df_calc["Nível"] = [
    respostas.get(f"{r['Pilar / Dimensão']}::{r['Código']}", np.nan) for _, r in df_calc.iterrows()
]
df_calc["Pontuação"] = df_calc["Nível"] * df_calc["Peso (%)"]

df_pilar = df_calc.groupby("Pilar / Dimensão", dropna=True).agg(
    Média=("Nível", "mean"),
    Peso=("Peso (%)", "sum"),
    Pontuação=("Pontuação", "sum"),
    Itens=("Nível", "count")
).reset_index()

soma_pesos = df_calc["Peso (%)"].sum()
nivel_global = float(np.nansum(df_calc["Nível"] * df_calc["Peso (%)"]) / soma_pesos) if soma_pesos else np.nan
interp = interpret_level(nivel_global) if not np.isnan(nivel_global) else "—"

# =========================
# 🧭 RESUMO EXECUTIVO
# =========================
k1, k2, k3, k4 = st.columns(4)
k1.metric("Nível Global", f"{nivel_global:.2f}" if not np.isnan(nivel_global) else "—")
k2.metric("Interpretação", interp)
k3.metric("% Requisitos ≥4", f"{100*(df_calc['Nível']>=4).mean():.1f}%")
k4.metric("% Requisitos ≤2", f"{100*(df_calc['Nível']<=2).mean():.1f}%")
st.markdown("---")

# =========================
# 📈 GRÁFICOS
# =========================
cL, cR = st.columns([2,1])

with cL:
    # Radar
    if len(df_pilar) > 0:
        cats = df_pilar["Pilar / Dimensão"].tolist()
        vals = df_pilar["Média"].fillna(0).tolist()
        radar = go.Figure()
        radar.add_trace(go.Scatterpolar(r=vals+[vals[0]], theta=cats+[cats[0]], fill='toself'))
        radar.update_layout(title="Radar — Maturidade por Pilar",
                            polar=dict(radialaxis=dict(visible=True, range=[0,5])),
                            showlegend=False)
        st.plotly_chart(radar, use_container_width=True)

    # Barras
    bar = go.Figure(data=[go.Bar(x=df_pilar["Pilar / Dimensão"], y=df_pilar["Média"],
                                 text=np.round(df_pilar["Média"],2), textposition="auto")])
    bar.update_layout(title="Barras — Média de Maturidade por Pilar", yaxis=dict(range=[0,5]))
    st.plotly_chart(bar, use_container_width=True)

with cR:
    # Pizza (Peso por Pilar)
    pie = go.Figure(data=[go.Pie(labels=df_pilar["Pilar / Dimensão"], values=df_pilar["Peso"], hole=.35)])
    pie.update_layout(title="Peso Relativo por Pilar")
    st.plotly_chart(pie, use_container_width=True)

st.markdown("---")

# =========================
# 🧾 TABELAS
# =========================
t1, t2 = st.columns(2)
with t1:
    st.subheader("Resumo por Pilar")
    st.dataframe(
        df_pilar.rename(columns={"Pilar / Dimensão":"Pilar","Média":"Média (1–5)","Peso":"Peso (%)"}),
        use_container_width=True
    )
with t2:
    st.subheader("Top 5 — Pontos Fortes / Áreas a Melhorar")
    fortes = df_pilar.nlargest(5, "Média")[["Pilar / Dimensão","Média"]].rename(columns={"Pilar / Dimensão":"Pilar","Média":"Média (1–5)"})
    fracos = df_pilar.nsmallest(5, "Média")[["Pilar / Dimensão","Média"]].rename(columns={"Pilar / Dimensão":"Pilar","Média":"Média (1–5)"})
    st.markdown("**Pontos Fortes**"); st.table(fortes)
    st.markdown("**Áreas de Melhoria**"); st.table(fracos)

# =========================
# 💾 GUARDAR RESULTADO (CSV)
# =========================
st.markdown("---")
st.subheader("Guardar resultado")
if nome.strip():
    export = df_calc[[
        "Pilar / Dimensão","Código","Requisito (NP 4427)","Descrição / Pergunta de Avaliação",
        "Nível","Peso (%)","Pontuação"
    ]].copy()
    export.insert(0, "Nome", nome)
    export.insert(1, "Email", email)
    export.insert(2, "Turma", turma)
    export["Nível Global"] = round(nivel_global, 2) if not np.isnan(nivel_global) else ""
    export["Interpretação"] = interp
    export["Timestamp"] = pd.Timestamp.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
    st.download_button(
        "💾 Guardar resultado (CSV)",
        data=export.to_csv(index=False).encode("utf-8"),
        file_name=f"NP4427_Autoavaliacao_{nome.strip().replace(' ','_')}.csv",
        mime="text/csv",
        use_container_width=True
    )
else:
    st.warning("Preenche o **Nome** no topo para permitir o download do resultado.")
