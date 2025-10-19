# app.py
# Autoavaliação de Maturidade NP 4427 — Streamlit (Excel por URL/Google Sheets)
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
st.set_page_config(
    page_title="Autoavaliação NP 4427 — STEM CoE",
    page_icon="✅",
    layout="wide",
)

# 👉 LIGAÇÃO DIRETA AO TEU GOOGLE SHEETS
EXCEL_URL = "https://docs.google.com/spreadsheets/d/1l576pvLwiC0kH3WverA2RDUra3x_k76Q/export?format=xlsx"

APP_TITLE = "Autoavaliação de Maturidade — NP 4427"
APP_SUBTITLE = "STEM Data Analytics & AI CoE · Data&AI4All"

PALETTE = {
    "primary": "#0C4A6E",
    "accent":  "#0EA5E9",
    "gray":    "#F1F5F9",
    "ok":      "#16A34A",
    "warn":    "#F59E0B",
    "bad":     "#DC2626",
}

# =========================
# 🎨 CABEÇALHO
# =========================
st.markdown(
    f"""
    <div style="padding:8px 0 0 0">
        <h1 style="margin:0;color:{PALETTE['primary']};">{APP_TITLE}</h1>
        <p style="margin:0;color:#334155;font-size:16px;">{APP_SUBTITLE}</p>
    </div>
    """,
    unsafe_allow_html=True
)
st.markdown("---")

# =========================
# 📥 FUNÇÕES
# =========================
@st.cache_data(show_spinner=True, ttl=600)
def load_excel_from_url(url: str) -> pd.ExcelFile:
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    return pd.ExcelFile(io.BytesIO(resp.content))

def interpret_level(x: float) -> str:
    if x < 2: return "Inicial"
    if x < 3: return "Básico"
    if x < 4: return "Padronizado"
    if x <= 4.5: return "Gerido"
    return "Otimizado"

# =========================
# 📦 CARREGAR MODELO
# =========================
with st.spinner("A carregar modelo NP 4427..."):
    xls = load_excel_from_url(EXCEL_URL)
df = pd.read_excel(xls, sheet_name="Checklist & Autoavaliação")

df["Peso (%)"] = pd.to_numeric(df["Peso (%)"], errors="coerce").fillna(0)
pillars = df["Pilar / Dimensão"].dropna().unique().tolist()

# =========================
# 🧑‍🎓 IDENTIFICAÇÃO
# =========================
st.subheader("Identificação do Participante")
nome = st.text_input("Nome", "")
email = st.text_input("Email (opcional)", "")
turma = st.text_input("Turma/Grupo (opcional)", "")
st.markdown("---")

# =========================
# 📝 FORMULÁRIO
# =========================
st.subheader("Preenchimento do Checklist (Nível 1–5)")
st.caption("Selecione o nível para cada requisito (1=Inicial, 5=Otimizado)")

respostas = {}
with st.form("avaliacao_form"):
    for pillar in pillars:
        bloco = df[df["Pilar / Dimensão"] == pillar]
        with st.expander(f"📌 {pillar}", expanded=False):
            for _, row in bloco.iterrows():
                codigo = row["Código"]
                requisito = row["Requisito (NP 4427)"]
                descricao = row["Descrição / Pergunta de Avaliação"]
                peso = row["Peso (%)"]
                chave = f"{pillar}::{codigo}"
                col1, col2 = st.columns([5, 1])
                with col1:
                    st.markdown(f"**{codigo} – {requisito}**  \n<small>{descricao}</small>", unsafe_allow_html=True)
                with col2:
                    respostas[chave] = st.select_slider(
                        "Nível", options=[1, 2, 3, 4, 5],
                        value=3, label_visibility="collapsed", key=chave
                    )
                st.caption(f"Peso: {peso:.0f}%")
    submitted = st.form_submit_button("Calcular Maturidade", use_container_width=True)

# =========================
# 📊 CÁLCULOS
# =========================
def calcular(df_base, respostas):
    df = df_base.copy()
    df["Nível"] = [respostas.get(f"{r['Pilar / Dimensão']}::{r['Código']}", np.nan)
                   for _, r in df.iterrows()]
    df["Pontuação"] = df["Nível"] * df["Peso (%)"]
    por_pilar = df.groupby("Pilar / Dimensão").agg(
        Média=("Nível", "mean"),
        Peso=("Peso (%)", "sum")
    ).reset_index()
    nivel_global = np.average(df["Nível"], weights=df["Peso (%)"])
    return df, por_pilar, nivel_global

if not submitted:
    st.info("Preenche o formulário e clica em **Calcular Maturidade**.")
    st.stop()

df_calc, df_pilar, nivel_global = calcular(df, respostas)
interp = interpret_level(nivel_global)

# =========================
# 🧭 RESUMO
# =========================
col1, col2, col3 = st.columns(3)
col1.metric("Nível Global", f"{nivel_global:.2f}")
col2.metric("Interpretação", interp)
col3.metric("Requisitos ≥4", f"{100*(df_calc['Nível']>=4).mean():.1f}%")
st.markdown("---")

# =========================
# 📈 GRÁFICOS
# =========================
radar = go.Figure()
radar.add_trace(go.Scatterpolar(
    r=df_pilar["Média"].tolist() + [df_pilar["Média"].iloc[0]],
    theta=df_pilar["Pilar / Dimensão"].tolist() + [df_pilar["Pilar / Dimensão"].iloc[0]],
    fill='toself'
))
radar.update_layout(title="Radar — Maturidade por Pilar",
                    polar=dict(radialaxis=dict(visible=True, range=[0,5])))
st.plotly_chart(radar, use_container_width=True)

bar = go.Figure(data=[go.Bar(
    x=df_pilar["Pilar / Dimensão"], y=df_pilar["Média"],
    text=np.round(df_pilar["Média"], 2), textposition="auto")])
bar.update_layout(title="Barras — Média de Maturidade", yaxis=dict(range=[0,5]))
st.plotly_chart(bar, use_container_width=True)

# =========================
# 💾 GUARDAR
# =========================
st.markdown("---")
if nome.strip():
    export = df_calc[[
        "Pilar / Dimensão", "Código", "Requisito (NP 4427)", "Descrição / Pergunta de Avaliação",
        "Nível", "Peso (%)", "Pontuação"
    ]]
    export.insert(0, "Nome", nome)
    export.insert(1, "Email", email)
    export.insert(2, "Turma", turma)
    export["Nível Global"] = nivel_global
    export["Interpretação"] = interp
    export["Data"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")
    st.download_button(
        label="💾 Guardar resultado (CSV)",
        data=export.to_csv(index=False).encode("utf-8"),
        file_name=f"NP4427_{nome.replace(' ','_')}.csv",
        mime="text/csv"
    )
else:
    st.warning("Preenche o nome para permitir guardar o resultado.")
