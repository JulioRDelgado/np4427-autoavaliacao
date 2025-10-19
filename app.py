# app.py
# Autoavaliação de Maturidade NP 4427 — Streamlit (Google Sheets por URL, com fallback URL/Upload)
import io
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import requests

st.set_page_config(page_title="Autoavaliação NP 4427 — STEM CoE", page_icon="✅", layout="wide")

APP_TITLE = "Autoavaliação de Maturidade — NP 4427"
APP_SUBTITLE = "STEM Data Analytics & AI CoE · Data&AI4All"

# Google Sheets (partilhado: Anyone with the link → Viewer)
EXCEL_URL_DEFAULT = (
    "https://docs.google.com/spreadsheets/d/"
    "1eWjwt2qXE_g3ZJIzc2PD7V16VzHuIGb__eqgcrQ04q4"
    "/export?format=xlsx"
)
SHEET_NAME = "Checklist & Autoavaliação"

# ------------------------
# Helpers
# ------------------------
def interpret_level(x: float) -> str:
    if x < 2: return "Inicial"
    if x < 3: return "Básico"
    if x < 4: return "Padronizado"
    if x <= 4.5: return "Gerido"
    return "Otimizado"

@st.cache_data(show_spinner=True, ttl=600)
def read_checklist_from_url(url: str) -> pd.DataFrame:
    """Descarrega o XLSX do URL e devolve diretamente a folha Checklist como DataFrame (pickle-able)."""
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    with pd.ExcelFile(io.BytesIO(r.content)) as xls:
        if SHEET_NAME not in xls.sheet_names:
            raise ValueError(f"Folha '{SHEET_NAME}' não encontrada. Folhas: {xls.sheet_names}")
        df = pd.read_excel(xls, sheet_name=SHEET_NAME)
    return df

@st.cache_data(show_spinner=True)
def read_checklist_from_bytes(file_bytes: bytes) -> pd.DataFrame:
    """Lê um .xlsx enviado (upload) e devolve a folha Checklist."""
    with pd.ExcelFile(io.BytesIO(file_bytes)) as xls:
        if SHEET_NAME not in xls.sheet_names:
            raise ValueError(f"Folha '{SHEET_NAME}' não encontrada. Folhas: {xls.sheet_names}")
        df = pd.read_excel(xls, sheet_name=SHEET_NAME)
    return df

# ------------------------
# Branding
# ------------------------
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

# ------------------------
# Carregamento do modelo (URL + fallback)
# ------------------------
st.caption("A app lê o modelo do Google Sheets. Se necessário, podes indicar outro URL ou carregar um .xlsx.")
c1, c2 = st.columns([3,2])
with c1:
    excel_url = st.text_input("URL do modelo (opcional)", value=EXCEL_URL_DEFAULT)
with c2:
    uploaded = st.file_uploader("Ou carregar .xlsx", type=["xlsx"])

with st.spinner("A carregar modelo NP 4427..."):
    try:
        if uploaded is not None:
            df = read_checklist_from_bytes(uploaded.read())
        else:
            df = read_checklist_from_url((excel_url or EXCEL_URL_DEFAULT).strip())
    except Exception as e:
        st.error(
            "❌ Falhou o carregamento do modelo. "
            "Confirma no Google Sheets: **Anyone with the link → Viewer**. "
            "Podes também usar o upload .xlsx. Detalhe: " + str(e)
        )
        st.stop()

# Verificações mínimas
required = ["Pilar / Dimensão","Código","Requisito (NP 4427)","Descrição / Pergunta de Avaliação","Peso (%)"]
miss = [c for c in required if c not in df.columns]
if miss:
    st.error(f"Colunas obrigatórias em falta no modelo: {miss}")
    st.stop()

df["Peso (%)"] = pd.to_numeric(df["Peso (%)"], errors="coerce").fillna(0.0)
pillars = df["Pilar / Dimensão"].dropna().unique().tolist()

# ------------------------
# Identificação
# ------------------------
st.subheader("Identificação do Participante")
colA, colB, colC = st.columns([3,3,2])
with colA: nome = st.text_input("Nome", "")
with colB: email = st.text_input("Email (opcional)", "")
with colC: turma = st.text_input("Turma/Grupo (opcional)", "")
st.markdown("---")

# ------------------------
# Formulário
# ------------------------
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

# ------------------------
# Cálculos
# ------------------------
df_calc = df.copy()
df_calc["Nível"] = [respostas.get(f"{r['Pilar / Dimensão']}::{r['Código']}", np.nan) for _, r in df_calc.iterrows()]
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

# ------------------------
# KPIs
# ------------------------
k1, k2, k3, k4 = st.columns(4)
k1.metric("Nível Global", f"{nivel_global:.2f}" if not np.isnan(nivel_global) else "—")
k2.metric("Interpretação", interp)
k3.metric("% Requisitos ≥4", f"{100*(df_calc['Nível']>=4).mean():.1f}%")
k4.metric("% Requisitos ≤2", f"{100*(df_calc['Nível']<=2).mean():.1f}%")
st.markdown("---")

# ------------------------
# Gráficos
# ------------------------
cL, cR = st.columns([2,1])

with cL:
    if len(df_pilar):
        cats = df_pilar["Pilar / Dimensão"].tolist()
        vals = df_pilar["Média"].fillna(0).tolist()
        radar = go.Figure()
        radar.add_trace(go.Scatterpolar(r=vals+[vals[0]], theta=cats+[cats[0]], fill='toself'))
        radar.update_layout(title="Radar — Maturidade por Pilar",
                            polar=dict(radialaxis=dict(visible=True, range=[0,5])),
                            showlegend=False)
        st.plotly_chart(radar, use_container_width=True)

    bar = go.Figure(data=[go.Bar(x=df_pilar["Pilar / Dimensão"], y=df_pilar["Média"],
                                 text=np.round(df_pilar["Média"],2), textposition="auto")])
    bar.update_layout(title="Barras — Média de Maturidade por Pilar", yaxis=dict(range=[0,5]))
    st.plotly_chart(bar, use_container_width=True)

with cR:
    pie = go.Figure(data=[go.Pie(labels=df_pilar["Pilar / Dimensão"], values=df_pilar["Peso"], hole=.35)])
    pie.update_layout(title="Peso Relativo por Pilar")
    st.plotly_chart(pie, use_container_width=True)

st.markdown("---")

# ------------------------
# Tabelas
# ------------------------
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

# ------------------------
# Guardar CSV
# ------------------------
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

