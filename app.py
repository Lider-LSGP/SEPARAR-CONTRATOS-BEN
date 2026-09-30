"""
╔══════════════════════════════════════════════════════════════════════╗
║          LIDER LIMPE — Separador de Relatório VA por Contrato        ║
║                                                                      ║
║  Fluxo:                                                              ║
║   1. Carrega a planilha enviada pelo usuário (2 formatos aceitos,    ║
║      com auto-detecção):                                             ║
║        • "Relatório VA" clássico (separadores por posto + subtotais) ║
║        • "Extrato Mapa Benefícios" (VA+CF+CB consolidado)            ║
║   2. Carrega o Mapeamento de Postos (Google Sheets, repo ou upload)  ║
║   3. Cruza cada POSTO com o mapeamento (Nome_EasyApp → CONTRATO)     ║
║   4. Gera uma planilha .xlsx por contrato (NOME | CPF | VALOR)       ║
║      — o arquivo SEM_CONTRATO inclui a coluna POSTO para facilitar   ║
║        o cadastro dos postos faltantes                               ║
║   5. Empacota tudo num .ZIP no formato:                              ║
║          "CONTRATO      MM-AAAA.xlsx"                                ║
╚══════════════════════════════════════════════════════════════════════╝
"""

from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from core import (
    carregar_mapeamento,
    carregar_mapeamento_google_sheets,
    parse_relatorio_va,
    aplicar_mapeamento,
    build_zip,
    COR_AZUL,
    COR_LARANJA,
    COR_CINZA,
)

APP_DIR = Path(__file__).parent
LOGO_PATH = APP_DIR / "assets" / "logo.png"
DEFAULT_MAP_PATH = APP_DIR / "data" / "Mapeamento Sistema.xls"

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}


def fmt_brl(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


@st.cache_data(ttl=3600, show_spinner=False)
def _carregar_mapeamento_gs_cached():
    """Cache de 1h do mapeamento vindo do Google Sheets.

    Para forçar atualização imediata após editar a planilha:
    menu ⋮ do app → 'Clear cache' → Rerun.
    """
    return carregar_mapeamento_google_sheets()


# ════════════════════════════════════════════════════════════════════════
# UI Streamlit
# ════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="LIDER LIMPE — Separador de VA por Contrato",
    page_icon=str(LOGO_PATH) if LOGO_PATH.exists() else "🧹",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    f"""
    <style>
    .block-container {{ padding-top: 1.6rem; padding-bottom: 2rem; }}
    .stButton>button {{
        background-color: #{COR_AZUL};
        color: white;
        border: 0;
        border-radius: 8px;
        padding: 0.5rem 1.1rem;
        font-weight: 600;
    }}
    .stButton>button:hover {{ background-color: #{COR_LARANJA}; color: white; }}
    .stDownloadButton>button {{
        background-color: #{COR_LARANJA};
        color: white;
        border: 0;
        border-radius: 8px;
        padding: 0.6rem 1.2rem;
        font-weight: 700;
    }}
    .stDownloadButton>button:hover {{ background-color: #{COR_AZUL}; color: white; }}
    .header-card {{
        background: linear-gradient(90deg, #{COR_AZUL} 0%, #2453b5 100%);
        padding: 1.2rem 1.4rem;
        border-radius: 14px;
        color: white;
        margin-bottom: 1rem;
    }}
    .header-card h1 {{ color: white !important; margin: 0; font-size: 1.6rem; }}
    .header-card p  {{ color: #e7ecf7; margin: 0.2rem 0 0 0; }}
    .metric-card {{
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-left: 5px solid #{COR_LARANJA};
        padding: 0.9rem 1rem;
        border-radius: 10px;
        height: 100%;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ────────── Sidebar ──────────
with st.sidebar:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), use_container_width=True)
    st.markdown("### Mapeamento de Postos")
    fonte_mapeamento = st.radio(
        "Fonte do mapeamento",
        options=["Google Sheets (automático)", "Arquivo do repositório", "Upload manual"],
        index=0,
        help=(
            "Google Sheets: lê direto da planilha compartilhada — edições feitas "
            "lá refletem aqui em até 1h (ou imediatamente com Clear cache → Rerun).\n"
            "Repositório: usa o data/Mapeamento Sistema.xls versionado no GitHub.\n"
            "Upload manual: envia um .xls/.xlsx só para esta execução."
        ),
    )
    mapeamento_upload = None
    if fonte_mapeamento == "Upload manual":
        mapeamento_upload = st.file_uploader(
            "Envie o Mapeamento Sistema (.xls/.xlsx)",
            type=["xls", "xlsx"],
            key="map_uploader",
        )

    if fonte_mapeamento == "Google Sheets (automático)":
        if st.button("🔄 Atualizar mapeamento agora", use_container_width=True):
            _carregar_mapeamento_gs_cached.clear()
            st.rerun()

    st.markdown("---")
    st.markdown("### Mês/Ano de referência")
    st.caption(
        "Por convenção, o VA é referente ao **mês posterior** ao mês de impressão "
        "do relatório."
    )
    hoje = date.today()
    prox_mes = hoje.month + 1
    prox_ano = hoje.year
    if prox_mes > 12:
        prox_mes, prox_ano = 1, prox_ano + 1
    col_m, col_a = st.columns(2)
    with col_m:
        mes = st.selectbox(
            "Mês",
            options=list(range(1, 13)),
            index=prox_mes - 1,
            format_func=lambda m: f"{m:02d} — {MESES_PT[m]}",
        )
    with col_a:
        ano = st.number_input(
            "Ano",
            min_value=2020, max_value=2099,
            value=prox_ano, step=1, format="%d",
        )

    st.markdown("---")
    st.caption(
        f"Arquivos terão o sufixo **`{mes:02d}-{ano:04d}`** "
        f"(ex.: `SALVAMAR      {mes:02d}-{ano:04d}.xlsx`)."
    )

# ────────── Header ──────────
col_logo, col_title = st.columns([1, 6])
with col_logo:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=110)
with col_title:
    st.markdown(
        """
        <div class="header-card">
            <h1>Separador de Relatório VA por Contrato</h1>
            <p>Envie o Relatório VA ou o Extrato de Benefícios e baixe um ZIP com uma planilha por contrato.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ────────── Upload do Relatório ──────────
st.markdown("#### 1. Envie a planilha")
relatorio_file = st.file_uploader(
    "Arraste aqui o Relatório VA ou o Extrato Mapa Benefícios (.xls ou .xlsx)",
    type=["xls", "xlsx"],
    key="rel_uploader",
)

if not relatorio_file:
    st.info("⬆️ Envie a planilha para começar — o formato é detectado automaticamente.")
    st.stop()

# ────────── Carregar mapeamento ──────────
st.markdown("#### 2. Conferência")

try:
    if fonte_mapeamento == "Google Sheets (automático)":
        try:
            mp = _carregar_mapeamento_gs_cached()
            fonte_map = "Google Sheets (atualização automática)"
        except Exception:
            # Fallback: se o Sheets estiver inacessível, usa o arquivo do repo
            if DEFAULT_MAP_PATH.exists():
                mp = carregar_mapeamento(str(DEFAULT_MAP_PATH))
                fonte_map = "repositório (fallback — Google Sheets indisponível)"
                st.warning(
                    "⚠️ Não consegui ler o Google Sheets agora. "
                    "Usei o arquivo do repositório como contingência."
                )
            else:
                raise
    elif fonte_mapeamento == "Arquivo do repositório":
        if not DEFAULT_MAP_PATH.exists():
            st.error(
                "O arquivo `data/Mapeamento Sistema.xls` não foi encontrado no repositório."
            )
            st.stop()
        mp = carregar_mapeamento(str(DEFAULT_MAP_PATH))
        fonte_map = "repositório (`data/Mapeamento Sistema.xls`)"
    else:  # Upload manual
        if mapeamento_upload is None:
            st.warning("Envie o Mapeamento Sistema no menu lateral.")
            st.stop()
        mp = carregar_mapeamento(mapeamento_upload)
        fonte_map = f"upload manual (`{mapeamento_upload.name}`)"
except Exception as e:
    st.error(f"Falha ao carregar o Mapeamento Sistema: {e}")
    st.stop()

# ────────── Processar Relatório ──────────
try:
    df, meta, fmt_detectado = parse_relatorio_va(relatorio_file)
except Exception as e:
    st.error(f"Falha ao processar a planilha: {e}")
    st.stop()

# Badge do formato detectado
_labels = {
    "relatorio_va": "📄 Relatório VA (formato clássico com separadores por posto)",
    "extrato_beneficios": "📊 Extrato Mapa Benefícios (VA + CF + CB consolidado)",
}
st.info(f"**Formato detectado:** {_labels.get(fmt_detectado, fmt_detectado)}")


if df.empty:
    st.warning("Nenhuma linha de colaborador foi encontrada no relatório.")
    st.stop()

df = aplicar_mapeamento(df, mp)

# ────────── Métricas ──────────
total_colab = len(df)
total_valor = df["VALOR"].sum()
contratos_unicos = sorted(df["CONTRATO"].unique().tolist())
contratos_validos = [c for c in contratos_unicos if c != "SEM_CONTRATO"]
sem_contrato_df = df[df["CONTRATO"] == "SEM_CONTRATO"]

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(
        f'<div class="metric-card"><b>Colaboradores</b><br>'
        f'<span style="font-size:1.6rem;color:#{COR_AZUL};">{total_colab}</span></div>',
        unsafe_allow_html=True,
    )
with c2:
    st.markdown(
        f'<div class="metric-card"><b>Valor total</b><br>'
        f'<span style="font-size:1.6rem;color:#{COR_AZUL};">{fmt_brl(total_valor)}</span></div>',
        unsafe_allow_html=True,
    )
with c3:
    st.markdown(
        f'<div class="metric-card"><b>Contratos identificados</b><br>'
        f'<span style="font-size:1.6rem;color:#{COR_AZUL};">{len(contratos_validos)}</span></div>',
        unsafe_allow_html=True,
    )
with c4:
    cor_alerta = COR_LARANJA if len(sem_contrato_df) else COR_CINZA
    st.markdown(
        f'<div class="metric-card" style="border-left-color:#{cor_alerta};">'
        f'<b>Sem mapeamento</b><br>'
        f'<span style="font-size:1.6rem;color:#{cor_alerta};">{len(sem_contrato_df)}</span></div>',
        unsafe_allow_html=True,
    )

st.markdown(" ")
st.caption(
    f"Mapeamento carregado de **{fonte_map}** — {len(mp)} postos cadastrados. "
    + (f"Período do relatório: **{meta['periodo']}**." if meta.get("periodo") else "")
)

# ────────── Alerta: postos sem mapeamento ──────────
if not sem_contrato_df.empty:
    postos_orfaos = sorted(sem_contrato_df["POSTO"].dropna().unique().tolist())
    with st.expander(
        f"⚠️ {len(postos_orfaos)} posto(s) sem CONTRATO no Mapeamento "
        f"({len(sem_contrato_df)} colaborador(es) — irão para `SEM_CONTRATO.xlsx` "
        "com a coluna POSTO)",
        expanded=True,
    ):
        st.markdown(
            "Estes postos **não foram encontrados** na coluna `Nome_EasyApp` do "
            "Mapeamento, ou estão cadastrados sem `CONTRATO`. "
            "Eles serão agrupados em um arquivo único `SEM_CONTRATO` **com a coluna "
            "POSTO visível**, para facilitar o cadastro. "
            "Atualize a aba **POSTOS** no Google Sheets e reprocessa o arquivo."
        )
        st.dataframe(
            pd.DataFrame({"Posto não mapeado": postos_orfaos}),
            use_container_width=True, hide_index=True,
        )

# ────────── Resumo por contrato ──────────
resumo = (
    df.groupby("CONTRATO", dropna=False)
      .agg(Colaboradores=("NOME", "count"),
           Valor_Total=("VALOR", "sum"))
      .reset_index()
      .sort_values("CONTRATO")
)
resumo_display = resumo.copy()
resumo_display["Valor_Total"] = resumo_display["Valor_Total"].map(fmt_brl)

st.markdown("#### 3. Resumo por Contrato")
st.dataframe(resumo_display, use_container_width=True, hide_index=True)

with st.expander("👀 Pré-visualizar dados consolidados (após limpeza)"):
    st.dataframe(df, use_container_width=True, hide_index=True, height=320)

# ────────── Gerar ZIP ──────────
st.markdown("#### 4. Gerar arquivos")
incluir_orfaos = st.checkbox(
    "Incluir arquivo `SEM_CONTRATO` no ZIP",
    value=True,
    help="Desligue para gerar apenas os contratos identificados.",
)

if st.button("🗂️  Gerar ZIP com planilhas por contrato", type="primary"):
    grupos = {c: g.reset_index(drop=True) for c, g in df.groupby("CONTRATO")}
    if not incluir_orfaos and "SEM_CONTRATO" in grupos:
        del grupos["SEM_CONTRATO"]

    if not grupos:
        st.error("Nenhum contrato para exportar.")
        st.stop()

    with st.spinner("Gerando arquivos…"):
        zip_bytes = build_zip(grupos, int(mes), int(ano))

    zip_name = f"RELATORIO VA POR CONTRATO {mes:02d}-{ano:04d}.zip"
    st.success(f"✅ {len(grupos)} planilha(s) gerada(s) com sucesso!")
    st.download_button(
        label=f"⬇️  Baixar {zip_name}",
        data=zip_bytes,
        file_name=zip_name,
        mime="application/zip",
    )

st.markdown("---")
st.caption(
    "© LIDER LIMPE — Ferramenta interna. "
    "Mapeamento de postos: edite direto no Google Sheets (aba POSTOS) — "
    "o app atualiza sozinho em até 1h, ou use o botão 🔄 no menu lateral."
)
