import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Performance Motoristas | J&T",
    page_icon="📦",
    layout="wide"
)

# ─── Estilo ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #0f0f0f; }
    .block-container { padding-top: 1.5rem; }
    h1 { color: #D32F2F; font-weight: 700; }
    h3 { color: #EEEEEE; }
    .stDataFrame { border-radius: 8px; }
    .metric-card {
        background: #1e1e1e;
        border: 1px solid #2d2d2d;
        border-radius: 10px;
        padding: 16px 20px;
        text-align: center;
    }
    .badge-ok   { background:#1b5e20; color:#a5d6a7; padding:3px 10px; border-radius:12px; font-size:0.85em; font-weight:600; }
    .badge-fail { background:#b71c1c; color:#ef9a9a; padding:3px 10px; border-radius:12px; font-size:0.85em; font-weight:600; }
    .badge-na   { background:#37474f; color:#b0bec5; padding:3px 10px; border-radius:12px; font-size:0.85em; font-weight:600; }
</style>
""", unsafe_allow_html=True)

# ─── Cabeçalho ────────────────────────────────────────────────────────────────
col_logo, col_title = st.columns([1, 9])
with col_logo:
    st.markdown("## 📦")
with col_title:
    st.markdown("# Performance de Motoristas")
    st.caption("Análise de critérios de elegibilidade | J&T Express – Belém/Ananindeua")

st.divider()

# ─── Upload ────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "📂 Carregar base (Carta de Porte de Entrega)",
    type=["xlsx"],
    help="Exporte do JMS: Carta de Porte de Entrega"
)

# ─── Parâmetros (sidebar) ─────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Parâmetros dos Critérios")
    min_pkgs_day  = st.number_input("📦 Mín. pacotes por dia",   min_value=1,  value=25, step=1)
    min_days      = st.number_input("📅 Mín. dias entregues",    min_value=1,  value=24, step=1)
    min_sundays   = st.number_input("☀️ Mín. domingos entregados", min_value=1, value=3,  step=1)
    st.divider()
    st.markdown("**Legenda**")
    st.markdown("🟢 Critério atingido  \n🔴 Critério não atingido")
    st.divider()
    st.caption("Base pode estar incompleta — resultados refletem os dados disponíveis.")


# ─── Processamento ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Processando base...")
def processar(file_bytes, min_pkg, min_d, min_sun):
    df = pd.read_excel(io.BytesIO(file_bytes))

    df['dt_entrega'] = pd.to_datetime(df['Horário da entrega'], errors='coerce')
    df = df.dropna(subset=['dt_entrega', 'Responsável pela entrega'])
    df['data']       = df['dt_entrega'].dt.date
    df['dia_semana'] = df['dt_entrega'].dt.dayofweek   # 6 = domingo

    motoristas = []
    for motorista, grp in df.groupby('Responsável pela entrega'):
        # pacotes por dia
        pkgs_dia = grp.groupby('data').size()
        min_dia  = int(pkgs_dia.min())
        total_pkgs = int(pkgs_dia.sum())
        c1 = bool((pkgs_dia >= min_pkg).all())

        # dias trabalhados
        dias_trabalhados = int(pkgs_dia.shape[0])
        c2 = dias_trabalhados >= min_d

        # domingos
        domingos = grp[grp['dia_semana'] == 6]['data'].nunique()
        c3 = domingos >= min_sun

        elegivel = c1 and c2 and c3

        motoristas.append({
            'Motorista':            motorista,
            'Total Pacotes':        total_pkgs,
            'Dias Trabalhados':     dias_trabalhados,
            'Mín. Pacotes/Dia':     min_dia,
            f'≥{min_pkg} pkg/dia':  c1,
            f'≥{min_d} dias':       c2,
            f'≥{min_sun} domingos': c3,
            'Domingos Entregues':   int(domingos),
            'ELEGÍVEL':             elegivel,
        })

    result = pd.DataFrame(motoristas).sort_values('Total Pacotes', ascending=False)
    return result, df

# ─── Gerador de Excel ──────────────────────────────────────────────────────────
def gerar_excel(df_result: pd.DataFrame, min_pkg, min_d, min_sun) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório Motoristas"

    RED      = "D32F2F"
    DARK     = "1a1a1a"
    MID      = "2d2d2d"
    LIGHT    = "3d3d3d"
    GREEN_OK = "1b5e20"
    RED_FAIL = "7f0000"
    WHITE    = "FFFFFF"
    YELLOW   = "F9A825"

    thin = Side(style='thin', color="555555")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Título
    ws.merge_cells("A1:J1")
    ws["A1"] = "J&T Express – Relatório de Performance de Motoristas"
    ws["A1"].font      = Font(name="Arial", bold=True, size=14, color=WHITE)
    ws["A1"].fill      = PatternFill("solid", fgColor=RED)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Subtítulo / critérios
    ws.merge_cells("A2:J2")
    ws["A2"] = (f"Critérios: ≥{min_pkg} pkg em todos os dias  |  "
                f"≥{min_d} dias no mês  |  ≥{min_sun} domingos")
    ws["A2"].font      = Font(name="Arial", italic=True, size=10, color="CCCCCC")
    ws["A2"].fill      = PatternFill("solid", fgColor=DARK)
    ws["A2"].alignment = Alignment(horizontal="center")

    ws.append([])  # linha 3 vazia

    # Cabeçalhos
    headers = [
        "Motorista", "Total Pacotes", "Dias Trabalhados",
        "Mín. Pacotes/Dia",
        f"≥{min_pkg} pkg/dia", f"≥{min_d} dias", f"≥{min_sun} domingos",
        "Domingos Entregues", "ELEGÍVEL"
    ]
    ws.append(headers)
    hdr_row = 4
    for col_idx, hdr in enumerate(headers, 1):
        cell = ws.cell(row=hdr_row, column=col_idx)
        cell.font      = Font(name="Arial", bold=True, color=WHITE, size=10)
        cell.fill      = PatternFill("solid", fgColor=MID)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border    = border
    ws.row_dimensions[hdr_row].height = 28

    # Dados
    bool_cols = {
        f"≥{min_pkg} pkg/dia", f"≥{min_d} dias",
        f"≥{min_sun} domingos", "ELEGÍVEL"
    }
    col_map = {h: i+1 for i, h in enumerate(headers)}

    for r_idx, row in enumerate(df_result.itertuples(index=False), start=5):
        elegivel = row[-1]  # última coluna = ELEGÍVEL
        row_bg   = "1e3a1e" if elegivel else "2d1a1a"

        for c_idx, (hdr, val) in enumerate(zip(headers, row), start=1):
            cell       = ws.cell(row=r_idx, column=c_idx)
            cell.border = border
            cell.font   = Font(name="Arial", size=9, color=WHITE)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill  = PatternFill("solid", fgColor=row_bg)

            if hdr in bool_cols:
                if val is True:
                    cell.value = "✔ SIM"
                    cell.font  = Font(name="Arial", size=9, bold=True, color="81C784")
                else:
                    cell.value = "✘ NÃO"
                    cell.font  = Font(name="Arial", size=9, bold=True, color="EF9A9A")
            else:
                cell.value = val

            if hdr == "Motorista":
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Larguras
    col_widths = [38, 14, 17, 17, 16, 14, 18, 18, 12]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Resumo
    total        = len(df_result)
    total_eleg   = int(df_result["ELEGÍVEL"].sum())
    summary_row  = ws.max_row + 2
    ws.cell(row=summary_row, column=1).value = "RESUMO"
    ws.cell(row=summary_row, column=1).font  = Font(name="Arial", bold=True, color=YELLOW, size=10)

    labels = [
        ("Total de Motoristas",    total),
        ("Elegíveis",              total_eleg),
        ("Não Elegíveis",          total - total_eleg),
    ]
    for offset, (label, val) in enumerate(labels):
        r = summary_row + offset + 1
        ws.cell(row=r, column=1).value = label
        ws.cell(row=r, column=1).font  = Font(name="Arial", size=9, color="CCCCCC")
        ws.cell(row=r, column=2).value = val
        ws.cell(row=r, column=2).font  = Font(name="Arial", bold=True, size=9, color=WHITE)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── Main ──────────────────────────────────────────────────────────────────────
if uploaded:
    file_bytes = uploaded.read()
    df_result, df_raw = processar(
        file_bytes, min_pkgs_day, min_days, min_sundays
    )

    # KPIs
    total       = len(df_result)
    elegiveis   = int(df_result["ELEGÍVEL"].sum())
    nao_eleg    = total - elegiveis

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.metric("🚴 Total Motoristas",  total)
    with k2:
        st.metric("✅ Elegíveis",         elegiveis)
    with k3:
        st.metric("❌ Não Elegíveis",     nao_eleg)
    with k4:
        dias_na_base = df_raw['data'].nunique()
        st.metric("📅 Dias na base",      dias_na_base)

    st.divider()

    # Filtro rápido
    filtro = st.radio(
        "Filtrar por:",
        ["Todos", "Apenas Elegíveis", "Apenas Não Elegíveis"],
        horizontal=True
    )
    if filtro == "Apenas Elegíveis":
        df_show = df_result[df_result["ELEGÍVEL"] == True]
    elif filtro == "Apenas Não Elegíveis":
        df_show = df_result[df_result["ELEGÍVEL"] == False]
    else:
        df_show = df_result

    # Tabela com cores
    c1_col = f"≥{min_pkgs_day} pkg/dia"
    c2_col = f"≥{min_days} dias"
    c3_col = f"≥{min_sundays} domingos"

    def color_bool(val):
        if val is True:
            return "background-color:#1b5e20; color:#a5d6a7; font-weight:600"
        elif val is False:
            return "background-color:#7f0000; color:#ef9a9a; font-weight:600"
        return ""

    def color_row(row):
        base = "background-color:#1e3a1e" if row["ELEGÍVEL"] else "background-color:#2d1a1a"
        return [base] * len(row)

    bool_cols_list = [c1_col, c2_col, c3_col, "ELEGÍVEL"]

    styled = (
        df_show.style
        .apply(color_row, axis=1)
        .applymap(color_bool, subset=bool_cols_list)
        .format({
            c1_col:    lambda v: "✔ SIM" if v else "✘ NÃO",
            c2_col:    lambda v: "✔ SIM" if v else "✘ NÃO",
            c3_col:    lambda v: "✔ SIM" if v else "✘ NÃO",
            "ELEGÍVEL":lambda v: "✔ ELEGÍVEL" if v else "✘ NÃO ELEGÍVEL",
        })
    )

    st.dataframe(styled, use_container_width=True, height=560)

    # Detalhamento por motorista
    with st.expander("🔍 Ver detalhamento por motorista (pacotes/dia)"):
        df_raw2 = pd.read_excel(io.BytesIO(file_bytes))
        df_raw2['dt_entrega'] = pd.to_datetime(df_raw2['Horário da entrega'], errors='coerce')
        df_raw2 = df_raw2.dropna(subset=['dt_entrega', 'Responsável pela entrega'])
        df_raw2['data'] = df_raw2['dt_entrega'].dt.date

        motorista_sel = st.selectbox(
            "Selecione o motorista:",
            sorted(df_raw2['Responsável pela entrega'].unique())
        )
        grp_sel = (
            df_raw2[df_raw2['Responsável pela entrega'] == motorista_sel]
            .groupby('data')
            .size()
            .reset_index(name='Pacotes')
        )
        grp_sel['Dia Semana'] = pd.to_datetime(grp_sel['data']).dt.day_name()
        grp_sel['Status'] = grp_sel['Pacotes'].apply(
            lambda x: "✔" if x >= min_pkgs_day else "✘"
        )
        st.dataframe(grp_sel, use_container_width=True)

    st.divider()

    # Download Excel
    excel_bytes = gerar_excel(df_result, min_pkgs_day, min_days, min_sundays)
    st.download_button(
        label="📥 Baixar Relatório Excel",
        data=excel_bytes,
        file_name="relatorio_performance_motoristas_jt.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )

else:
    st.info("👆 Faça o upload da base exportada do JMS para iniciar a análise.")
    st.markdown("""
    **Critérios analisados (configuráveis na barra lateral):**
    - 📦 **Mín. pacotes/dia** — motorista deve ter entregue pelo menos 25 pacotes em *todos* os dias que trabalhou
    - 📅 **Mín. dias no mês** — motorista deve ter entregado em pelo menos 24 dias distintos
    - ☀️ **Mín. domingos** — motorista deve ter entregado em pelo menos 3 domingos distintos
    """)