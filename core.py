"""
Núcleo de processamento — sem dependência do Streamlit.

Pode ser usado standalone (CLI/scripts) ou pelo app.py.

Fontes de mapeamento suportadas:
  - Google Sheets (planilha compartilhada, atualização automática) — padrão
  - Arquivo local no repositório (data/Mapeamento Sistema.xls) — fallback
  - Upload manual (.xls/.xlsx) — uso pontual

Formatos de planilha de entrada suportados (auto-detectados):
  - "relatorio_va": Relatório VA clássico (separadores 'Posto de Trabalho:',
    subtotais, rodapé longo)
  - "extrato_beneficios": Extrato Mapa Benefícios (VA+CF+CB consolidado,
    sem separadores, apenas 1 linha de TOTAL GERAL no final)
"""

from __future__ import annotations

import io
import re
import zipfile
import unicodedata
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side


MAPEAMENTO_SHEET = "POSTOS"

# Cores da identidade visual LIDER LIMPE
COR_AZUL = "1B3A8C"      # azul escuro do logo
COR_LARANJA = "F47B20"   # laranja do logo
COR_CINZA = "6C7686"     # cinza secundário


# ════════════════════════════════════════════════════════════════════════
# Utilidades
# ════════════════════════════════════════════════════════════════════════
def normalize_text(s) -> str:
    """
    Normaliza string para comparação (sem acento, upper, trim, sem espaços duplos).
    Também unifica os vários tipos de traço Unicode (hífen, travessão, en-dash,
    em-dash, minus...) para o hífen comum, e converte espaços invisíveis
    (NBSP etc.) em espaço normal — evita falhas de comparação quando o texto
    foi editado com autocorreção do Word/Google Sheets.
    """
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s).strip()
    # Unificar todos os tipos de traço no hífen comum "-"
    s = re.sub(r"[\u2010\u2011\u2012\u2013\u2014\u2015\u2212\uFE58\uFE63\uFF0D]", "-", s)
    # Espaços invisíveis (NBSP, thin space, etc.) → espaço normal
    s = re.sub(r"[\u00A0\u2000-\u200A\u202F\u205F\u3000]", " ", s)
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII")
    s = re.sub(r"\s+", " ", s).upper()
    return s


def normalize_cpf(cpf) -> str:
    """Mantém apenas dígitos do CPF e completa com zeros à esquerda até 11."""
    if cpf is None or (isinstance(cpf, float) and pd.isna(cpf)):
        return ""
    digits = re.sub(r"\D", "", str(cpf))
    if not digits:
        return ""
    return digits.zfill(11)


def safe_filename(name: str) -> str:
    """Remove caracteres inválidos para nome de arquivo Windows/Linux."""
    name = str(name).strip()
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name or "SEM_NOME"


def parse_valor(v):
    """Converte VALOR (float/int/str BR) para float. Vazios viram 0.0."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("R$", "").replace(" ", "")
    if not s:
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def read_excel_any(file, sheet_name=0, header=None) -> pd.DataFrame:
    """Lê .xls (xlrd) ou .xlsx (openpyxl) de forma robusta."""
    name = getattr(file, "name", str(file)).lower()
    if name.endswith(".xls"):
        engines = ["xlrd", "openpyxl"]
    elif name.endswith(".xlsx"):
        engines = ["openpyxl", "xlrd"]
    else:
        engines = ["xlrd", "openpyxl"]

    last_err = None
    for eng in engines:
        try:
            if hasattr(file, "seek"):
                file.seek(0)
            return pd.read_excel(file, sheet_name=sheet_name, header=header, engine=eng)
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"Não foi possível ler o arquivo Excel. Último erro: {last_err}")


# ════════════════════════════════════════════════════════════════════════
# Mapeamento de Postos → Contrato
# ════════════════════════════════════════════════════════════════════════
# Fonte principal: planilha Google Sheets compartilhada (atualização automática).
# Precisa estar como "Qualquer pessoa com o link pode ver".
GOOGLE_SHEETS_ID = "1Rn3pYB9uHQysMMfNGgHccJPALDMqf9f4sOYk7MyZaxc"
GOOGLE_SHEETS_URL = (
    f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEETS_ID}"
    f"/gviz/tq?tqx=out:csv&sheet={MAPEAMENTO_SHEET}"
)


def _df_para_dict_mapeamento(df: pd.DataFrame) -> dict:
    """
    Transforma DataFrame do mapeamento em dict:
        { normalize_text(Nome_EasyApp) -> CONTRATO }
    """
    cols_lower = {c.lower().strip(): c for c in df.columns if isinstance(c, str)}
    col_nome = cols_lower.get("nome_easyapp")
    col_contrato = cols_lower.get("contrato")
    if col_nome is None or col_contrato is None:
        raise ValueError(
            "A aba 'POSTOS' do Mapeamento precisa ter as colunas "
            "'Nome_EasyApp' e 'CONTRATO'."
        )

    mp: dict = {}
    for _, row in df.iterrows():
        nome = row[col_nome]
        contrato = row[col_contrato]
        if pd.isna(nome) or pd.isna(contrato):
            continue
        key = normalize_text(nome)
        if not key:
            continue
        # Se um Nome_EasyApp se repete com contratos diferentes, prevalece o primeiro
        if key not in mp:
            mp[key] = str(contrato).strip()
    return mp


def carregar_mapeamento_google_sheets(url: str = GOOGLE_SHEETS_URL) -> dict:
    """
    Lê a aba POSTOS diretamente do Google Sheets (export CSV público).
    Requer que a planilha esteja como 'qualquer pessoa com o link pode ver'.
    """
    df = pd.read_csv(url, dtype=str)
    return _df_para_dict_mapeamento(df)


def carregar_mapeamento(file_or_path) -> dict:
    """
    Lê a aba POSTOS do Mapeamento Sistema (.xls/.xlsx local ou upload)
    e devolve dict: { normalize_text(Nome_EasyApp) -> CONTRATO }
    """
    df = read_excel_any(file_or_path, sheet_name=MAPEAMENTO_SHEET, header=0)
    return _df_para_dict_mapeamento(df)


# ════════════════════════════════════════════════════════════════════════
# Núcleo: parsing das planilhas de entrada
# ════════════════════════════════════════════════════════════════════════
def detect_format(raw: pd.DataFrame) -> str:
    """
    Auto-detecta o formato da planilha de entrada.

    Retorna:
      - "extrato_beneficios": Extrato Mapa Benefícios (VA+CF+CB consolidado,
        colunas: Nome do Funcionário | CPF/CNPJ | Posto de Trabalho | Total Benefício (R$))
      - "relatorio_va": Relatório VA original (com separadores 'Posto de Trabalho:',
        subtotais e rodapé longo)
    """
    for i in range(min(5, len(raw))):
        row_vals = [normalize_text(v) for v in raw.iloc[i].tolist()]
        row_text = " | ".join(row_vals)
        # Extrato: cabeçalho tem "NOME DO FUNCIONARIO" + "BENEFICIO"
        if "NOME DO FUNCIONARIO" in row_text and "BENEFICIO" in row_text:
            return "extrato_beneficios"
        # Relatório VA: NOME + CPF + LOTE juntos
        if "NOME" in row_vals and "CPF" in row_vals and "LOTE" in row_vals:
            return "relatorio_va"
    return "relatorio_va"  # default (mantém compatibilidade)


def parse_extrato_beneficios(raw: pd.DataFrame):
    """
    Parser do 'Extrato Mapa Benefícios' (formato consolidado VA+CF+CB).

    Layout esperado:
      Col 0: Nome do Funcionário
      Col 1: CPF/CNPJ
      Col 2: Posto de Trabalho
      Col 3: Total Benefício (R$)

    Só há 1 linha de rodapé ("TOTAL GERAL:"), sem separadores nem subtotais.
    """
    # Achar linha do cabeçalho
    header_row = 0
    for i in range(min(5, len(raw))):
        row_vals = [normalize_text(v) for v in raw.iloc[i].tolist()]
        if "NOME DO FUNCIONARIO" in row_vals:
            header_row = i
            break

    header_vals = [normalize_text(v) for v in raw.iloc[header_row].tolist()]

    def find_col(*names):
        for nm in names:
            if nm in header_vals:
                return header_vals.index(nm)
        for nm in names:  # busca parcial (fallback)
            for i, h in enumerate(header_vals):
                if nm in h:
                    return i
        return None

    idx_nome = find_col("NOME DO FUNCIONARIO", "NOME")
    idx_cpf = find_col("CPF/CNPJ", "CPF")
    idx_posto = find_col("POSTO DE TRABALHO", "POSTO TRABALHO", "POSTO")
    idx_valor = find_col("TOTAL BENEFICIO (R$)", "TOTAL BENEFICIO", "BENEFICIO", "VALOR")

    if any(v is None for v in (idx_nome, idx_cpf, idx_posto, idx_valor)):
        raise ValueError(
            "Cabeçalho não reconhecido no Extrato de Benefícios. Esperado: "
            "Nome do Funcionário, CPF/CNPJ, Posto de Trabalho, Total Benefício (R$)."
        )

    body = raw.iloc[header_row + 1:].reset_index(drop=True)

    rows = []
    for i in range(len(body)):
        nome_raw = body.iat[i, idx_nome]
        cpf_raw = body.iat[i, idx_cpf]
        posto_raw = body.iat[i, idx_posto]
        valor_raw = body.iat[i, idx_valor]

        # Rodapé "TOTAL GERAL:" — NOME/CPF vazios
        nome_vazio = nome_raw is None or (isinstance(nome_raw, float) and pd.isna(nome_raw))
        cpf_vazio = cpf_raw is None or (isinstance(cpf_raw, float) and pd.isna(cpf_raw))
        if nome_vazio and cpf_vazio:
            continue
        posto_str = str(posto_raw or "").strip().upper()
        if "TOTAL GERAL" in posto_str or "TOTAL/GERAL" in posto_str:
            continue

        rows.append({
            "NOME": str(nome_raw).strip() if not pd.isna(nome_raw) else "",
            "CPF": normalize_cpf(cpf_raw),
            "POSTO": str(posto_raw).strip() if posto_raw is not None and not pd.isna(posto_raw) else "",
            "VALOR": parse_valor(valor_raw),
        })

    df = pd.DataFrame(rows, columns=["NOME", "CPF", "POSTO", "VALOR"])
    return df, {"periodo": None, "total_geral": None, "codigo_mapa": None}


def _parse_relatorio_va_classico(raw: pd.DataFrame):
    """
    Parser do Relatório VA clássico (com separadores 'Posto de Trabalho:',
    subtotais e rodapé longo).
    Recebe o DataFrame bruto já lido, para permitir auto-detecção.
    """

    # 1) Achar linha de cabeçalho com NOME e CPF
    header_row = None
    for i in range(min(15, len(raw))):
        row_vals = [normalize_text(v) for v in raw.iloc[i].tolist()]
        if "NOME" in row_vals and "CPF" in row_vals:
            header_row = i
            break
    if header_row is None:
        header_row = 0

    header_vals = [normalize_text(v) for v in raw.iloc[header_row].tolist()]

    def find_col(*names):
        for nm in names:
            if nm in header_vals:
                return header_vals.index(nm)
        return None

    idx_nome = find_col("NOME")
    idx_cpf = find_col("CPF")
    idx_posto = find_col("POSTO TRABALHO", "POSTO DE TRABALHO", "POSTO")
    idx_lote = find_col("LOTE")
    idx_valor = find_col("VALOR")

    if any(v is None for v in (idx_nome, idx_cpf, idx_posto, idx_valor)):
        raise ValueError(
            "Cabeçalho não reconhecido. Esperado pelo menos: "
            "NOME, CPF, POSTO TRABALHO, VALOR."
        )

    body = raw.iloc[header_row + 1:].copy().reset_index(drop=True)

    # 2) Identificar onde começa o rodapé
    meta = {"periodo": None, "total_geral": None, "codigo_mapa": None}
    footer_idx_candidates = []
    for i in range(len(body)):
        c0 = str(body.iat[i, 0]) if not pd.isna(body.iat[i, 0]) else ""
        c3 = ""
        if idx_lote is not None and idx_lote < body.shape[1]:
            c3 = str(body.iat[i, idx_lote]) if not pd.isna(body.iat[i, idx_lote]) else ""
        low = c0.lower()
        if low.startswith("período") or low.startswith("periodo"):
            meta["periodo"] = c0.split(":", 1)[-1].strip()
            footer_idx_candidates.append(i)
        elif low.startswith("código/mapa") or low.startswith("codigo/mapa"):
            meta["codigo_mapa"] = c0.split(":", 1)[-1].strip()
            footer_idx_candidates.append(i)
        elif low.startswith("empresa:") or low.startswith("relatório de va") or \
             low.startswith("relatorio de va") or low.startswith("data/hora"):
            footer_idx_candidates.append(i)
        elif "total/geral" in c3.lower() or "total geral" in c3.lower():
            footer_idx_candidates.append(i)

    cutoff = min(footer_idx_candidates) if footer_idx_candidates else len(body)

    # 3) Filtrar
    rows = []
    for i in range(cutoff):
        nome_raw = body.iat[i, idx_nome]
        cpf_raw = body.iat[i, idx_cpf]
        posto_raw = body.iat[i, idx_posto]
        valor_raw = body.iat[i, idx_valor]
        lote_raw = body.iat[i, idx_lote] if idx_lote is not None else None

        # separador "Posto de Trabalho:"
        if isinstance(nome_raw, str) and nome_raw.strip().lower().startswith("posto de trabalho"):
            continue
        # subtotal
        if isinstance(lote_raw, str) and "subtotal" in lote_raw.strip().lower():
            continue
        # linha em branco / só obs
        nome_vazio = nome_raw is None or (isinstance(nome_raw, float) and pd.isna(nome_raw))
        cpf_vazio = cpf_raw is None or (isinstance(cpf_raw, float) and pd.isna(cpf_raw))
        if nome_vazio and cpf_vazio:
            continue

        rows.append({
            "NOME": str(nome_raw).strip() if not pd.isna(nome_raw) else "",
            "CPF": normalize_cpf(cpf_raw),
            "POSTO": str(posto_raw).strip() if not pd.isna(posto_raw) else "",
            "VALOR": parse_valor(valor_raw),
        })

    df = pd.DataFrame(rows, columns=["NOME", "CPF", "POSTO", "VALOR"])
    return df, meta


def parse_relatorio_va(file):
    """
    Ponto de entrada público. Lê o arquivo, detecta o formato e chama
    o parser apropriado.

    Retorna (df_limpo, meta, formato_detectado).
      - df_limpo: DataFrame com NOME | CPF | POSTO | VALOR
      - meta: dict com metadados quando disponíveis
      - formato_detectado: "relatorio_va" | "extrato_beneficios"
    """
    raw = read_excel_any(file, sheet_name=0, header=None)
    fmt = detect_format(raw)

    if fmt == "extrato_beneficios":
        df, meta = parse_extrato_beneficios(raw)
    else:
        df, meta = _parse_relatorio_va_classico(raw)

    return df, meta, fmt


def aplicar_mapeamento(df: pd.DataFrame, mp: dict) -> pd.DataFrame:
    """Adiciona coluna CONTRATO; postos não mapeados → 'SEM_CONTRATO'."""
    def lookup(posto):
        return mp.get(normalize_text(posto), "SEM_CONTRATO")
    df = df.copy()
    df["CONTRATO"] = df["POSTO"].map(lookup)
    return df


# ════════════════════════════════════════════════════════════════════════
# Geração das planilhas por contrato
# ════════════════════════════════════════════════════════════════════════
def build_xlsx_contrato(nome_contrato: str, df_contrato: pd.DataFrame) -> bytes:
    """
    Gera .xlsx em memória com estilo + linha de TOTAL.

    Todas as planilhas (contratos normais e SEM_CONTRATO) saem com 5 colunas:
        MATRICULA | NOME | CPF | POSTO | VALOR
    A coluna MATRICULA é gerada em branco (reservada para preenchimento manual).
    """
    headers = ["MATRICULA", "NOME", "CPF", "POSTO", "VALOR"]
    col_valor = 5
    n_cols = len(headers)

    wb = Workbook()
    ws = wb.active
    ws.title = safe_filename(nome_contrato)[:31] or "Planilha1"
    ws.append(headers)

    header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    header_fill = PatternFill("solid", fgColor=COR_AZUL)
    thin = Side(border_style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Estilizar cabeçalho
    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Linhas de dados — MATRICULA sempre em branco
    for _, r in df_contrato.iterrows():
        ws.append(["", r["NOME"], r["CPF"], r.get("POSTO", ""), r["VALOR"]])

    # Formatação das linhas
    last_row = ws.max_row
    for row in range(2, last_row + 1):
        ws.cell(row=row, column=1).number_format = "@"  # MATRICULA como texto
        ws.cell(row=row, column=3).number_format = "@"  # CPF como texto
        ws.cell(row=row, column=col_valor).number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'
        for col in range(1, n_cols + 1):
            ws.cell(row=row, column=col).border = border
            if col == 1:
                align = "center"   # MATRICULA
            elif col == 2:
                align = "left"     # NOME
            elif col == 3:
                align = "center"   # CPF
            elif col == col_valor:
                align = "right"    # VALOR
            else:
                align = "left"     # POSTO
            ws.cell(row=row, column=col).alignment = Alignment(
                horizontal=align, vertical="center",
            )

    # Linha de TOTAL
    total_row = last_row + 1
    label_col = col_valor - 1  # célula imediatamente antes do VALOR (POSTO)
    ws.cell(row=total_row, column=label_col, value="TOTAL").font = Font(bold=True)
    ws.cell(row=total_row, column=label_col).alignment = Alignment(horizontal="right")
    for col in range(1, n_cols + 1):
        ws.cell(row=total_row, column=col).fill = PatternFill("solid", fgColor="F2F2F2")
    total_cell = ws.cell(row=total_row, column=col_valor, value=float(df_contrato["VALOR"].sum()))
    total_cell.font = Font(bold=True, color=COR_AZUL)
    total_cell.number_format = 'R$ #,##0.00'

    # Larguras de coluna
    ws.column_dimensions["A"].width = 14    # MATRICULA
    ws.column_dimensions["B"].width = 45    # NOME
    ws.column_dimensions["C"].width = 16    # CPF
    ws.column_dimensions["D"].width = 40    # POSTO
    ws.column_dimensions["E"].width = 16    # VALOR
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def build_zip(grupos: dict, mes: int, ano: int) -> bytes:
    """
    grupos: { contrato -> DataFrame }
    Gera ZIP com cada contrato em "CONTRATO      MM-AAAA.xlsx".
    """
    buf = io.BytesIO()
    sufixo = f"{mes:02d}-{ano:04d}"
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for contrato, df_c in sorted(grupos.items()):
            xlsx_bytes = build_xlsx_contrato(contrato, df_c)
            fname = f"{safe_filename(contrato)}      {sufixo}.xlsx"
            zf.writestr(fname, xlsx_bytes)
    buf.seek(0)
    return buf.getvalue()
