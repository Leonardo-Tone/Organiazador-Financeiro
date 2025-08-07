import gspread
from oauth2client.service_account import ServiceAccountCredentials
from orgfinancas import renda_pessoal, despesa_pessoal, renda_a_parte, despesa_a_parte, nome_arquivo
import pandas as pd
from functions import enviar_df_para_planilha, formatar_planilha, obter_ultima_data_preenchida, proxima_linha_vazia
from openpyxl.utils import column_index_from_string
import re
from dadospessoais import id_da_planilha
# Setup Google Sheets API
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'service_account.json', scope
)
mes_para_aba = {
    "01": "Jan.",
    "02": "Fev.",
    "03": "Mar.",
    "04": "Abr.",
    "05": "Mai.",
    "06": "Jun.",
    "07": "Jul.",
    "08": "Ago.",
    "09": "Set.",
    "10": "Out.",
    "11": "Nov.",
    "12": "Dez."
}
match = re.search(r"(\d{2})-(\d{2})-(\d{4})", nome_arquivo)
if not match:
    raise ValueError("Não foi possível extrair a data do nome do arquivo.")

_, mes_str, _ = match.groups()
aba_nome = mes_para_aba.get(mes_str)

if aba_nome is None:
    raise ValueError(f"Mês '{mes_str}' não reconhecido.")

client = gspread.authorize(creds)
spreadsheet = client.open_by_key(id_da_planilha)
sheet = spreadsheet.worksheet(aba_nome)

coluna_data_por_categoria = {
    "A": "D",
    "K": "N",
    "AA": "AD",
    "AK": "AN"
}

# Dicionário de dados e colunas associadas
conjuntos = [
    (renda_pessoal, "A"),
    (despesa_pessoal, "K"),
    (renda_a_parte, "AA"),
    (despesa_a_parte, "AK")
]

for df, col in conjuntos:
    if df.empty:
        continue

    # obter última data da aba naquela categoria
    col_data = coluna_data_por_categoria[col]
    ult_data = obter_ultima_data_preenchida(sheet, column_index_from_string(col_data), linha_inicial=5)

    if ult_data is not None:
        # filtra o df para datas posteriores
        df = df[df["Data Lançamento"] > ult_data]

    if df.empty:
        continue

    # Encontra a próxima linha vazia
    linha_destino = proxima_linha_vazia(sheet, col, linha_inicial=5)

    # Envia os dados a partir dessa linha
    enviar_df_para_planilha(df, col, linha_destino, sheet)
    formatar_planilha(sheet, col, linha_destino, len(df))


