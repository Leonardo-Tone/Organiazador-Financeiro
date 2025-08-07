import pandas as pd
import unicodedata
from gspread_formatting import format_cell_range, CellFormat, NumberFormat
from gspread_dataframe import set_with_dataframe
from openpyxl.utils import column_index_from_string, get_column_letter
import re
import os
from dadospessoais import *

def normalizar(texto):
    texto = str(texto).lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto

def tokenize(texto):
    # captura apenas sequências de letras e números
    return set(re.findall(r'\w+', texto))

def tem_match(descricao, nome_alvo):
    desc_norm = normalizar(descricao)
    nome_norm = normalizar(nome_alvo)
    
    # exceções globais
    if any(exc in desc_norm for exc in excecoes):
        return True

    # tokeniza em palavras alfanuméricas
    palavras_desc = tokenize(desc_norm)
    palavras_nome = tokenize(nome_norm)
    if not palavras_nome:
        return False

    # interseção e proporção
    intersec = palavras_desc.intersection(palavras_nome)
    proporcao = len(intersec) / len(palavras_nome)

    # opcional: para nomes completos, exigir substring contínua
    full_match = nome_norm.replace(" ", "")
    desc_junto = desc_norm.replace(" ", "")
    if len(palavras_nome) > 2 and full_match in desc_junto:
        return True

    return proporcao >= 0.75


def marcar_reembolso(descricao):  
    for nome in nomes_reembolsaveis:
        if tem_match(descricao, nome):
            return "Reembolsável"
    return "Pessoal"

def classificar(descricao):
    for categoria, termos in categorias.items():
        if any(termo.lower() in descricao.lower() for termo in termos):
            return categoria

def detalhes(descricao):
    for termo, detalhe in detalhes_map.items():
        if termo.lower() in descricao.lower():
            return detalhe
    return "Sem detalhe"

# Função para extrair apenas a parte desejada da descrição
def extrair_fonte(descricao):
    if pd.isna(descricao):
        return ""
    
    partes = descricao.split("-")
    if len(partes) > 1:
        return partes[1].strip()
    
    match = re.search(r'"([^"]+)"', descricao)
    if match:
        return match.group(1)
    
    return descricao.strip()

# Função para enviar df processado à planilha
def enviar_df_para_planilha(df, coluna_inicial, linha_inicial, sheet):
    # Se estiver vazio, nada a fazer
    if df.empty:
        return

    df = df.copy()
    
    # Seleciona e renomeia as colunas
    df = df[["Valor", "Descrição", "Detalhes", "Data Lançamento", "Categoria"]]
    df.columns = ["Receita/Gasto", "Fonte", "Detalhe", "Data", "Categoria"]

    # Extrai a parte da fonte da forma desejada
    df["Fonte"] = df["Fonte"].apply(extrair_fonte)
    valores = df.values.tolist()

    # Atualiza a planilha a partir da célula inicial
    cell_range = f"{coluna_inicial}{linha_inicial}"
    df["Data"] = pd.to_datetime(
        df["Data"], format="%d/%m/%Y", dayfirst=True, errors="coerce"
    )

    set_with_dataframe(
        worksheet=sheet,
        dataframe=df,
        row=linha_inicial,
        col=column_index_from_string(coluna_inicial),
        include_index=False,
        include_column_header=False,
        resize=False,
        allow_formulas=False
    )


def formatar_planilha(sheet, coluna_inicial, linha_inicial, num_linhas):
    # Se não há linhas, pula formatação
    if num_linhas <= 0:
        return

    # converte "A", "AA" etc para índice e volta para letra
    idx_inicial = column_index_from_string(coluna_inicial)
    idx_final   = idx_inicial + 3
    coluna_valor = coluna_inicial
    coluna_data  = get_column_letter(idx_final)  # Data é a 4ª depois da inicial

    # monta intervalos
    linha_final = linha_inicial + num_linhas - 1
    intervalo_valor = f"{coluna_valor}{linha_inicial}:{coluna_valor}{linha_final}"
    intervalo_data  = f"{coluna_data}{linha_inicial}:{coluna_data}{linha_final}"

    # formatação de número (Receita/Gasto)
    fmt_valor = CellFormat(
        numberFormat=NumberFormat(type="NUMBER", pattern="#,##0.00")
    )
    format_cell_range(sheet, intervalo_valor, fmt_valor)

    # formatação de data
    fmt_data = CellFormat(
        numberFormat=NumberFormat(type="DATE", pattern="DD/MM/YYYY")
    )
    format_cell_range(sheet, intervalo_data, fmt_data)

def obter_ultima_data_preenchida(sheet, coluna_data, linha_inicial):

    # Pega os valores da coluna, ignorando cabeçalho
    valores = sheet.col_values(coluna_data)[linha_inicial - 1:]  # linha 5 → índice 4
    datas = []

    for val in valores:
        try:
            datas.append(pd.to_datetime(val, dayfirst=True, errors="coerce"))
        except:
            continue

    # Remove valores vazios ou inválidos
    datas_validas = [d for d in datas if pd.notna(d)]
    if not datas_validas:
        return None

    return max(datas_validas)

def proxima_linha_vazia(sheet, coluna_letra, linha_inicial):
    # Retorna a próxima linha vazia em uma coluna, começando de linha_inicial.
    col_index = column_index_from_string(coluna_letra)
    valores = sheet.col_values(col_index)[linha_inicial - 1:]  # começa em linha_inicial (ex: 5 → índice 4)
    
    for i, valor in enumerate(valores):
        if not valor.strip():  # encontrou célula vazia
            return linha_inicial + i

    return linha_inicial + len(valores)


def obter_nome_arquivo_csv(mes: str, ano: str, pasta=".") -> str:
    """
    Procura um arquivo de extrato no formato padrão para o mês e ano informados.

    Args:
        mes (str): Nome do mês por extenso com inicial maiúscula (ex: "Abril").
        ano (str): Ano com 4 dígitos (ex: "2025").
        pasta (str): Caminho da pasta onde estão os arquivos (padrão: atual).

    Returns:
        str: Nome do arquivo CSV correspondente.
    """
    mes_para_num = {
        "Janeiro": "01",
        "Fevereiro": "02",
        "Março": "03",
        "Abril": "04",
        "Maio": "05",
        "Junho": "06",
        "Julho": "07",
        "Agosto": "08",
        "Setembro": "09",
        "Outubro": "10",
        "Novembro": "11",
        "Dezembro": "12"
    }

    mes_formatado = mes.capitalize()
    mes_num = mes_para_num.get(mes_formatado)
    if not mes_num:
        raise ValueError(f"Mês inválido: '{mes}'. Use o nome por extenso com inicial maiúscula.")

    padrao = re.compile(rf"Extrato-\d{{2}}-{mes_num}-{ano}-a-\d{{2}}-{mes_num}-{ano}-CSV\.csv")
    arquivos = os.listdir(pasta)
    matches = [f for f in arquivos if padrao.match(f)]

    if not matches:
        raise FileNotFoundError(f"Nenhum arquivo CSV encontrado para {mes}/{ano} na pasta '{pasta}'.")

    if len(matches) > 1:
        raise RuntimeError(f"Mais de um arquivo encontrado para {mes}/{ano}: {matches}")

    return matches[0]

def porquinho(df, incluir=True):
    if incluir:
        return df["Descrição"].str.contains(padrao, case=False, na=False, regex=True)
    else:
        return ~df["Descrição"].str.contains(padrao, case=False, na=False, regex=True)