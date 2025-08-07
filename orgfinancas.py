from functions import *
import pandas as pd 

mes = "Agosto"
ano = "2025"

# Obtém o nome do arquivo automaticamente
nome_arquivo = obter_nome_arquivo_csv(mes, ano)

# Lê o CSV
df = pd.read_csv(
    nome_arquivo,
    sep=';',
    skiprows=4,
)

df["Data Lançamento"] = pd.to_datetime(
        df["Data Lançamento"], format="%d/%m/%Y", dayfirst=True, errors="coerce"
    )

df = df.assign(
    Reembolso=df["Descrição"].apply(marcar_reembolso),
    Categoria=df["Descrição"].apply(classificar),
    Detalhes=df["Descrição"].apply(detalhes),
    Valor=df["Valor"]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .astype(float)
)

# Renda pessoal: valor positivo e NÃO reembolsável
renda_pessoal = df[
    (df["Valor"] > 0) & 
    (df["Reembolso"] == "Pessoal") & 
    porquinho(df, incluir=False)
]

# Despesa pessoal: valor negativo e NÃO reembolsável, ignorando CDB
despesa_pessoal = df[
    (df["Valor"] < 0) & 
    (df["Reembolso"] == "Pessoal") &
    porquinho(df, incluir=False)
]
despesa_pessoal = despesa_pessoal.assign(
    Valor=despesa_pessoal["Valor"].abs())

# Renda a parte: valor positivo e reembolsável
renda_a_parte = df[
    (df["Valor"] > 0) & 
    (df["Reembolso"] == "Reembolsável")
]

# Despesa a parte: valor negativo e reembolsável
despesa_a_parte = df[
    (df["Valor"] < 0) & 
    (df["Reembolso"] == "Reembolsável")
]
despesa_a_parte = despesa_a_parte.assign(
    Valor=despesa_a_parte["Valor"].abs())

print(df.iloc[:, :5])  # Exibe as primeiras 5 colunas do DataFrame para verificação
#quando tiver mais de um arquivo do mesmo mes deve pegar o mais recente