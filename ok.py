import pandas as pd

try:
    df = pd.read_excel(
        r"C:\Users\rafae\Documents\Trabalho_ Roland_dashboard\Question_Socio.xlsx",
        sheet_name="Sheet1"
    )
    print("Arquivo lido com sucesso!")
    print("Dimensões:", df.shape)
    print("Colunas:", df.columns.tolist())
    print(df.head())
except Exception as e:
    print("Erro ao carregar o arquivo Excel:", e)
