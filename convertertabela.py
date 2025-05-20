import pandas as pd
from docx import Document

def lerarquivoxlxs(nomearquivo):
    # Leitura do arquivo .xlsx
    df = pd.read_excel(nomearquivo, engine='openpyxl')  
    return df

def formatObjectID(objectid, format="P0000"):
    objectid = str(objectid)
    tamanho = len(format) - len(objectid)
    formatado = format[:tamanho] + objectid
    return formatado

def formatDirection(direction):
    #retornar xxº xx' xx''
    graus, minutos, segundos = direction.split("-")
    formatado = f"{graus}º {minutos}' {segundos}''"
    return formatado

def formatDistance(distance):
    return distance



df = lerarquivoxlxs("Vertices.xlsx")
print(df.head())
print(df.tail())

inicio = 0
fim = 10

# Cria o documento Word
doc = Document()
doc.add_heading("Últimas Linhas da Tabela", level=1)

for n in range(inicio,fim):
    print(formatObjectID(df.loc[n, "OBJECTID_1"]))
    print(formatDirection(df.loc[n, "Direction"]))
    print(formatDistance(df.loc[n, "Distance"]))


# ultimas_linhas = df.tail().values
# # Adiciona cada linha como parágrafo
# for linha in ultimas_linhas:
#     doc.add_paragraph(", ".join(map(str, linha)))

# Salva o documento
# doc.save("saida.docx")

#produção de texto