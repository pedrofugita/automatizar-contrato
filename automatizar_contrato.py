# BIBLIOTECAS:
from docx import Document
from datetime import datetime
import pandas as pd
from docx2pdf import convert
import os

# IMPORTA DOCUMENTOS:
tabela = pd.read_excel("Informações.xlsx")
for linha in tabela.index:
    documento = Document("Contrato.docx")

    nome = tabela.loc[linha, "Nome"]
    item1 = tabela.loc[linha, "Item1"]
    item2 = tabela.loc[linha, "Item2"]
    item3 = tabela.loc[linha, "Item3"]

    # DICIONÁRIO:
    referencias = {
        "XXXX": nome,
        "YYYY": item1,
        "ZZZZ": item2,
        "WWWW": item3,
        "DD": str(datetime.now().day),
        "MM": str(datetime.now().month),
        "AAAA": str(datetime.now().year)
    }

    # SUBSTITUI NO DOCUMENTO ORIGINAL:
    for paragrafo in documento.paragraphs:
        for codigo in referencias:
            paragrafo.text = paragrafo.text.replace(codigo, referencias[codigo])

    # SALVA COMO NOVO DOCUMENTO DOCX:
    documento.save(f"ContratosProntos/Contrato - {nome}.docx")

    # CONVERTE WORD PARA PDF:
    convert(f"ContratosProntos/Contrato - {nome}.docx")

# ORGANIZA OS ARQUIVOS:
lista_arquivos = os.listdir("ContratosProntos")
for arquivo in lista_arquivos:
    if ".docx" in arquivo:
        os.rename(f"ContratosProntos/{arquivo}", f"ContratosProntos/Words/{arquivo}")
    if ".pdf" in arquivo:
        os.rename(f"ContratosProntos/{arquivo}", f"ContratosProntos/PDFs/{arquivo}")