import pandas as pd
from docx import Document
from docx.shared import Pt
import locale

def substituir_dados_documento(dados, template_path, output_path):
    document = Document(template_path)
    
    for paragrafo in document.paragraphs:
        for marcador, valor in dados.items():
            if marcador in paragrafo.text:
                for run in paragrafo.runs:
                    if marcador in run.text:
                        if marcador == '<<CIDADE>>':
                            valor_formatado = str(valor).upper()
                        else:
                            valor_formatado = formatar_valor(valor)
                        texto_formatado = run.text.replace(marcador, valor_formatado)
                        run.text = texto_formatado
                        # Manter a formatação do marcador
                        run.font.size = Pt(12)
    
    document.save(output_path)

def formatar_valor(valor):
    locale.setlocale(locale.LC_ALL, '') 
    if isinstance(valor, float):
        valor_formatado = locale.format_string('%.2f', valor, grouping=True)
        return valor_formatado
    else:
        return str(valor)

template = 'CAMINHO PARA O TEMPLATE EM WORD.docx'

df = pd.read_excel('CAMINHO PARA A PLANILHA.xlsx', sheet_name='NOME DA PLANILHA')

for index, row in df.iterrows():
    dados = row.to_dict()

    nome_arquivo = 'COLOCAR O NOME DO ARQUIVO AQUI '+ dados['<<NOME>>'] + '.docx'
    
    output = 'COLOCAR O NOME DA PASTA AQUI' + nome_arquivo

    substituir_dados_documento(dados, template, output)
