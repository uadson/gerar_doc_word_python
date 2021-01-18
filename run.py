# 1. Importando as bibliotecas necessárias para execução do projeto
import os
import shutil
from datetime import datetime, date
from time import sleep

from docx import Document
from openpyxl import load_workbook

from templates import *


# 2. Caminho dos diretórios
# 2.1 Planilha que será usada como base de dados
DB = 'db.xlsx'

# 2.2 Arquivo de texto contendo os números que serão inseridos
# nos documentos
arq = 'num.txt'

# 2.3 Caminho do diretório principal
MAIN_PATH = f'{os.getcwd()}'

# 2.4 Caminho dos templates
TEMP_PATH = os.path.join(MAIN_PATH, 'templates')

# 2.5 Caminho dos documentos
DOC_PATH = os.path.join(MAIN_PATH, 'doc_emitidos')

# 3. Formatando a data
dia = date.today().day
mes = date.today().month
ano = date.today().year
data = date.today()
data_atual = data.strftime('%d%m%y')
hora = datetime.now()
hora_atual = hora.strftime('%H%M')

meses = {
	1: 'janeiro',
	2: 'fevereiro',
	3: 'março',	
	4: 'abril',
	5: 'maio',
	6: 'junho',
	7: 'julho',
	8: 'agosto',
	9: 'setembro',
   10: 'outubro',
   11: 'novembro',
   12: 'dezembro'
}

# 4 Criando uma lista com os arquivos do diretório
files = os.listdir(DOC_PATH)
# variável que receberá um valor de um elemento pertecente a lista files
file = f'{ano}' in files
# verificação da existência desse elemento na lista files
if file == True:
	new_folder = file
else:
	new_folder = os.mkdir(f'{DOC_PATH}' + f'\\{ano}')

# 4.1 Criando caminho da nova pasta
FOLDER_PATH = os.path.join(DOC_PATH, f'{ano}')

files = os.listdir(FOLDER_PATH)
file = f'{ano}{mes}' in files
if file == True:
	new_folder = file
else:
	new_folder = os.mkdir(f'{FOLDER_PATH}' + f'\\{ano}{mes}')

FINAL_PATH = os.path.join(FOLDER_PATH, f'{ano}{mes}')

# 5. Tipos de Documentos
tipo_doc = {
	1: 'TIPO DE DOC 1',
	2: 'TIPO DE DOC 2'
}

# 6. Lendo os dados da planilha(coluna e celula)
planilha = load_workbook(DB)
plan = planilha['processos']

# 7. Criando uma lista com os dados coletados da planilha
protocolo = []
i = 0 
for celula in plan['A'][1:15]:
    if celula.value != '':
        num = int(celula.value)
        protocolo.append(num)
        i += 1

# 8. Obtendo o valor total de itens adicionados à lista
total = len(protocolo)


# 9. Obtendo uma lista dos documentos presentes no diretório de templates cujo o inicio 
# do nome de cada arquivo seja 'doc' e adicionando a um dicionário.
lista = os.listdir(TEMP_PATH)
docs = {}
k = 1
for temp in lista:
	if temp.startswith('doc'):
		docs[k] = temp
	k += 1
		

# 10. Selecionando o template que será utilizado para gerar os documentos.
# imprimindo as opções na tela
for k, v in docs.items():
	print(f'[{k}] - {v}')


while True:
	try:
		template = int(input('\nSelecione um documento: '))
	except ValueError:
		print('Digite um número válido!')
	else:
		if template > len(docs) or template <= 0:
			print('Documento inexistente.')
		else:
			break


# 11. Obtendo uma lista com as linhas do texto que será inserido no documento
doc2 = Document(f'{os.getcwd()}' + f'\\templates\\{docs.get(template)}')    
fullText = []
for linha in doc2.paragraphs:
    fullText.append(linha.text)
texto = '\n'.join(fullText)


# 12. Selecionando o tipo de documento
for k, v in tipo_doc.items():
	print(f'[{k}] - {v}')


while True:	
	try:
		tipoDoc = int(input('\nSelecione o tipo de documento: '))
	except ValueError:
		print('Digite um número válido!')
	else:
		if tipoDoc > len(tipo_doc) or tipoDoc <= 0:
			print('Documento inexistente.')
		else:
			break


# 13. Inserindo um número ao documento
# abrindo arquivo em modo leitura
file = open(arq, 'r')
# uma variável receberá cada linha do arquivo convertida para o tipo inteiro
for f in file:
	a = int(f) + 1
# fechando o arquivo	
file.close()


# 13.1 Informando um assunto para o documento			
assunto = input('Informe o Assunto: ').upper().strip()


# 14. Gerando documentos 
for i in range(total):
	# abrindo arquivo em modo r+ para que o conteúdo seja inserido
	# sempre na primeira linha 
	file = open(arq, 'r+')
	# indicando que o registro será feito na primeira linha
	file.seek(0)
	# escrevendo no arquivo
	file.write(str(a))
	# acessando um arquivo modelo para a partir dele criar outros documentos
	doc1 = Document(f'{os.getcwd()}' + '\\modelos\\modelo.docx')
	# adiciona o primeiro parágrafo
	doc1.add_paragraph(f'DOCUMENTO  N. :  {a} / {ano}')
	# adiciona uma linha branco
	doc1.add_paragraph('')
	# adicion o terceiro parágrafo
	doc1.add_paragraph(f'ASSUNTO               :  {assunto}')
	# adiciona uma linha em branco
	doc1.add_paragraph('')
	# adiciona o quinto parágrafo
	doc1.add_paragraph(f'PROTOCOLO        :  {protocolo[i]}')
	# formatação dos dois primeiros parágrafos escritos
	doc1.paragraphs[1].runs[0].bold = True
	doc1.paragraphs[3].runs[0].bold = True
	doc1.paragraphs[5].runs[0].bold = True
	# adiciona duas linhas em branco
	doc1.add_paragraph('')
	doc1.add_paragraph('')
	# adicionando o texto
	doc1.add_paragraph(texto)
	# adiciona duas linhas em branco
	doc1.add_paragraph('')
	doc1.add_paragraph('')
	# adicionando data
	doc1.add_paragraph('>>> Aqui pode ser a identificação do órgão/departamento, etc.')
	doc1.add_paragraph(f'>>> Local, ao(s) {dia} dia(s) do mês de {meses.get(mes)} de {ano}.')
	# adicionando 5 linhas em branco
	doc1.add_paragraph('')
	doc1.add_paragraph('')
	doc1.add_paragraph('')
	doc1.add_paragraph('')
	doc1.add_paragraph('')
	# nome do responsável
	doc1.add_paragraph('>>> Nome do Responsável')
	# formatando a linha
	doc1.paragraphs[18].runs[0].bold = True
	doc1.add_paragraph('>>> Cargo que Ocupa')
	doc1.add_paragraph('Inscrição no Conselho, etc')
	# salvando documento
	# esse formato fica a livre escolha
	doc1.save(f'doc - {protocolo[i]}{hora_atual}.docx')
	# incrementação do numéro de documento	
	a += 1

	# fechando o arquivo
	file.close()

# 15.1 Movendo arquivos
docs = os.listdir(MAIN_PATH)
for doc in docs:
	if doc.startswith('doc -'):
		shutil.move(doc, FINAL_PATH)