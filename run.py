# 1. Importando as bibliotecas necessárias para execução do projeto
from docx import Document
from xlrd import open_workbook
import os
import shutil
from datetime import date
from datetime import datetime
from templates import *
from time import sleep


# 2. Caminho dos diretórios
# 2.1 Planilha que será usada como base de dados
DB = 'db.xlsx'


# 2.2 Caminho dos templates
TEMP_PATH = f'{os.getcwd()}' + '\\templates'


# 2.3 Caminho dos documentos
DOC_PATH = f'{os.getcwd()}' + '\\doc_emitidos'


# 2.4 Caminho do diretório principal
MAIN_PATH = f'{os.getcwd()}'


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


# 4. Tipos de Documentos
tipo_doc = {
	1: 'TIPO DE DOC 1',
	2: 'TIPO DE DOC 2'
}


# 5. Lendo os dados da planilha(coluna e celula)
planilha = open_workbook(DB)
plan = planilha.sheet_by_index(0)
col_valor = plan.col_values(0)


# 6. Criando uma lista com os dados coletados da planilha
protocolo = []
i = 0 
for celula in col_valor[1:]:
    if celula != '':
        num = int(celula)
        protocolo.append(num)
        i += 1


# 7. Obtendo o valor total de itens adicionados à lista
total = len(protocolo)


# 8. Obtendo uma lista dos documentos presentes no diretório de templates cujo o inicio 
# do nome de cada arquivo seja 'doc' e adicionando a um dicionário.
lista = os.listdir(TEMP_PATH)
docs = {}
k = 1
for temp in lista:
	if temp.startswith('doc'):
		docs[k] = temp
	k += 1
		

# 9. Selecionando o template que será utilizado para gerar os documentos.
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


# 10. Obtendo uma lista com as linhas do texto que será inserido no documento
doc2 = Document(f'{os.getcwd()}' + f'\\templates\\{docs.get(template)}')    
fullText = []
for linha in doc2.paragraphs:
    fullText.append(linha.text)
texto = '\n'.join(fullText)


# 11. Selecionando o tipo de documento
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


# 12. Solicitando o número do documento ao usuário
while True:
	try:
		num_doc = int(input('Informe o número do documento: '))
	except ValueError:
		print('\nDigite um número válido!')
	else:
		if num_doc:
			break

			
assunto = input('Informe o Assunto: ').upper().strip()


# 13. Gerando documentos 
for i in range(total):
	# acessando um arquivo modelo para a partir dele criar outros documentos
	doc1 = Document(f'{os.getcwd()}' + '\\modelos\\modelo.docx')
	# adiciona o primeiro parágrafo
	doc1.add_paragraph(f'DOCUMENTO  N. :  {num_doc} / {ano}')
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
	num_doc += 1


# 14. Aguardando 2 segundos para início de processamento
sleep(2)
# aqui optei por colocar dois segundos de espera pois, os arquivos
# são gerados primeiramente para o diretório principal e depois
# escolhir movê-los para outra pasta para que o diretório principal
# não ficasse cheio de arquivos.


# 15. Movendo os documentos gerados para pasta doc_emitidos
docs = os.listdir(MAIN_PATH)
for doc in docs:
	if doc.startswith('doc'):
		shutil.move(doc, DOC_PATH)


# OBS.: caso tem mais de um modelo de documento, pode-se usar 
# o mesmo procedimento usado para obter a lista de templates

''' Exemplo:

lista = os.listdir(MODELOS_PATH)
docs = {}
k = 1
for modelo in lista:
	if modelo.startswith('modelo'):
		docs[k] = modelo
	k += 1
'''

# E depois listar os modelos do diretório para escolha

''' Exemplo:

for k, v in docs.items():
	print(f'[{k}] - {v}')

modelos = int(input('\nSelecione um modelo: '))
'''
	