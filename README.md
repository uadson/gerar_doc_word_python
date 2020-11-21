# Gerar e Manipular Arquivos Word com Python I
Script simples para manipular / gerar arquivos do tipo .docx com Python

Neste script tentei elaborar de modo simples com foco em programação procedural, conforme meu nível de conhecimento, um modelo para automatizar pelo menos em parte, a elaboração de documentos.
Em alguns departamentos e principalmente em órgãos públicos, ainda há uso de documentos impressos e a elaboração dos mesmos em muitos se torna uma tarefa repetitiva. Resolvi elaborar um script que pudesse auxiliar de forma rápida, prática e eficiente nessas tarefas que geralmente repetitivas.

O Python é uma linguagem poderosa, com múltiplos recursos e que auxilia de modo eficiente na automatização destas tarefas repetitivas do dia-a-dia.

Optei por utilizar como base de dados uma planilha de excel do tipo .xlsx, aonde os dados serão inseridos de forma manual (para a resolução do meu problema).
Descrição do Script

1. Arquitetura.

Criei uma arquitetura de modo a organizar a distribuição dos arquivos.

Ficou da seguinte forma:

A) Um arquivo .xlsx (db.xlsx) que será utilizado para inserção de dados que está contido em:

B) Um diretório principal (gerar_doc_word_python) que chamei de MAIN_PATH e seus pacotes:

- doc_emitidos: local em que serão armazenados todos os arquivos gerados ao final da execução do script. Eles serão gerados primeiramente no diretório principal, movidos para uma pasta cujo nome será a hora em que os arquivos foram gerados, logo depois serão movidos para outra pasta cujo nome é data em que os arquivos foram gerados e finalmente movidos para pasta doc_emitidos.


	doc_emitidos/
				folder_data_atual/
								 folder_hora_atual/
								 	              arquivos.docx	


- modelos: aqui há uma arquivo .docx como modelo. Explico. Precisei emitir muitos documentos e nestes documentos existe uma logomarca, como o arquivo é em word, isso fica restrito ao cabeçalho e ao rodapé do documento, que são inalteráveis. Por isso optei por um modelo somente com cabeçalho e rodapé.

- templates: nesta pasta existem alguns arquivos .docx, e cada arquivo com um texto diferente. 

E por fim o script que será executado run.py.

Estrutura completa ao final da execução:



gerar_doc_word_python/
				     doc_emitidos/
								folder_data_atual/
								 				folder_hora_atual/
								 	              				 arquivos.docx
modelos/
		modelos.docx

templates/
		 templates.docx

run.py




2. Execução.

A execução do script consiste em:

1º inserção de dados na planilha;

2º na execução, os dados da planilha serão lidos e adicionados em uma lista;

3º os arquivos em templates serão lidos também e adicionados em um dicionário;

4º após a seleção do template e inserção do assunto do documento, 
o arquivo modelo será "aberto" para que o algoritmo inscreva os dados coletados nele e em seguida finaliza o processo, salvando-o.


Espero ajudar de alguma forma. \o


REFERÊNCIA: 
Livro - Automatize Tarefas Maçantes com Python - Programação Prática para Verdadeiros Iniciantes - AL Sweigart


