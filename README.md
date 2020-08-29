# gerar_doc_word_python
Script simples para manipular / gerar arquivos do tipo .docx com Python

Neste script tentei elaborar de modo simples e, conforme meu nível de conhecimento, um modelo para automatizar pelo menos em parte, a elaboração de documentos.
Em alguns departamentos e principalmente em órgãos públicos, ainda há uso de documentos impressos e a elaboração dos mesmos em muitos se torna uma tarefa repetitiva. Resolvi elaborar um script que pudesse auxiliar de forma rápida, prática e eficiente nessas tarefas que geralmente repetitivas.

O Python é uma linguagem poderosa, com múltiplos recursos e que auxilia de modo eficiente na automatização destas tarefas repetitivas do dia-a-dia.

Optei por utilizar como base de dados uma planilha de excel do tipo .xlsx, aonde os dados serão inseridos de forma manual (para a resolução do meu problema).
Descrição do Script

1. Arquitetura.

Criei uma arquitetura, que no meu ponto de vista, poderia organizar a distribuição dos arquivos.

Ficou da seguinte forma:

- arquivo .xlsx (db.xlsx) que será utilizado para inserção de dados.

Temos um diretório principal (gerar_doc) que chamei de MAIN_PATH e seus pacotes:

- doc_emitidos: local em que serão armazenados os arquivos gerados ao final da execução do script. Eles serão gerados primeiramente no diretório principal e depois movidos para pasta doc_emitidos.

- modelos: aqui há uma arquivo .docx como modelo. Explico. Precisei emitir muitos documentos e nestes documentos existe uma logomarca, como o arquivo é em word, isso fica restrito ao cabeçalho e ao rodapé do documento, que são inalteráveis. Por isso optei por um modelo somente com cabeçalho e rodapé.

- templates: nesta pasta existem alguns arquivos .docx, e cada arquivo com um texto diferente. 

E por fim o script run.py.

2. Execução.

A execução do script consiste em:

a) inserção de dados na planilha;

b) na execução, os dados da planilha serão lidos e adicionados em uma lista;

c) os arquivos em templates serão lidos também e adicionados em um dicionário;

d) após a seleção do template e inserção de alguns dados solicitados como número de documento e assunto do documento, 
o arquivo modelo será "aberto" para que o algoritmo inscreva os dados coletados nele e em seguida finaliza o processo, salvando-o.


Espero ter ajudado. \o


REFERÊNCIA: 
Livro - Automatize Tarefas Maçantes com Python - Programação Prática para Verdadeiros Iniciantes - AL Sweigart


