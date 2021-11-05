## Gerar e Manipular Arquivos Word com Python - 08/2020
[![Build Status](https://app.travis-ci.com/uadson/gerar_doc_word_python.svg?branch=master)](https://app.travis-ci.com/uadson/gerar_doc_word_python)	[![Updates](https://pyup.io/repos/github/uadson/gerar_doc_word_python/shield.svg)](https://pyup.io/repos/github/uadson/gerar_doc_word_python/)	[![Python 3](https://pyup.io/repos/github/uadson/gerar_doc_word_python/python-3-shield.svg)](https://pyup.io/repos/github/uadson/gerar_doc_word_python/)

#### Objetivo:
	
Desenvolver um programa que ajude a reduzir o tempo empreendido em uma atividade repetitiva de elaboração de documentos, em que somente dados como data e número do documento são atualizados, enquanto que o seu contexto permanece o mesmo.

#### Habilidades Python: 

Listas, Dicionários, Tratamento de Excessões, Manipulação de Strings, Loops For e While e Estruturas de Decisão (if/elif/else)


#### Módulos / Bibliotecas

os, datetime, openpyxl, python-docx


#### Estrutura:


###### templates

Diretório no qual ficará os documentos (formato .docx) somente com o contexto inserido, sem cabeçalho ou rodapé, ou formatação.


###### modelos

Diretório com um documento (modelo) no formato .docx, que conterá um cabeçalho e rodapé. O contexto dos templates será inserido neste documento.


###### doc_emitidos

Diretório em que os documentos serão salvos depois de gerados pelo script


##### doc_duplicados

Diretório para o qual serão enviados arquivos repetidos.


#### Tarefas:


##### 1 - Organizar os documentos em diretórios

ao executar o script, diretórios fazendo referência ao ano, mês e dia atuais, são criados com o objetivo de organizar cada documento criado no diretório correspondente a data de sua emissão.


##### 2 - Coletando dados

os dados serão coletados de uma planilha e inseridos em uma lista.


##### 3 - Identificando os templates

cada template tem um contexto diferente, então para identificá-los, no nome de cada arquivo terá
um prefixo 'doc'. O algoritimo fará a leitura dos arquivos cujo o nome inicia com o prefixo e os inserirá em um dicionário que será visualizado na tela.


##### 4 - Coleta do contexto do template

após a identificação do template, o algoritimo coletará todo o contexto do mesmo que será armazenado em uma variável.


##### 5 - Adicionando um número e um assunto ao documento.

um número será solicitado para o que os documentos possuam uma sequência númerica, assim também como um assunto.


##### 6 - Gerando os documentos

os documentos serão gerados contendo todos os dados coletados e salvos em uma pasta correspondente a data de sua emissão.
