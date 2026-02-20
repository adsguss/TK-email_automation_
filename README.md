# 📩 Sistema de Automação de Pedidos via E-mail

## 📌 Sobre o Projeto

Este projeto consiste em uma aplicação desktop desenvolvida em Python com o objetivo de automatizar a leitura e processamento de pedidos recebidos por e-mail.

A aplicação conecta-se à caixa de entrada via protocolo IMAP, identifica pedidos estruturados no corpo da mensagem, extrai as informações relevantes e gera automaticamente uma planilha Excel organizada.

O projeto simula um cenário real administrativo/comercial onde pedidos são recebidos por e-mail e precisam ser organizados manualmente.

---

## 🚀 Funcionalidades

- Conexão com servidor de e-mail via IMAP (Gmail)
- Leitura automática da inbox
- Processamento de mensagens multipart (text/plain)
- Tratamento de encoding (UTF-8 / Latin-1)
- Extração de dados utilizando Expressões Regulares (Regex)
- Captura de múltiplos campos:
  - Nome
  - Produto
  - Quantidade
  - Valor
  - Data
  - Telefone
- Estruturação de dados com Pandas (DataFrame)
- Geração automática de relatório em Excel (.xlsx)
- Interface gráfica desenvolvida com Tkinter
- Execução assíncrona com Threading para evitar travamento da interface
- Tratamento de exceções

---

## 🧠 Conceitos Aplicados

- Protocolo IMAP
- Parsing de mensagens MIME
- Pattern Matching com Regex
- Manipulação de dados com Pandas
- Escrita de arquivos Excel com Openpyxl
- Multithreading em aplicações desktop
- Tratamento de encoding e charset
- Estruturação de aplicações com separação de responsabilidades

---

## 🛠️ Tecnologias Utilizadas

- Python 3
- Tkinter
- IMAP (imaplib)
- email (MIME parsing)
- Regex (re)
- Pandas
- Openpyxl
- Threading

---

## 📂 Estrutura do Projeto

automacao-email-pedidos/

│

├── main.py

├── requirements.txt

└── README.md


---

## ⚙️ Como Executar o Projeto

## 1️⃣ Clone o repositório

bash
git clone [https://github.com/seuusuario/automacao-email-pedidos.git](https://github.com/adsguss/TK-email_automation_/tree/main)

----
----

## 2️⃣ Acesse a pasta do projeto
cd automacao-email-pedidos

----

## 3️⃣ Instale as dependências
pip install -r requirements.txt

----

## 4️⃣ Execute o sistema
python main.py

----

# 🔐 Configuração do Gmail (Importante)
Para utilizar a aplicação com Gmail:

Ative a verificação em duas etapas na sua conta Google

Gere uma App Password

Utilize essa App Password no campo de senha da aplicação

---- 

# ⚠️ Nunca utilize sua senha principal do Gmail.

---- 

# 📊 Formato esperado do e-mail

O sistema identifica pedidos com o seguinte padrão no corpo da mensagem:

Nome: Gustavo
Produto: Mouse Gamer
Quantidade: 2
Valor: 150
Data: 19/02/2026
Telefone: 21999999999

----

# 🎯 Objetivo do Projeto

O objetivo foi transformar um processo manual de leitura e digitação de pedidos em uma solução automatizada, reduzindo:

Tempo operacional

Retrabalho

Erros humanos

Atividades repetitivas

----

# 📈 Possíveis Evoluções Futuras

Filtragem apenas de e-mails não lidos

Deploy em ambiente web

Integração com banco de dados

Containerização com Docker

Logs estruturados

Testes automatizados



