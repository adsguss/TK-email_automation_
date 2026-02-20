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

### 1️⃣ Clone o repositório

```bash
git clone https://github.com/seuusuario/automacao-email-pedidos.git


## 2️⃣ Acesse a pasta do projeto

