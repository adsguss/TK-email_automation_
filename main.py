import tkinter as tk
from tkinter import ttk, messagebox
from threading import Thread
import imaplib
import email
import re
import pandas as pd
from email.header import decode_header
import os


# =============================
# FUNÇÃO PRINCIPAL
# =============================
def automacao(email_user, email_pass, log_widget):
    try:
        log_widget.insert(tk.END, "Conectando ao servidor IMAP...\n")
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(email_user, email_pass)
        mail.select("inbox")

        log_widget.insert(tk.END, "Buscando e-mails...\n")

        status, messages = mail.search(None, "ALL")
        email_ids = messages[0].split()

        dados_extraidos = []

        for num in email_ids[-10:]:
            status, msg_data = mail.fetch(num, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])

                    body = ""

                    # =============================
                    # TRATAMENTO DE CHARSET
                    # =============================
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                payload = part.get_payload(decode=True)
                                charset = part.get_content_charset()

                                if payload:
                                    try:
                                        body = payload.decode(charset or "utf-8", errors="ignore")
                                    except:
                                        body = payload.decode("latin-1", errors="ignore")
                    else:
                        payload = msg.get_payload(decode=True)
                        charset = msg.get_content_charset()

                        if payload:
                            try:
                                body = payload.decode(charset or "utf-8", errors="ignore")
                            except:
                                body = payload.decode("latin-1", errors="ignore")

                    # =============================
                    # NOVO REGEX BASEADO NO SEU MODELO
                    # =============================
                    nome = re.search(r"Nome:\s*(.+)", body, re.IGNORECASE)
                    produto = re.search(r"Produto:\s*(.+)", body, re.IGNORECASE)
                    quantidade = re.search(r"Quantidade:\s*(\d+)", body, re.IGNORECASE)
                    valor = re.search(r"Valor:\s*(\d+)", body, re.IGNORECASE)
                    data = re.search(r"Data:\s*(.+)", body, re.IGNORECASE)
                    telefone = re.search(r"Telefone:\s*(\d+)", body, re.IGNORECASE)

                    if nome and produto and quantidade and valor and data and telefone:
                        dados_extraidos.append({
                            "Nome": nome.group(1).strip(),
                            "Produto": produto.group(1).strip(),
                            "Quantidade": quantidade.group(1),
                            "Valor": valor.group(1),
                            "Data": data.group(1).strip(),
                            "Telefone": telefone.group(1)
                        })

        if dados_extraidos:
            log_widget.insert(tk.END, "Gerando planilha Excel...\n")

            df = pd.DataFrame(dados_extraidos)
            caminho = os.path.join(os.getcwd(), "pedidos_extraidos.xlsx")
            df.to_excel(caminho, index=False)

            log_widget.insert(tk.END, f"Arquivo salvo em: {caminho}\n")
        else:
            log_widget.insert(tk.END, "Nenhum pedido encontrado.\n")

        mail.logout()
        log_widget.insert(tk.END, "Processo finalizado!\n")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


# =============================
# FUNÇÃO DO BOTÃO
# =============================
def iniciar_automacao():
    email_user = entry_email.get()
    email_pass = entry_senha.get()

    if not email_user or not email_pass:
        messagebox.showwarning("Atenção", "Preencha e-mail e senha!")
        return

    Thread(
        target=automacao,
        args=(email_user, email_pass, log_area),
        daemon=True
    ).start()


# =============================
# INTERFACE
# =============================
janela = tk.Tk()
janela.title("Sistema de Automação de Pedidos")
janela.geometry("600x500")

titulo = ttk.Label(janela, text="Automação Administrativa", font=("Arial", 16))
titulo.pack(pady=10)

frame_login = ttk.Frame(janela)
frame_login.pack(pady=10)

ttk.Label(frame_login, text="E-mail:").grid(row=0, column=0, padx=5, pady=5)
entry_email = ttk.Entry(frame_login, width=40)
entry_email.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(frame_login, text="Senha/App Password:").grid(row=1, column=0, padx=5, pady=5)
entry_senha = ttk.Entry(frame_login, width=40, show="*")
entry_senha.grid(row=1, column=1, padx=5, pady=5)

btn_iniciar = ttk.Button(janela, text="Iniciar Automação", command=iniciar_automacao)
btn_iniciar.pack(pady=15)

log_area = tk.Text(janela, height=15)
log_area.pack(padx=10, pady=10, fill="both", expand=True)

janela.mainloop()