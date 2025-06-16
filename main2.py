import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.scrolledtext as scrolledtext
import threading
import pandas as pd
import time
import win32com.client as win32
import os
import json

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Prospecção Automática JR Solutions")
        self.root.geometry("1000x700")

        # Cores
        self.cor_fundo = "#c1c4c4"
        self.cor_primaria = "#216eaa"
        self.cor_texto = "#060606"
        self.cor_botao = "#3fc3ed"
        self.cor_botao_disabled = "#8c7c84"

        self.estilo = {
            "botao": {"bg": self.cor_botao, "fg": self.cor_texto, "relief": "flat"},
            "label": {"bg": self.cor_fundo, "fg": self.cor_texto}
        }

        self.root.configure(bg=self.cor_fundo)

        # Variáveis de controle
        self.planilha_path = ""
        self.anexo_path = []
        self.enviando = False
        self.pausado = False
        self.encerrar = False

        # Configurações salvas
        self.email_subject_default = "APRESENTAÇÃO JRSOLUTIONS / OHM CENTRO DE REPAROS LTDA. "
        self.email_body_default = """\n🚩🚩Coloque o corpo do seu e-mail 🚩🚩"""
        self.delay_padrao = 75

        self.load_config()

        self.canvas = tk.Canvas(self.root, bg=self.cor_fundo, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollable_frame = tk.Frame(self.canvas, bg=self.cor_fundo)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.create_widgets()

    def load_config(self):
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                config = json.load(f)
                self.email_subject_default = config.get("titulo", self.email_subject_default)
                self.email_body_default = config.get("corpo", self.email_body_default)
                self.delay_padrao = config.get("delay", self.delay_padrao)
        except:
            pass

    def save_config(self):
        config = {
            "titulo": self.entry_subject.get().strip(),
            "corpo": self.txt_corpo_email.get("1.0", tk.END).strip(),
            "delay": self.delay_padrao
        }
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)

    def create_widgets(self):
        self.btn_planilha = tk.Button(self.scrollable_frame, text="Selecionar Planilha", command=self.selecionar_planilha, **self.estilo["botao"])
        self.btn_planilha.pack(pady=5, padx=10, fill='x')

        self.lbl_planilha = tk.Label(self.scrollable_frame, text="Nenhuma planilha selecionada", **self.estilo["label"])
        self.lbl_planilha.pack(padx=10)

        self.btn_anexo = tk.Button(self.scrollable_frame, text="Selecionar Anexo", command=self.selecionar_anexo, **self.estilo["botao"])
        self.btn_anexo.pack(pady=5, padx=10, fill='x')

        self.lbl_anexo = tk.Label(self.scrollable_frame, text="Nenhum anexo selecionado", justify="left", **self.estilo["label"])
        self.lbl_anexo.pack(padx=10)

        tk.Label(self.scrollable_frame, text="Título do E-mail:", **self.estilo["label"]).pack(pady=(15, 0), padx=10, anchor='w')
        self.entry_subject = tk.Entry(self.scrollable_frame, width=80, bg="white", fg=self.cor_texto, relief="solid")
        self.entry_subject.pack(padx=10, pady=5, fill='x')
        self.entry_subject.insert(0, self.email_subject_default)

        tk.Label(self.scrollable_frame, text="Corpo do E-mail:", **self.estilo["label"]).pack(pady=(15, 0), padx=10, anchor='w')
        self.txt_corpo_email = scrolledtext.ScrolledText(self.scrollable_frame, width=80, height=20, bg="white", fg=self.cor_texto, relief="solid")
        self.txt_corpo_email.pack(padx=10, pady=5, fill='both', expand=True)
        self.txt_corpo_email.insert(tk.END, self.email_body_default)

        tk.Label(self.scrollable_frame, text="Tempo entre envios (segundos):", **self.estilo["label"]).pack(pady=(10, 0), padx=10, anchor='w')
        self.delay_entry = tk.Entry(self.scrollable_frame, width=10)
        self.delay_entry.pack(padx=10, anchor='w')
        self.delay_entry.insert(0, str(self.delay_padrao))

        btn_frame = tk.Frame(self.scrollable_frame, bg=self.cor_fundo)
        btn_frame.pack(pady=10, padx=10, fill='x')

        self.btn_iniciar = tk.Button(btn_frame, text="Iniciar Envio", command=self.iniciar_envio, **self.estilo["botao"])
        self.btn_iniciar.pack(side="left", padx=5, fill='x', expand=True)

        self.btn_pausar = tk.Button(btn_frame, text="Pausar", command=self.pausar_envio, state=tk.DISABLED, bg=self.cor_botao_disabled, fg=self.cor_texto, relief="flat")
        self.btn_pausar.pack(side="left", padx=5, fill='x', expand=True)

        self.btn_retornar = tk.Button(btn_frame, text="Retomar", command=self.retomar_envio, state=tk.DISABLED, bg=self.cor_botao_disabled, fg=self.cor_texto, relief="flat")
        self.btn_retornar.pack(side="left", padx=5, fill='x', expand=True)

        self.btn_encerrar = tk.Button(btn_frame, text="Encerrar", command=self.encerrar_envio, state=tk.DISABLED, bg=self.cor_botao_disabled, fg=self.cor_texto, relief="flat")
        self.btn_encerrar.pack(side="left", padx=5, fill='x', expand=True)

        self.lbl_status = tk.Label(self.scrollable_frame, text="Status: Aguardando início", **self.estilo["label"])
        self.lbl_status.pack(pady=10, padx=10, anchor='w')

    def selecionar_planilha(self):
        path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if path:
            self.planilha_path = path
            self.lbl_planilha.config(text=os.path.basename(path))

    def selecionar_anexo(self):
        paths = filedialog.askopenfilenames(title="Selecionar Arquivos para Anexar")
        if paths:
            self.anexo_path = list(paths)
            arquivos = "\n".join([os.path.basename(p) for p in self.anexo_path])
            self.lbl_anexo.config(text=f"{len(self.anexo_path)} arquivo(s) selecionado(s):\n{arquivos}")

    def iniciar_envio(self):
        if not self.planilha_path or not self.anexo_path:
            messagebox.showwarning("Aviso", "Selecione a planilha e o anexo antes de iniciar.")
            return

        try:
            self.delay_padrao = int(self.delay_entry.get())
        except ValueError:
            messagebox.showerror("Erro", "Tempo entre envios deve ser um número inteiro.")
            return

        self.save_config()

        if self.enviando:
            messagebox.showinfo("Info", "Envio já está em andamento.")
            return

        self.enviando = True
        self.pausado = False
        self.encerrar = False

        self.btn_iniciar.config(state=tk.DISABLED, bg=self.cor_botao_disabled)
        self.btn_pausar.config(state=tk.NORMAL, bg=self.cor_botao)
        self.btn_retornar.config(state=tk.DISABLED, bg=self.cor_botao_disabled)
        self.btn_encerrar.config(state=tk.NORMAL, bg=self.cor_botao)

        thread = threading.Thread(target=self.enviar_emails, daemon=True)
        thread.start()

    def pausar_envio(self):
        if self.enviando:
            self.pausado = True
            self.btn_pausar.config(state=tk.DISABLED, bg=self.cor_botao_disabled)
            self.btn_retornar.config(state=tk.NORMAL, bg=self.cor_botao)
            self.lbl_status.config(text="Status: Pausado")

    def retomar_envio(self):
        if self.enviando and self.pausado:
            self.pausado = False
            self.btn_pausar.config(state=tk.NORMAL, bg=self.cor_botao)
            self.btn_retornar.config(state=tk.DISABLED, bg=self.cor_botao_disabled)
            self.lbl_status.config(text="Status: Enviando...")

    def encerrar_envio(self):
        if self.enviando:
            self.encerrar = True
            self.lbl_status.config(text="Status: Encerrando...")

    def enviar_emails(self):
        try:
            df = pd.read_excel(self.planilha_path)
            for col in ['nome', 'email', 'empresa']:
                if col not in df.columns:
                    self.root.after(0, lambda: messagebox.showerror("Erro", f"Coluna '{col}' não encontrada na planilha."))
                    self.resetar_botoes()
                    return
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Erro ao ler planilha: {e}"))
            self.resetar_botoes()
            return

        try:
            outlook = win32.Dispatch('outlook.application')
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Erro ao iniciar Outlook: {e}"))
            self.resetar_botoes()
            return

        total = len(df)
        indice = 0

        while indice < total and not self.encerrar:
            if self.pausado:
                time.sleep(1)
                continue

            row = df.iloc[indice]
            nome = row.get('nome', '')
            email = row.get('email', '')
            empresa = row.get('empresa', '')

            if not email:
                indice += 1
                continue

            try:
                mail = outlook.CreateItem(0)
                mail.To = email
                titulo = self.entry_subject.get().strip()
                mail.Subject = f"{titulo} para {nome}" if empresa else titulo
                mail.Body = self.txt_corpo_email.get("1.0", tk.END).strip()
                for arquivo in self.anexo_path:
                    mail.Attachments.Add(arquivo)
                mail.Send()

                indice += 1
                self.root.after(0, lambda i=indice: self.lbl_status.config(text=f"Status: Enviando {i} de {total}"))
                time.sleep(self.delay_padrao)

            except Exception as e:
                self.root.after(0, lambda e=e: messagebox.showerror("Erro", f"Erro ao enviar para {email}: {e}"))
                indice += 1

        self.resetar_botoes()
        if self.encerrar:
            self.root.after(0, lambda: self.lbl_status.config(text="Status: Envio encerrado pelo usuário."))
        else:
            self.root.after(0, lambda: self.lbl_status.config(text="Status: Envio concluído."))

    def resetar_botoes(self):
        self.enviando = False
        self.pausado = False
        self.encerrar = False
        self.root.after(0, lambda: self.btn_iniciar.config(state=tk.NORMAL, bg=self.cor_botao))
        self.root.after(0, lambda: self.btn_pausar.config(state=tk.DISABLED, bg=self.cor_botao_disabled))
        self.root.after(0, lambda: self.btn_retornar.config(state=tk.DISABLED, bg=self.cor_botao_disabled))
        self.root.after(0, lambda: self.btn_encerrar.config(state=tk.DISABLED, bg=self.cor_botao_disabled))


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()
