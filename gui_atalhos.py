# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import tempfile
import win32com.client
import shutil
import re

shortcuts = []
startup_path = os.path.join(os.getenv('APPDATA'), 'Microsoft\\Windows\\Start Menu\\Programs\\Startup')

def adicionar_programa():
    caminho = filedialog.askopenfilename(title="Selecione o executável", filetypes=[("Executáveis", "*.exe")])
    if caminho:
        nome = os.path.splitext(os.path.basename(caminho))[0]
        lista.insert(tk.END, f"{nome} -> {caminho}")
        shortcuts.append((caminho, nome))

def criar_atalhos():
    shell = win32com.client.Dispatch("WScript.Shell")
    erros = []
    criados = []

    for caminho, nome in shortcuts:
        try:
            if not os.path.isfile(caminho):
                raise FileNotFoundError("Arquivo não encontrado: " + caminho)

            # Verificação de caracteres problemáticos (ex: emojis)
            if any(ord(c) > 127 for c in caminho):
                raise ValueError("Caminho contém símbolos especiais (como emojis), o que impede a criação do atalho.")

            caminho = os.path.abspath(caminho)
            link_path = os.path.join(startup_path, f"{nome}.lnk")

            try:
                atalho = shell.CreateShortcut(link_path)
                atalho.TargetPath = caminho
                atalho.WorkingDirectory = os.path.dirname(caminho)
                atalho.Save()
            except Exception:
                # Alternativa: criar em local temporário e mover
                temp_link = os.path.join(tempfile.gettempdir(), f"{nome}.lnk")
                atalho = shell.CreateShortcut(temp_link)
                atalho.TargetPath = caminho
                atalho.WorkingDirectory = os.path.dirname(caminho)
                atalho.Save()

                shutil.move(temp_link, link_path)

            criados.append(f"{nome}.lnk")

        except Exception as e:
            erros.append(f"{nome} → {caminho}\nErro: {str(e)}")

    if erros:
        messagebox.showerror("Erros ao criar atalhos", "\n\n".join(erros))
    elif criados:
        messagebox.showinfo("Sucesso", "Atalhos criados com sucesso:\n\n" + "\n".join(criados))
    else:
        messagebox.showwarning("Nenhum atalho", "Nenhum atalho foi criado.")
    root.destroy()

root = tk.Tk()
root.title("Criador de Atalhos (Inicialização)")
root.geometry("600x400")

btn_adicionar = tk.Button(root, text="Adicionar Programa", command=adicionar_programa)
btn_adicionar.pack(pady=10)

lista = tk.Listbox(root, width=80)
lista.pack(pady=10, expand=True, fill=tk.BOTH)

btn_criar = tk.Button(root, text="Criar Atalhos", command=criar_atalhos)
btn_criar.pack(pady=10)

root.mainloop()