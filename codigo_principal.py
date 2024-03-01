import customtkinter as ctk
from customtkinter import CTk, CTkButton, CTkProgressBar, CTkLabel, CTkEntry, set_appearance_mode, set_default_color_theme
from tkinter import filedialog, messagebox
import os
from threading import Thread
from codigo_secundario import executar_codigo_secundario

def executar_codigo_secundario_thread():
    button_gerar.configure(state="disabled") 
    processo_conclusao.set(0) 
    Thread(target=executar_codigo_secundario_com_progresso).start() 

def executar_codigo_secundario_com_progresso():
    executar_codigo_secundario() 
    processo_conclusao.set(100) 
    button_gerar.configure(state="normal")
    messagebox.showinfo("Concluído", "O processo foi concluído com sucesso!") 

def buscar_arquivo():
    filename = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos RAR", "*.rar"), ("Todos os arquivos", "*.*")))
    if filename:
        input_busca.delete(0, "end")
        input_busca.insert(0, filename) 

root = CTk()
root.title('')
root.geometry('385x200')
set_appearance_mode("dark")
set_default_color_theme("dark-blue")
root.resizable(False, False)

title = CTkLabel(root, text='Gerador Tarifador', text_color='white', font=('Bold', 20))
title.place(x=110, y=10)

input_busca = CTkEntry(root, width=250, height=26, placeholder_text='Caminho do arquivo .ZIP', font=('Bold', 12))
input_busca.place(x=30, y=70)

button_busca = CTkButton(root, text='Localizar', width=30, height=25, command=buscar_arquivo)
button_busca.place(x=290, y=70)

button_baixar = CTkButton(root, text='Baixar .ZIP', width=40, height=25)
button_baixar.place(x=30, y=140)

button_gerar = CTkButton(root, text='Converter', width=40, height=25, command=executar_codigo_secundario_thread)  
button_gerar.place(x=210, y=140)

processo_conclusao = CTkProgressBar(root, width=250, height=10)
processo_conclusao.place(x=30, y=100)
processo_conclusao.set(0) 

root.mainloop()
