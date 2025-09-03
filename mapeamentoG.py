import os
import subprocess
import tkinter as tk
from tkinter import messagebox
import win32com.client

# Configurações da rede
REDE_PATH = r"\\127.166.0.0\wilson"
USUARIO = "admwgatos"
SENHA = "matrox2563"
UNIDADE = "M:"  # Letra usada para mapear a unidade de rede

# Nome do aplicativo a procurar
NOME_ARQUIVO = "PCINF000.EXE"

def mapear_rede():
    try:
        # Primeiro desconectar caso já exista mapeamento
        subprocess.run(f'net use {UNIDADE} /delete /y', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        # Mapear novamente
        comando = f'net use {UNIDADE} "{REDE_PATH}" /user:{USUARIO} {SENHA}'
        resultado = subprocess.run(comando, shell=True, capture_output=True, text=True)

        if resultado.returncode != 0:
            messagebox.showerror("Erro", f"Falha ao mapear rede:\n{resultado.stderr}")
            return False

        return True

    except Exception as e:
        messagebox.showerror("Erro", str(e))
        return False

def procurar_arquivo():
    # Busca o arquivo dentro da unidade mapeada
    for raiz, _, arquivos in os.walk(UNIDADE + "\\"):
        for arquivo in arquivos:
            if arquivo.upper() == NOME_ARQUIVO.upper():
                return os.path.join(raiz, arquivo)
    return None

def criar_atalho(caminho_alvo, caminho_atalho):
    shell = win32com.client.Dispatch("WScript.Shell")
    atalho = shell.CreateShortCut(caminho_atalho)
    atalho.Targetpath = caminho_alvo
    atalho.WorkingDirectory = os.path.dirname(caminho_alvo)
    atalho.IconLocation = caminho_alvo
    atalho.save()

def executar():
    if not mapear_rede():
        return

    # Caminho da Área de Trabalho do usuário
    desktop_path = os.path.join(os.environ["USERPROFILE"], "Desktop")
    destino_atalho = os.path.join(desktop_path, f"{os.path.splitext(NOME_ARQUIVO)[0]}.lnk")

    if os.path.exists(destino_atalho):
        messagebox.showinfo("Info", "O atalho já existe na Área de Trabalho.\nSomente o mapeamento foi realizado.")
    else:
        origem = procurar_arquivo()
        if origem:
            criar_atalho(origem, destino_atalho)
            messagebox.showinfo("Sucesso", f"Atalho para {NOME_ARQUIVO} criado na Área de Trabalho.")
        else:
            messagebox.showwarning("Aviso", f"{NOME_ARQUIVO} não encontrado na pasta mapeada.")

# Criando a interface gráfica
janela = tk.Tk()
janela.title("Mapeador de Rede")
janela.geometry("300x150")

botao = tk.Button(janela, text="Executar", command=executar, height=2, width=15)
botao.pack(expand=True)

janela.mainloop()
