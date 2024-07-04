import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import re
import os
from datetime import datetime
import subprocess
import shutil

def extrair_informacoes_especificas(pdf_path, output_dir):
    # Abre o PDF
    documento = fitz.open(pdf_path)
    
    informacoes_encontradas = []

    # Padrões regex para encontrar as informações de lotação específicas
    padrao_lotacao_seduc = r"lotação:\s*SEDUC\s*[-‑]\s*39\s*Coordenadoria\s*Regional\s*de\s*Educação"
    padrao_lotacao_cre = r"lotação:\s*39ª\s*CRE"
    padrao_assunto = r"assunto:\s*(.+)"
    padrao_nome = r"nome:\s*(.+)"
    padrao_id = r"Id\.Func\./Vínculo:\s*(.+)"

    # Buffer para armazenar linhas antes da lotação desejada
    buffer_linhas = []

    # Leia e extraia texto de cada página
    for num_pagina in range(len(documento)):
        pagina = documento.load_page(num_pagina)
        texto = pagina.get_text("text")  # Extrai texto em modo de linha
        linhas = texto.splitlines()  # Divide o texto em linhas

        for idx, linha in enumerate(linhas):
            # Verifica se a linha atual contém a lotação desejada
            if re.search(padrao_lotacao_seduc, linha, re.IGNORECASE) or re.search(padrao_lotacao_cre, linha, re.IGNORECASE):
                # Extrai informações das linhas anteriores, se disponíveis
                if buffer_linhas:
                    paragrafo = "\n".join(buffer_linhas)

                    assunto = re.search(padrao_assunto, paragrafo, re.IGNORECASE)
                    nome = re.search(padrao_nome, paragrafo, re.IGNORECASE)
                    id_match = re.search(padrao_id, paragrafo, re.IGNORECASE)

                    # Extraia as informações se encontradas
                    info_assunto = assunto.group(1).strip().upper() if assunto else "N/A"
                    info_nome = nome.group(1).strip().upper() if nome else "N/A"
                    info_id = id_match.group(1).strip().upper() if id_match else "N/A"

                    # Armazene as informações
                    informacoes_encontradas.append({
                        "assunto": info_assunto,
                        "nome": info_nome,
                        "id": info_id,
                        "pagina": num_pagina + 1
                    })

                # Limpa o buffer após encontrar a lotação desejada
                buffer_linhas = []
            
            else:
                # Adiciona a linha atual ao buffer
                buffer_linhas.append(linha)

                # Mantém o buffer com no máximo 6 linhas
                if len(buffer_linhas) > 6:
                    buffer_linhas.pop(0)

    # Caminho completo para o arquivo de saída na pasta de downloads
    output_path = os.path.join(output_dir, "informacoes_extraidas.txt")

    # Escreva as informações encontradas em um arquivo de texto
    with open(output_path, "w", encoding="utf-8") as f:
        for info in informacoes_encontradas:
            f.write(f"ASSUNTO: {info['assunto']}\n")
            f.write(f"NOME: {info['nome']}\n")
            f.write(f"ID FUNCIONAL: {info['id']}\n")
            f.write(f"PÁGINA: {info['pagina']}\n")
            f.write("\n")
    
    print(f"Extração concluída. Informações salvas em '{output_path}'")
    documento.close()

    return output_path

def selecionar_pdf():
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal do Tkinter
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    return file_path

def selecionar_output_dir():
    # Pasta de downloads do usuário
    output_dir = os.path.expanduser("~/Downloads")
    return output_dir

# Seleciona o arquivo PDF para extração de informações
pdf_path = selecionar_pdf()

if pdf_path:
    # Seleciona o diretório de saída (downloads)
    output_dir = selecionar_output_dir()

    # Chama a função de extração
    output_file = extrair_informacoes_especificas(pdf_path, output_dir)

    # Abre o arquivo de texto após salvar
    if output_file:
        subprocess.Popen(["notepad.exe", output_file])
    else:
        messagebox.showerror("Erro", "Falha ao salvar o arquivo de texto.")
else:
    print("Operação cancelada: Arquivo PDF não selecionado.")
