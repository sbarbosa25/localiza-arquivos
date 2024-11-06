import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime

# Função para selecionar diretório
def selecionar_diretorio(titulo):
    """Abre um diálogo para selecionar o diretório."""
    return filedialog.askdirectory(title=titulo)

# Função para selecionar arquivo Excel
def selecionar_arquivo_excel():
    """Abre um diálogo para selecionar um arquivo Excel."""
    return filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])

# Função para copiar pastas e gerar log
def copiar_conteudo_pastas(diretorio, excel_path, destino):
    """Copia o conteúdo das pastas listadas na terceira coluna do Excel para o diretório de destino."""
    pastas_copiadas = 0
    pastas_existentes = 0
    pastas_nao_encontradas = []
    planilha = pd.read_excel(excel_path)

    # Considera que a terceira coluna (índice 1) contém os nomes das pastas. Você pode alterar a coluna de acordo com sua necessidade.
    nomes_pastas = planilha.iloc[:, 1].dropna().unique()

    for nome_pasta in nomes_pastas:
        caminho_pasta = os.path.join(diretorio, nome_pasta)
        destino_pasta = os.path.join(destino, nome_pasta)

        if os.path.exists(caminho_pasta) and os.path.isdir(caminho_pasta):
            # Verifica se a pasta de destino já existe
            if os.path.exists(destino_pasta):
                pastas_existentes += 1
                print(f"Pasta {nome_pasta} já existe em {destino_pasta}. Pulando para o próximo item.")
                continue  # Se a pasta já existir, pula para o próximo item

            # Cria a pasta de destino se não existir
            os.makedirs(destino_pasta)

            # Copia todo o conteúdo da pasta original para a pasta de destino
            for item in os.listdir(caminho_pasta):
                src_item = os.path.join(caminho_pasta, item)
                dst_item = os.path.join(destino_pasta, item)

                try:
                    if os.path.isdir(src_item):
                        shutil.copytree(src_item, dst_item, dirs_exist_ok=True)
                    else:
                        shutil.copy2(src_item, dst_item)
                except shutil.Error as e:
                    print(f"Erro ao copiar {src_item}: {e}")
                except OSError as e:
                    # Tratamento específico para o erro de nuvem (WinError 380)
                    if "WinError 380" in str(e):
                        print(f"Arquivo na nuvem não disponível localmente: {src_item}. Pulando...")
                    else:
                        print(f"Erro ao copiar {src_item}: {e}")

            pastas_copiadas += 1
            print(f"Pasta {nome_pasta} copiada para {destino_pasta}")
        else:
            # Se a pasta não for encontrada, adiciona à lista de pastas não encontradas
            pastas_nao_encontradas.append(nome_pasta)

    # Geração do log
    gerar_log(pastas_copiadas, pastas_existentes, pastas_nao_encontradas)

    return pastas_copiadas, pastas_existentes, len(pastas_nao_encontradas)

# Função para gerar log
def gerar_log(pastas_copiadas, pastas_existentes, pastas_nao_encontradas):
    """Gera um arquivo de log informando quantas pastas foram copiadas, quantas já existiam e quantas não foram encontradas."""
    log_filename = f"log_copia_pastas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(log_filename, 'w') as log_file:
        log_file.write(f"Log de Cópia de Pastas - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        log_file.write(f"Pastas copiadas: {pastas_copiadas}\n")
        log_file.write(f"Pastas já existentes (não copiadas): {pastas_existentes}\n")
        log_file.write(f"Pastas não encontradas: {len(pastas_nao_encontradas)}\n")

        if pastas_nao_encontradas:
            log_file.write("\nLista de pastas não encontradas:\n")
            for pasta in pastas_nao_encontradas:
                log_file.write(f"- {pasta}\n")

    print(f"Log gerado: {log_filename}")

# Função principal
def main():
    # Configuração da interface gráfica
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal

    diretorio = selecionar_diretorio("Selecione o diretório contendo as pastas")
    excel_path = selecionar_arquivo_excel()
    destino = selecionar_diretorio("Selecione o diretório de destino para as pastas copiadas")

    if diretorio and excel_path and destino:
        pastas_copiadas, pastas_existentes, pastas_nao_encontradas = copiar_conteudo_pastas(diretorio, excel_path, destino)
        messagebox.showinfo("Resultado", f"{pastas_copiadas} pasta(s) copiada(s).\n"
                                         f"{pastas_existentes} pasta(s) já existia(m) e não foram copiadas.\n"
                                         f"{pastas_nao_encontradas} pasta(s) não encontrada(s).")
    else:
        messagebox.showwarning("Operação cancelada", "Um ou mais diretórios/arquivo não foram selecionados.")

if __name__ == "__main__":
    main()

