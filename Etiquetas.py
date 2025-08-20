import xml.etree.ElementTree as ET
from openpyxl import load_workbook
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import sys
import json


def get_most_recent_xml(pasta):
    try:
        arquivos_xml = [os.path.join(pasta, f) for f in os.listdir(pasta) if f.endswith('.xml')]
        if not arquivos_xml:
            return None
        mais_recente = max(arquivos_xml, key=os.path.getmtime)
        return mais_recente
    except FileNotFoundError:
        messagebox.showerror("Vish, deu erro!", f"A pasta '{pasta}' não foi encontrada. Confere o caminho, por favor.")
        return None

def processar_xml(arquivo_dados_xml, arquivo_modelo, arquivo_saida, celulas): 
    # Primeiro, vamos ver se o arquivo XML não está vazio
    if os.stat(arquivo_dados_xml).st_size == 0:
        messagebox.showwarning("Hmm, arquivo vazio", f"O arquivo XML '{arquivo_dados_xml}' está vazio. Melhor pular esse.")
        return

    try:
        arvore = ET.parse(arquivo_dados_xml)
        raiz = arvore.getroot()
        # Esse 'ns' é o que faz o Python entender o formato do XML de nota fiscal
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

        dest = raiz.find('.//nfe:dest', ns)
        transp = raiz.find('.//nfe:transp/nfe:transporta', ns)
        
        # Pega os dados do destinatário
        cliente = dest.find('nfe:xNome', ns).text if dest is not None and dest.find('nfe:xNome', ns) is not None else "Não Encontrado"
        cidade = dest.find('nfe:enderDest/nfe:xMun', ns).text if dest is not None and dest.find('nfe:enderDest/nfe:xMun', ns) is not None else "Não Encontrado"
        estado = dest.find('nfe:enderDest/nfe:UF', ns).text if dest is not None and dest.find('nfe:enderDest/nfe:UF', ns) is not None else "Não Encontrado"

        # Pega os dados da transportadora, peso e volume
        transpo = transp.find('nfe:xNome', ns).text if transp is not None else "Não Encontrado"
        
        try:
            volume = raiz.find('.//nfe:transp/nfe:vol/nfe:qVol', ns).text
        except AttributeError:
            volume = "Não Informado"
        try:
            # Pega o peso e limpa ele pra ficar bonitinho
            peso = raiz.find('.//nfe:transp/nfe:vol/nfe:pesoL', ns).text
            if peso:
                try:
                    peso_numerico = float(peso)
                    if peso_numerico.is_integer():
                        peso = str(int(peso_numerico))
                    else:
                        peso = str(peso_numerico)
                except (ValueError, TypeError):
                    pass
            else:
                peso = "Não Informado"
        except AttributeError:
            peso = "Não Informado"

        # O número da nota fiscal
        numero_nota = raiz.find('.//nfe:ide/nfe:nNF', ns).text
        
        # --- Vamos ver se deu tudo certo, imprimindo aqui pra conferir --
        print("--- Ó, o que a gente conseguiu pegar do XML ---")
        print(f"CLIENTE: {cliente}")
        print(f"CIDADE: {cidade}")
        print(f"ESTADO: {estado}")
        print(f"TRANSPORTADORA: {transpo}")
        print(f"NÚMERO DA NOTA: {numero_nota}")
        print(f"VOLUME: {volume}")
        print(f"PESO: {peso}")
        print("-" * 30)

        wb_saida = load_workbook(filename=arquivo_modelo)
        ws_saida = wb_saida.active
        
        ws_saida[celulas['cliente']] = f"CLIENTE: {cliente}"
        ws_saida[celulas['cidade_estado']] = f"CIDADE: {cidade}             ESTADO: {estado}"
        ws_saida[celulas['transportadora']] = f"TRANSP: {transpo}"
        ws_saida[celulas['numero_nota']] = f"NOTA: {numero_nota}"
        ws_saida[celulas['volume_peso']] = f"VOL: {volume}                         PESO: {peso}Kg"

    
        wb_saida.save(arquivo_saida)
        messagebox.showinfo("Sucesso!", f"Etiqueta salva com sucesso em '{arquivo_saida}'. Prontinho pra imprimir!")

        # Pergunta se a gente quer imprimir
        resposta = messagebox.askyesno("Imprimir Etiqueta", "Etiqueta gerada com sucesso. Quer imprimir agora?")
        if resposta:
            try:
                # Manda o comando de impressão pro sistema
                os.startfile(arquivo_saida, "print")
                messagebox.showinfo("Impressão", "A janela de impressão vai abrir. É só escolher a impressora e 'Imprimir'.")
            except Exception as e:
                messagebox.showerror("Erro de Impressão", f"Não consegui iniciar a impressão. Erro: {e}")
        else:
            messagebox.showinfo("Impressão", "Impressão cancelada. Você pode imprimir depois.")

    except FileNotFoundError as e:
        messagebox.showerror("Arquivo não encontrado", f"O arquivo '{e.filename}' não foi encontrado. Confere o caminho, por favor.")
    except ET.ParseError as e:
        messagebox.showerror("Erro no XML", f"O arquivo XML pode estar vazio ou corrompido. Detalhes: {e}")
    except Exception as e:
        messagebox.showerror("Ops, deu erro!", f"Ocorreu um erro inesperado: {e}")

def save_config(config):
    try:
        with open("config.json", "w") as f:
            json.dump(config, f)
    except Exception as e:
        print(f"Não deu pra salvar a configuração: {e}")

def load_config():
    try:
        with open("config.json", "r") as f:
            config = json.load(f)
            return config
    except FileNotFoundError:
        # Na primeira vez que rodar, a gente usa os valores padrão
        return {
            "caminho_xml": r'C:\nQuestorEmp\NFe\Geradas',
            "caminho_modelo": r'd:\User\Desktop\ETIQUETA.xlsx',
            "caminho_saida": r'd:\User\Desktop\ETIQUETA_PREENCHIDA.xlsx',
            "celula_cliente": "A4",
            "celula_cidade_estado": "A5",
            "celula_transportadora": "A6",
            "celula_numero_nota": "A7",
            "celula_volume_peso": "A8"
        }
    except Exception as e:
        print(f"Não consegui carregar a configuração: {e}")
        return {}


# --- A interface do programa, o que a gente vê na tela ---
def main():
    root = tk.Tk()
    root.title("Gerador de Etiquetas NFe")
    root.geometry("400x600")

    config = load_config()

    # Variaveis pra guardar os caminhos e células
    caminho_xml_var = tk.StringVar(value=config.get("caminho_xml"))
    caminho_modelo_var = tk.StringVar(value=config.get("caminho_modelo"))
    caminho_saida_var = tk.StringVar(value=config.get("caminho_saida"))
    
    celula_cliente_var = tk.StringVar(value=config.get("celula_cliente"))
    celula_cidade_estado_var = tk.StringVar(value=config.get("celula_cidade_estado"))
    celula_transportadora_var = tk.StringVar(value=config.get("celula_transportadora"))
    celula_numero_nota_var = tk.StringVar(value=config.get("celula_numero_nota"))
    celula_volume_peso_var = tk.StringVar(value=config.get("celula_volume_peso"))

    # A parte de selecionar os caminhos
    frame_caminhos = tk.Frame(root)
    frame_caminhos.pack(pady=10)

    tk.Label(frame_caminhos, text="Pasta das NFes (XML):").pack()
    tk.Entry(frame_caminhos, textvariable=caminho_xml_var, width=50).pack()
    tk.Button(frame_caminhos, text="Procurar Pasta", command=lambda: caminho_xml_var.set(filedialog.askdirectory())).pack()

    tk.Label(frame_caminhos, text="Arquivo Modelo da Etiqueta:").pack()
    tk.Entry(frame_caminhos, textvariable=caminho_modelo_var, width=50).pack()
    tk.Button(frame_caminhos, text="Procurar Arquivo", command=lambda: caminho_modelo_var.set(filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")]))).pack()

    tk.Label(frame_caminhos, text="Arquivo de Saída da Etiqueta:").pack()
    tk.Entry(frame_caminhos, textvariable=caminho_saida_var, width=50).pack()
    tk.Button(frame_caminhos, text="Salvar Como...", command=lambda: caminho_saida_var.set(filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")]))).pack()

    # A parte de selecionar as células
    frame_celulas = tk.Frame(root)
    frame_celulas.pack(pady=10)
    
    tk.Label(frame_celulas, text="Células pra preencher:").pack()
    
    tk.Label(frame_celulas, text="Cliente:").pack()
    tk.Entry(frame_celulas, textvariable=celula_cliente_var, width=10).pack()
    
    tk.Label(frame_celulas, text="Cidade/Estado:").pack()
    tk.Entry(frame_celulas, textvariable=celula_cidade_estado_var, width=10).pack()
    
    tk.Label(frame_celulas, text="Transportadora:").pack()
    tk.Entry(frame_celulas, textvariable=celula_transportadora_var, width=10).pack()
    
    tk.Label(frame_celulas, text="Número da Nota:").pack()
    tk.Entry(frame_celulas, textvariable=celula_numero_nota_var, width=10).pack()
    
    tk.Label(frame_celulas, text="Volume/Peso:").pack()
    tk.Entry(frame_celulas, textvariable=celula_volume_peso_var, width=10).pack()
    
    def on_run():
        pasta_monitorar = caminho_xml_var.get()
        arquivo_modelo = caminho_modelo_var.get()
        arquivo_saida = caminho_saida_var.get()
        
        celulas = {
            'cliente': celula_cliente_var.get(),
            'cidade_estado': celula_cidade_estado_var.get(),
            'transportadora': celula_transportadora_var.get(),
            'numero_nota': celula_numero_nota_var.get(),
            'volume_peso': celula_volume_peso_var.get()
        }

        if not pasta_monitorar or not arquivo_modelo or not arquivo_saida or any(not v for v in celulas.values()):
            messagebox.showerror("Faltou preencher!", "Por favor, preenche todos os campos antes de continuar.")
            return

        print("Beleza, bora começar!")
        arquivo_mais_recente = get_most_recent_xml(pasta_monitorar)
        
        if arquivo_mais_recente:
            print(f"Achou o arquivo XML mais recente: {arquivo_mais_recente}")
            processar_xml(arquivo_mais_recente, arquivo_modelo, arquivo_saida, celulas)
        else:
            messagebox.showinfo("Nenhum Arquivo", f"Não achamos nenhum arquivo XML na pasta '{pasta_monitorar}'.")
        
        print("\nPronto, terminamos!")

        # A gente salva as configurações pra usar na próxima vez
        current_config = {
            "caminho_xml": pasta_monitorar,
            "caminho_modelo": arquivo_modelo,
            "caminho_saida": arquivo_saida,
            "celula_cliente": celulas['cliente'],
            "celula_cidade_estado": celulas['cidade_estado'],
            "celula_transportadora": celulas['transportadora'],
            "celula_numero_nota": celulas['numero_nota'],
            "celula_volume_peso": celulas['volume_peso']
        }
        save_config(current_config)

    tk.Button(root, text="Executar", command=on_run).pack(pady=10)
    
    # Isso aqui faz o console sumir no Windows pra não atrapalhar
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
        except:
            pass

    root.mainloop()

if __name__ == "__main__":
    main()
