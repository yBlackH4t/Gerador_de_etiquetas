# Gerador de Etiquetas de NFe

Este é um programa em Python com interface gráfica (GUI) que automatiza a criação de etiquetas para notas fiscais eletrônicas (NFe). Ele lê um arquivo XML, extrai as informações do destinatário, transportadora, numero da NF, volume, peso e etc... E preenche um modelo de etiqueta em Excel, pronto para impressão.

---

### **Como Usar**

Você tem duas opções para usar o programa:

#### **Opção 1: Baixar o Executável (Recomendado)**

Se você não tem Python instalado ou prefere uma solução mais simples, basta baixar o arquivo executável (`.exe`) e rodar.

1.  Vá para a página de [**Releases**](https://github.com/yBlackH4t/Gerador_de_etiquetas/releases/download/v1.0.0/Etiquetas.exe) do projeto.
2.  Baixe o arquivo `Etiquetas.exe` na versão mais recente.
3.  Execute o arquivo e a interface gráfica será aberta.



#### **Opção 2: Rodar o Código Fonte**

Se você é um desenvolvedor ou prefere inspecionar o código, pode clonar este repositório e rodar o script diretamente.

**Pré-requisitos:**
* Python 3.x
* As bibliotecas `openpyxl` e `PyInstaller`.

**Instalação das dependências:**

```sh
pip install openpyxl
pip install pyinstaller
```

**Executando o script:**
```
python Etiquetas.py
```

## Funcionalidades

* **Extração Automática:** Busca o arquivo XML de NFe mais recente em uma pasta e extrai as informações (cliente, transportadora, peso, volume, etc.).
* **Preenchimento de Modelo:** Preenche um modelo de etiqueta em formato Excel (.xlsx) com os dados extraídos.
* **Impressão Direta:** Pergunta se o usuário deseja enviar a etiqueta preenchida para impressão.
* **Memorização de Caminhos:** Salva os últimos caminhos de pastas e arquivos utilizados para facilitar o uso futuro.

**Desenvolvido por:** _Guilherme 'yBlackH4t' Souza._
