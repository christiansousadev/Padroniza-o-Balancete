# Conversor de Balancetes - Fortes para Accountfy

Automação contábil desenvolvida em Python para padronização de balancetes exportados do ERP Fortes, preparando-os para importação no Accountfy ou análise em Excel.

## 🚀 Funcionalidades
* **Detecção Dinâmica:** Identifica automaticamente o início dos dados, ignorando cabeçalhos de relatórios.
* **Tratamento de Dados:** Remove notação científica e limpa máscaras de contas contábeis.
* **Exportação Dual:** Suporta geração de arquivos `.xlsx` formatados ou `.csv` no padrão Databuilder (Accountfy).
* **Interface Moderna:** UI responsiva que não trava durante o processamento (uso de threading).

## 🛠️ Requisitos
* Python 3.11+
* Bibliotecas listadas em `requirements.txt`

## 📦 Como rodar (Desenvolvimento)
1. Instale as dependências:
   ```bash
   pip install -r requirements.txt

## Comando instalador
pyinstaller --noconsole --onefile --collect-all customtkinter --add-data "icon.ico;." --icon="icon.ico" --name "Conversor Balancete" app.py

# projeto no github
* git init
* git add .
* git commit -m "feat: inicio do projeto"
* git push origin main