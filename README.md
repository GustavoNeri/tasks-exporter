# Instruções do Projeto

## Resumo do projeto
Este projeto contém scripts para atualizar tarefas e gerar relatórios a partir de dados de tarefas. Ele processa páginas web, extrai informações relevantes e salva resultados em formatos do Office:
- `atualizar_tarefas_excel.py` — trabalha com planilhas Excel (`.xlsx`) usando `openpyxl`
- `atualizar_tarefas_word.py` — gera ou atualiza documentos Word (`.docx`) usando `python-docx`

O projeto depende de bibliotecas Python para requisições HTTP, parsing de HTML e manipulação de arquivos de Office.

## Dependências
As dependências são listadas em `requirements.txt`:
- `requests`
- `beautifulsoup4`
- `openpyxl`
- `python-docx`

## Instalação das dependências

### Ubuntu 20.04
No terminal, navegue até a pasta do projeto e execute:

```bash
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt
```

Se quiser instalar apenas para o usuário atual:

```bash
python3 -m pip install --user -r requirements.txt
```

### Windows 11 Home
No PowerShell ou Prompt de Comando, navegue até a pasta do projeto e execute:

```powershell
py -3 -m pip install --upgrade pip
py -3 -m pip install -r requirements.txt
```

Se `py` não estiver disponível, use:

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Como executar
Após instalar as dependências, execute os scripts diretamente:

```bash
python3 atualizar_tarefas_excel.py
python3 atualizar_tarefas_word.py
```

Ou, no Windows:

```powershell
py -3 atualizar_tarefas_excel.py
py -3 atualizar_tarefas_word.py
```

> Ajuste o comando de execução conforme o Python instalado no seu sistema.

