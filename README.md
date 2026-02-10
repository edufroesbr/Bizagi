# Automação Bizagi ONS

Este repositório contém scripts de automação para o processamento e aprovação de casos no sistema Bizagi do ONS (Operador Nacional do Sistema Elétrico). A aplicação utiliza Playwright para interação com o navegador e validação de documentos legais e financeiros.

## 🚀 Funcionalidades

- **Login Automatizado**: Autenticação automática no portal Bizagi.
- **Processamento em Lote (Ajuste)**: Script para buscar casos específicos, tomar posse, marcar documentos para ajuste e preencher observações automaticamente.
- **Aprovação em Lote**: Script para automatizar o fluxo de aprovação de múltiplos casos.
- **Validação Financeira**: Integração com planilhas Excel para validar valores de débitos e CPB/CUST.
- **Validação de Documentos**: Verificação de conformidade de documentos (CADIN, Protesto, Termo de Compromisso) em arquivos PDF.
- **Relatórios**: Geração automática de relatórios de processamento em formato CSV.

## 🛠️ Estrutura do Projeto

- `run_batch_v3.py`: Script principal para processamento de ajustes em lote.
- `run_approve_batch.py`: Script para aprovação em lote de casos.
- `bizagi_bot.py`: Classe principal que orquestra as interações com o Playwright.
- `validator.py`: Lógica de validação de documentos e conformidade com a Resolução 1125.
- `excel_helper.py`: Utilitários para leitura e busca em planilhas Excel.
- `case_reporter.py`: Módulo responsável pela geração do relatório `case_report.csv`.
- `config.py`: Configurações globais da aplicação (URLs, caminhos, etc).

## 📋 Pré-requisitos

- Python 3.10+
- Playwright (`pip install playwright`)
- Navegador Chromium instalado (`playwright install chromium`)
- Planilha Mestre de Contratos no caminho especificado em `config.py`.

## 📖 Como Usar

### Processamento de Ajustes
Execute o script `run_batch_v3.py`:
```bash
python run_batch_v3.py
```
O script solicitará a lista de IDs de casos ou utilizará a lista padrão definida no código.

### Aprovação de Casos
Execute o script `run_approve_batch.py`:
```bash
python run_approve_batch.py
```

## 📄 Relatórios
Após a execução, os resultados serão salvos no arquivo `case_report.csv` no diretório raiz do projeto.
