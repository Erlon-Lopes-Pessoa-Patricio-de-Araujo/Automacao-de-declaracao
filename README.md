# Automacao-de-declaracao

# DocWise V1 - Document Automation and Conversion Tool

**DocWise** é uma aplicação robusta para automação de documentos, com funcionalidades de geração em lote de declarações personalizadas, conversão de arquivos entre DOCX e PDF, integração com OCR para PDFs escaneados e interface gráfica amigável baseada em `Tkinter`.

## 🎯 Funcionalidades Principais

- 📝 **Geração de declarações em lote** a partir de modelos Word (.docx) e planilhas Excel (.xlsx).
- 🔁 **Conversão de documentos**:
  - DOCX → PDF
  - PDF → DOCX (com suporte a OCR)
- 🧠 **Detecção automática de placeholders** no modelo.
- 📁 **Criação de subpastas e nomes personalizados** com base em colunas da planilha.
- 📦 **Geração de arquivos ZIP** com os documentos gerados.
- 🔧 **Configurações persistentes** e histórico de execuções via SQLite e JSON.
- 🔌 **Compatibilidade com Microsoft Word e LibreOffice** (modo headless).
- 🔍 **OCR com Tesseract** para PDF escaneado (opcional).

## 🖥️ Tecnologias Utilizadas

- `Python 3.x`
- `Tkinter` para interface gráfica
- `pandas`, `docx`, `reportlab`, `pdf2docx`, `docx2pdf`, `pytesseract`, `psutil`, entre outras
- `SQLite` para gerenciamento de histórico
- `LibreOffice` (modo headless) para conversões alternativas

## 📂 Estrutura do Projeto

```bash
.
├── app9corrigido.py       # Código principal da aplicação
├── config_v4.json         # Arquivo de configuração persistente (gerado automaticamente)
├── execution_history.db   # Histórico das execuções (gerado automaticamente)
├── templates/             # Pasta recomendada para armazenar seus modelos .docx
└── output/                # Pasta de saída recomendada

📅 Recursos futuros sugeridos

Suporte a mais formatos (ODT, HTML)

Sincronização com nuvem

Exportação direta para e-mail

Detecção e tradução automática de campos

📄 Licença

Este projeto é distribuído como software fechado. Para uso pessoal, institucional ou colaboração, entre em contato com o desenvolvedor.

Desenvolvido para automatizar tarefas repetitivas de geração e gestão de documentos em ambientes acadêmicos e administrativos. Com foco em desempenho, confiabilidade e facilidade de uso.
