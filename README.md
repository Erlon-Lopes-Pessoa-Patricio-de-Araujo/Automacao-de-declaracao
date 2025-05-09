# DocWise V1 - Document Automation and Conversion Tool

**DocWise** é uma aplicação desktop para automação de documentos, com foco em geração de declarações em lote, conversão entre formatos e manuseio inteligente de arquivos do Microsoft Word, PDF e Excel. A interface intuitiva e os recursos avançados tornam esta ferramenta ideal para uso em ambientes acadêmicos, administrativos ou empresariais.

---

# 🌟 Funcionalidades

- 📃 **Geração em lote de documentos** usando modelos .docx e planilhas .xlsx.
- 🔁 **Conversão de arquivos**:
  - DOCX → PDF (via Word ou LibreOffice)
  - PDF → DOCX (com suporte a OCR)
- 🔍 **Detecção automática de placeholders** (`{{Nome}}`, `{{CPF}}`, etc.).
- 📁 **Organização de saída em subpastas** com base em colunas.
- 📅 **Agendamento de tarefas** para execução futura.
- 🧠 **Histórico de execuções** com salvamento automático (SQLite).
- 🧾 **Relatórios em PDF ou CSV** após o processamento.
- 🛠️ **Compatibilidade com Word e LibreOffice** (modo headless).
- 🖥️ **Interface gráfica (GUI)** com `tkinter`.

---

## 📊 Requisitos

- **Python 3.8+**
- **LibreOffice** instalado (opcional)
- **Tesseract OCR** (opcional, para PDFs escaneados)

Instale as dependências com:

```bash
pip install -r requirements.txt
````

---

## 🚀 Como usar

1. **Execute o programa:**

2. **Na interface:**

   * Selecione um modelo `.docx` com os campos personalizados (`{{Nome}}`, `{{Curso}}`, etc.).
   * Carregue uma planilha `.xlsx` com os dados.
   * Escolha a pasta de saída.
   * Configure o nome dos arquivos e opções extras (ZIP, subpastas).
   * Clique em **Gerar declarações** ou **Converter arquivos**.

---

# 🔌 Integrações e Tecnologias

* `tkinter`: Interface gráfica
* `pandas`: Leitura de planilhas Excel
* `python-docx`: Manipulação de arquivos Word
* `docx2pdf`, `pdf2docx`: Conversão entre DOCX e PDF
* `reportlab`: Geração de relatórios PDF
* `sqlite3`, `psutil`, `schedule`, `threading`, `logging`: gerenciamento e desempenho
* `pytesseract` + `Pillow`: OCR para PDFs escaneados

---
# 📂 Estrutura sugerida

```
.
├── app9corrigido.py         # Código-fonte principal
├── config_v4.json           # Configurações persistentes
├── execution_history.db     # Banco de dados SQLite
├── requirements.txt         # Dependências
├── /templates               # Modelos Word (.docx)
└── /output                  # Arquivos gerados
```

---

# 📌 Recursos futuros sugeridos

* Suporte a mais formatos (ODT, HTML)
* Integração com armazenamento em nuvem
* Envio automático por e-mail
* Tradução automática de campos

---

# 📄 Licença

Este projeto é distribuído como software fechado. Para uso pessoal, institucional ou colaborações, entre em contato com o desenvolvedor.

---

> Desenvolvido para automatizar tarefas repetitivas de geração e gestão de documentos em ambientes acadêmicos e administrativos. Foco em desempenho, confiabilidade e facilidade de uso.

```

---

Se quiser, posso gerar também um `requirements.txt` personalizado com base nas bibliotecas do seu código. Deseja isso?
```
