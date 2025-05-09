# DocWise V1 - Document Automation and Conversion Tool

**DocWise** é uma aplicação desktop para automação de documentos, com foco em geração de declarações em lote, conversão entre formatos e manuseio inteligente de arquivos do Microsoft Word, PDF e Excel. A interface intuitiva e os recursos avançados tornam esta ferramenta ideal para uso em ambientes acadêmicos, administrativos ou empresariais.

> ✨ Este projeto é **open-source** e pode ser usado livremente, desde que seja mantida uma **citação ou crédito** ao autor original.

> ✍️ Foi criado com uma ideia simples para automatizar alguns processos no trabalho, mas acabou se tornando um sistema maior, que ainda tem muito o que melhorar.

---

# 🌟 Funcionalidades

* 📃 **Geração em lote de documentos** usando modelos .docx e planilhas .xlsx.
* 🔄 **Conversão de arquivos**:

  * DOCX → PDF (via Word ou LibreOffice)
  * PDF → DOCX (com suporte a OCR)
* 🔍 **Detecção automática de placeholders** (`{{Nome}}`, `{{CPF}}`, etc.).
* 📁 **Organização de saída em subpastas** com base em colunas.
* 📆 **Agendamento de tarefas** para execução futura.
* 🏋️ **Histórico de execuções** com salvamento automático (SQLite).
* ⚖️ **Relatórios em PDF ou CSV** após o processamento.
* 🔧 **Compatibilidade com Word e LibreOffice** (modo headless).
* 🔮 **Interface gráfica (GUI)** com `tkinter`.

---

# 📊 Requisitos

* **Python 3.8+**
* **LibreOffice** instalado (opcional)
* **Tesseract OCR** (opcional, para PDFs escaneados)

Instale as dependências com:

```bash
pandas
python-docx
docx2pdf
pdf2docx
Pillow
pytesseract
reportlab
schedule
psutil
comtypes
tk

```

---

# 🚀 Como usar

1. **Execute o programa:**

```bash
Prototipo_automacao.py
```

2. **Na interface:**

   * Selecione um modelo `.docx` com os campos personalizados (`{{Nome}}`, `{{Curso}}`, etc.).
   * Carregue uma planilha `.xlsx` com os dados.
   * Escolha a pasta de saída.
   * Configure o nome dos arquivos e opções extras (ZIP, subpastas).
   * Clique em **Gerar declarações** ou **Converter arquivos**.

---

# 🔌 Integrações e Tecnologias

* **`tkinter`**: Interface gráfica (GUI)
* **`pandas`**: Leitura de planilhas Excel
* **`python-docx`**: Manipulação de arquivos .docx
* **`docx2pdf`**, **`pdf2docx`**: Conversão entre formatos
* **`reportlab`**: Geração de relatórios PDF
* **`psutil`**, **`schedule`**, **`sqlite3`**, **`threading`**, **`logging`**: gerenciamento e desempenho
* **`pytesseract`** e **`Pillow`**: OCR de PDFs escaneados

---

# 📂 Estrutura sugerida

```bash
.
├── Prototipo_automacao.py   # Código-fonte principal
├── config_v4.json           # Configurações persistentes
├── execution_history.db     # Banco de dados SQLite
├── requirements.txt         # Dependências
├── /templates               # Modelos Word (.docx)
└── /output                  # Arquivos gerados
```

---

# 📄 Versão Executável (.exe)

Um arquivo `.exe` está disponível para facilitar o uso sem necessidade de instalar Python ou dependências.

🔗 **Link para download**: *(em breve / adicionar aqui quando hospedado)*

---

# 🎓 Aprendizados e Contexto

Esse projeto foi criado como um **projeto pessoal**, enquanto eu trabalhava como **estagiário na UPE**, na **secretaria de Mestrado e Doutorado do campus Mata Norte**.

Fui aprendendo aos poucos, pesquisando, testando, errando e corrigindo. Utilizei:

* **IAs (ferramentas de mensagens)** como assistentes
* **Vídeos tutoriais** e aulas online
* **Livros e documentações técnicas**

Aprendi muito com isso, tanto sobre código quanto sobre organização de sistemas reais. Ainda tem **muito o que melhorar**, mas já me orgulho do que consegui construir.

---

# 📅 Recursos futuros sugeridos

* Suporte a mais formatos (ODT, HTML)
* Sincronização com nuvem
* Exportação direta para e-mail
* Detecção e tradução automática de campos

---

# 💼 Licença

Este projeto é **open-source** e pode ser usado por qualquer pessoa.

> Apenas **mantenha uma citação ou crédito** ao autor original, por respeito ao trabalho.

---

> Desenvolvido para automatizar tarefas repetitivas de geração e gestão de documentos em ambientes acadêmicos e administrativos. Com foco em desempenho, confiabilidade e facilidade de uso.
