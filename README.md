<div align="center">

<img src="assets/logo.png" alt="LIDER LIMPE" width="140"/>

# Separador de Benefícios por Contrato

**App Streamlit da LIDER LIMPE** — envie a planilha de benefícios, o app
identifica o **CONTRATO** de cada colaborador pelo posto de trabalho e gera
um **ZIP** com uma planilha `.xlsx` por contrato no formato:

```
CONTRATO      MM-AAAA.xlsx
```

[![Streamlit](https://img.shields.io/badge/Streamlit-Cloud-FF4B4B?logo=streamlit&logoColor=white)](https://streamlit.io/cloud)
[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Uso](https://img.shields.io/badge/Uso-Interno%20LIDER%20LIMPE-1B3A8C)]()

</div>

---

## ✨ O que o app faz

1. Recebe a planilha enviada (`.xls` ou `.xlsx`) e **detecta o formato
   automaticamente**:
   - 📄 **Relatório VA clássico** — remove cabeçalho extra, linhas separadoras
     (`Posto de Trabalho: XXX`), subtotais e o rodapé (Período, Empresa,
     Código/Mapa, Total/Geral, Data/Hora).
   - 📊 **Extrato Mapa Benefícios** (VA + CF + CB consolidado em um único
     valor) — leitura direta das colunas `Nome do Funcionário | CPF/CNPJ |
     Posto de Trabalho | Total Benefício (R$)`, removendo a linha
     `TOTAL GERAL:`.
2. Cruza cada **Posto de Trabalho** com o **Mapeamento de Postos**
   (`Nome_EasyApp → CONTRATO`).
3. Gera, para cada contrato, um `.xlsx` com `NOME | CPF | VALOR` + linha
   de **TOTAL**.
   - O arquivo **`SEM_CONTRATO`** (postos ainda não cadastrados no
     mapeamento) sai com uma coluna extra: `NOME | CPF | POSTO | VALOR`,
     pra você identificar exatamente quais postos cadastrar.
4. Empacota tudo num único **ZIP**:

```
RELATORIO VA POR CONTRATO MM-AAAA.zip
├── SALVAMAR      07-2026.xlsx
├── CESAN         07-2026.xlsx
├── SEM_CONTRATO  07-2026.xlsx
└── ...
```

> 💡 **Mês/Ano**: o app sugere automaticamente o **mês seguinte** ao atual,
> pois o benefício é referente ao próximo mês. Dá pra trocar na barra lateral.

---

## 🗺️ Mapeamento de Postos (3 fontes)

Na barra lateral você escolhe de onde vem o mapeamento `Posto → Contrato`:

| Fonte | Quando usar |
|---|---|
| **Google Sheets (automático)** ⭐ padrão | Uso do dia a dia. Lê a aba **POSTOS** da planilha compartilhada — qualquer pessoa que editar lá atualiza o app para todo mundo |
| Arquivo do repositório | Fallback: usa o `data/Mapeamento Sistema.xls` versionado no GitHub (também serve como contingência automática se o Sheets estiver fora do ar) |
| Upload manual | Envio pontual de um `.xls`/`.xlsx` só para aquela execução |

### Como atualizar o mapeamento (Google Sheets)

1. Abra a planilha compartilhada e vá até a aba **POSTOS**.
2. Edite/inclua linhas: `Nome_EasyApp` = nome do posto **exatamente como
   aparece na planilha de entrada**; `CONTRATO` = nome do contrato que vai
   para o nome do arquivo.
3. Salve — o app reflete a mudança em até **1 hora** (cache) ou imediatamente
   clicando em **🔄 Atualizar mapeamento agora** na barra lateral.

> ⚠️ A planilha precisa estar compartilhada como
> **"Qualquer pessoa com o link pode ver"** (Arquivo → Compartilhar).
> Se ficar restrita, o app cai automaticamente para o arquivo do repositório
> e avisa na tela.

---

## 📂 Estrutura do projeto

```
SEPARAR-CONTRATOS-BEN/
├── app.py                       ← UI Streamlit (arquivo principal)
├── core.py                      ← Lógica (parsing, mapeamento, exportação)
├── requirements.txt
├── .python-version
├── README.md
├── .gitignore
├── .gitattributes
├── .streamlit/config.toml       ← Tema azul/laranja LIDER LIMPE
├── assets/logo.png
└── data/Mapeamento Sistema.xls  ← Fallback do mapeamento
```

---

## 🚀 Como rodar

### Local

```bash
git clone https://github.com/SEU_USUARIO/SEPARAR-CONTRATOS-BEN.git
cd SEPARAR-CONTRATOS-BEN

python -m venv .venv
.\.venv\Scripts\activate        # Windows
# source .venv/bin/activate     # Mac/Linux

pip install -r requirements.txt
streamlit run app.py
```

### Streamlit Community Cloud

1. Faça **push** do repositório para o GitHub.
2. Acesse <https://share.streamlit.io> → **New app**.
3. Aponte para o repositório, branch `main`, arquivo `app.py`.
4. Pronto — sem secrets, sem credenciais, sem configuração extra.

---

## 🧾 Formatos de entrada aceitos

### A) Relatório VA clássico

| NOME | CPF | POSTO TRABALHO | LOTE | VALOR | obs |
|---|---|---|---|---|---|
| ANE CAROLINE… | 161.422.807-85 | ADM - LIDER LIMPE | Não informado | 525,14 | |

Com linhas `Posto de Trabalho: XXX` entre blocos, `Subtotal:` ao fim de cada
bloco e rodapé de ~6 linhas (Período, Empresa, Código/Mapa, Total/Geral) —
tudo descartado automaticamente.

### B) Extrato Mapa Benefícios (VA + CF + CB)

| Nome do Funcionário | CPF/CNPJ | Posto de Trabalho | Total Benefício (R$) |
|---|---|---|---|
| ALAN VITOR… | 194.396.547-16 | ADM - LIDER LIMPE | 525,14 |

Com apenas 1 linha de rodapé (`TOTAL GERAL:`), descartada automaticamente.

---

## 📤 Resultado gerado

Contratos normais:

| NOME | CPF | VALOR |
|---|---|---|
| GABRIEL ALMEIDA SANTOS | 12854998758 | R$ 628,30 |
| **TOTAL** | | **R$ 70.206,36** |

`SEM_CONTRATO` (com a coluna POSTO para facilitar o cadastro):

| NOME | CPF | POSTO | VALOR |
|---|---|---|---|
| ALAN VITOR DOS ANJOS BISPO | 19439654716 | ADM - OPCAO | R$ 525,14 |

- `CPF` salvo como **texto** (preserva zeros à esquerda).
- `VALOR` formatado como **moeda BRL** nativa do Excel.

---

## ⚠️ Postos sem contrato

Se algum posto não existir na aba **POSTOS** do mapeamento (ou estiver sem
`CONTRATO`), o app:

- Exibe um **alerta** com a lista dos postos faltantes.
- Agrupa esses colaboradores no arquivo `SEM_CONTRATO      MM-AAAA.xlsx`
  (com a coluna POSTO).
- Permite desmarcar a inclusão desse arquivo no ZIP.

Depois de cadastrar os postos na aba POSTOS do Google Sheets, clique em
**🔄 Atualizar mapeamento agora** e reprocessa o arquivo.

---

## 🛠️ Stack

| Camada | Tecnologia |
|---|---|
| UI | [Streamlit](https://streamlit.io/) |
| Parsing | `pandas`, `openpyxl`, `xlrd` (legacy .xls) |
| Mapeamento | Google Sheets (export CSV público) ou arquivo `.xls` local |
| Geração XLSX | `openpyxl` |
| Empacotamento | `zipfile` (stdlib) |

---

## 📜 Licença

Uso interno **LIDER LIMPE**. Não redistribuir.

---

<div align="center">

Desenvolvido com 💙 + 🧡 para a equipe LIDER LIMPE.

</div>
