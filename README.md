<div align="center">

<img src="assets/logo.png" alt="LIDER LIMPE" width="140"/>

# Separador de Relatório VA por Contrato

**App Streamlit para LIDER LIMPE** — sobe a planilha de Vale Alimentação (VA),
identifica o **CONTRATO** de cada colaborador pelo posto de trabalho e exporta
um **ZIP** com uma planilha `.xlsx` por contrato no formato:

```
CONTRATO      MM-AAAA.xlsx
```

[![Streamlit](https://img.shields.io/badge/Streamlit-Cloud-FF4B4B?logo=streamlit&logoColor=white)](https://streamlit.io/cloud)
[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/Uso-Interno%20LIDER%20LIMPE-1B3A8C)]()

</div>

---

## ✨ O que o app faz

1. Recebe a planilha **Relatório VA** (`.xls` ou `.xlsx`) gerada pelo sistema.
2. Remove automaticamente:
   - Linha de cabeçalho original
   - **Linhas separadoras** de posto (`Posto de Trabalho: XXX`)
   - **Linhas de subtotal**
   - **Rodapé** (Período, Empresa, Código/Mapa, Total/Geral, Data/Hora de impressão)
3. Cruza cada **POSTO TRABALHO** com a planilha **Mapeamento Sistema**
   (aba `POSTOS`, colunas `Nome_EasyApp → CONTRATO`).
4. Gera, para cada contrato, uma planilha `.xlsx` com **apenas três colunas**:
   `NOME | CPF | VALOR` (mais uma linha de **TOTAL**).
5. Empacota tudo em um único **ZIP**, com nome:

   ```
   RELATORIO VA POR CONTRATO MM-AAAA.zip
   └── CONTRATO_A      MM-AAAA.xlsx
   └── CONTRATO_B      MM-AAAA.xlsx
   └── ...
   ```

> 💡 **Mês/Ano**: por padrão o app sugere o **mês seguinte** ao atual,
> porque o VA é referente ao próximo mês (ex.: relatório gerado em junho =
> arquivos com sufixo `07-AAAA`). Você pode trocar manualmente na barra lateral.

---

## 📂 Estrutura do projeto

```
lider-va-splitter/
├── app.py                    ← UI Streamlit (sobe esse arquivo no Streamlit Cloud)
├── core.py                   ← Lógica de leitura/limpeza/exportação (puro Python)
├── requirements.txt
├── README.md
├── .gitignore
├── .gitattributes
├── .streamlit/
│   └── config.toml           ← Tema visual (azul/laranja LIDER LIMPE)
├── assets/
│   └── logo.png              ← Logo exibido na barra lateral e no topo
└── data/
    └── Mapeamento Sistema.xls   ← Mapeamento POSTO→CONTRATO versionado
```

---

## 🚀 Como rodar

### Local (Windows / Mac / Linux)

```bash
# 1. Clone o repositório
git clone https://github.com/SEU_USUARIO/lider-va-splitter.git
cd lider-va-splitter

# 2. (opcional) crie um ambiente virtual
python -m venv .venv
.\.venv\Scripts\activate          # Windows
# source .venv/bin/activate       # Mac / Linux

# 3. Instale as dependências
pip install -r requirements.txt

# 4. Suba o app
streamlit run app.py
```

O app abre em `http://localhost:8501`.

### Streamlit Community Cloud (recomendado)

1. Faça **push** deste repositório para o seu GitHub.
2. Acesse <https://share.streamlit.io> → **New app**.
3. Aponte para o repositório, branch `main`, arquivo `app.py`.
4. Pronto — o app fica disponível em uma URL pública.

> Como o `Mapeamento Sistema.xls` está versionado em `data/`, ao subir uma
> versão nova no GitHub o app passa a usar a nova versão automaticamente.

---

## 🗂️ Atualizando o Mapeamento Sistema

Tem duas formas:

### A) Atualizar pelo GitHub (recomendado, fica versionado)

1. Edite o arquivo `data/Mapeamento Sistema.xls` localmente (ou substitua por uma nova versão).
2. **Commit + push** para o GitHub.
3. O app, no próximo carregamento, já lê a nova versão.

### B) Subir manualmente na hora

1. Na barra lateral, desligue **“Usar Mapeamento do repositório”**.
2. Use o uploader para enviar o `.xls`/`.xlsx` daquela execução.

> ⚠️ O mapeamento **precisa ter a aba `POSTOS`** com as colunas
> `Nome_EasyApp` e `CONTRATO`. As demais colunas/abas são ignoradas.

---

## 🧾 Formato esperado da Planilha VA de entrada

A planilha precisa seguir o layout do sistema (igual ao `MODELO_TESTE.xls`):

| Coluna A   | Coluna B         | Coluna C        | Coluna D            | Coluna E  | Coluna F     |
|------------|------------------|-----------------|---------------------|-----------|--------------|
| **NOME**   | **CPF**          | **POSTO TRABALHO** | **LOTE**         | **VALOR** | _(opcional: obs)_ |
| ANE CAROLINE…  | 161.422.807-85   | ADM - LIDER LIMPE | Não informado | 525,14   |              |
| …          | …                | …                 | …             | …        |              |

Entre cada bloco de posto há uma linha tipo `Posto de Trabalho: XXX` e
no final do bloco uma linha de `Subtotal:` — ambas são **descartadas
automaticamente**, junto com o rodapé do relatório.

---

## 🧪 Resultado gerado

Cada arquivo `.xlsx` dentro do ZIP tem **exatamente** este formato:

| NOME                       | CPF         | VALOR     |
|----------------------------|-------------|-----------|
| GABRIEL ALMEIDA SANTOS     | 12854998758 | R$ 628,30 |
| LUCIANO KENG QUEIROZ       | 01716129737 | R$ 505,48 |
| SAMUEL SOARES DE CARVALHO  | 18271622790 | R$ 628,30 |
| …                          | …           | …         |
| **TOTAL**                  |             | **R$ 70.206,36** |

- `CPF` salvo como **texto** (preserva zero à esquerda).
- `VALOR` formatado como **moeda BRL** com separadores nativos do Excel.

---

## ⚠️ Postos sem contrato

Se algum posto da planilha de entrada **não tiver** correspondência na coluna
`Nome_EasyApp` do Mapeamento (ou estiver lá com `CONTRATO` vazio), o app:

- Exibe um **alerta** com a lista dos postos não encontrados.
- Agrupa esses colaboradores em um único arquivo **`SEM_CONTRATO      MM-AAAA.xlsx`**.
- Você pode optar por **não incluir** esse arquivo no ZIP, no checkbox final.

> Recomendado: ao ver postos no alerta, **atualize o Mapeamento no GitHub**
> e rode o app de novo.

---

## 🛠️ Stack técnica

| Camada       | Tecnologia                            |
|--------------|---------------------------------------|
| UI           | [Streamlit](https://streamlit.io/)    |
| Parsing      | `pandas`, `openpyxl`, `xlrd` (legacy .xls) |
| Geração XLSX | `openpyxl` (estilos + formatação)     |
| Empacotamento| `zipfile` (stdlib)                    |

---

## 📜 Licença

Uso interno **LIDER LIMPE**. Não redistribuir.

---

<div align="center">

Desenvolvido com 💙 + 🧡 para a equipe LIDER LIMPE.

</div>
