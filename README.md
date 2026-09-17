# Gerador de Termo de Responsabilidade (TR)

Este script Python automatiza a geração de documentos do Word (`.docx`) a partir de dados estruturados em uma planilha do Excel (`.xlsx`). Ele lê as informações dos colaboradores/equipamentos, preenche um modelo predefinido (substituindo textos no corpo do documento e em tabelas) e salva os arquivos organizados automaticamente em pastas separadas por filial.

## 🛠️ Tecnologias e Bibliotecas Utilizadas

* **Python 3.x**

* [**Pandas**](https://pandas.pydata.org/): Para leitura e manipulação dos dados da planilha Excel.

* [**python-docx**](https://python-docx.readthedocs.io/): Para manipulação e substituição de textos no modelo do Word.

* [**openpyxl**](https://openpyxl.readthedocs.io/): Dependência necessária do Pandas para leitura de arquivos `.xlsx`.

## 📋 Pré-requisitos

Certifique-se de ter o Python instalado em sua máquina e instale as dependências necessárias executando no terminal:

```
pip install pandas python-docx openpyxl

```

## 📁 Estrutura de Arquivos Recomendada

O script utiliza **caminhos dinâmicos**, portanto basta manter os arquivos de entrada na mesma pasta do script `gerador_tr.py`:

```
📂 Seu Projeto/
├── gerador_tr.py       # Script principal
├── Base.xlsx          # Planilha com os dados dos colaboradores/equipamentos
├── MODELO.docx         # Modelo base do Termo de Responsabilidade
└── 📂 Termos_Gerados/  # (Criado automaticamente pelo script)
    ├── 📂 Filial_A/
    │   └── Termo_de_Responsabilidade_João_Silva.docx
    └── 📂 Filial_B/
        └── Termo_de_Responsabilidade_Maria_Santos.docx

```

## 📊 Estrutura da Planilha (`Base.xlsx`)

A planilha precisa conter as seguintes colunas (com a grafia exata):

| Coluna | Descrição | 
 | ----- | ----- | 
| `Nome` | Nome do colaborador (usado também no nome do arquivo) | 
| `Cargo` | Cargo ou função do colaborador | 
| `Cpf` | CPF (tratado como texto para preservar zeros à esquerda e pontuação) | 
| `Cnpj` | CNPJ da empresa | 
| `Modelo` | Modelo do equipamento | 
| `Tag` | Identificador / Patrimônio / Serial do equipamento | 
| `Tipo` | Tipo de dispositivo (ex: Notebook, Smartphone) | 
| `Filial` | Nome da filial (usada para organizar os arquivos em subpastas) | 
| `Estado` | Unidade Federativa (UF) | 
| `Endereco` | Endereço da filial ou colaborador | 
| `Numero` | Número | 
| `Bairro` | Bairro | 
| `CEP` | Código de Endereçamento Postal | 

## 📄 Formatação do Modelo Word (`MODELO.docx`)

O arquivo `MODELO.docx` deve conter marcadores de posição (*placeholders*) que correspondem às chaves que o script irá substituir:

* `<<Nome>>`

* `<<Cargo>>`

* `<<Cpf>>`

* `<<Cnpj>>`

* `<<Modelo>>`

* `<<Tag>>`

* `<<Tipo>>`

* `<<Filial>>`

* `<<Estado>>`

* `<<Endereco>>`

* `<<Numero>>`

* `<<Bairro>>`

* `<<CEP>>`

*Nota: Os marcadores funcionam tanto no corpo normal do texto (parágrafos) quanto dentro de tabelas.*

## ⚙️ Principais Funcionalidades do Script

1. **Caminhos Dinâmicos**: O script utiliza `os.path.dirname(os.path.abspath(__file__))` para localizar os arquivos automaticamente, funcionando em qualquer sistema operacional ou computador sem necessidade de alterar caminhos fixos.

2. **Higienização de Nomes de Arquivo**: Possui a função `limpar_nome_arquivo()` que remove caracteres proibidos pelo sistema operacional (como `\ / : * ? " < > |`) dos campos `Nome` e `Filial`, prevenindo erros durante o salvamento.

3. **Substituição Abrangente**: Varre parágrafos comuns e células de tabelas dentro do documento `.docx`.

4. **Organização por Pastas**: Cria automaticamente a pasta `Termos_Gerados` e subpastas para cada filial encontrada na planilha.

## 🚀 Como Executar

Execute o comando no terminal na pasta do projeto:

```
python gerador_tr.py

```

Os documentos preenchidos serão exibidos no terminal à medida que forem sendo gerados e ficarão salvos dentro da pasta `Termos_Gerados/`.
