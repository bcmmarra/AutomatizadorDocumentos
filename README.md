# 📄 Automatizador de Documentos Word (DOCX) via Excel

Este projeto em Python é uma solução robusta para a geração automatizada de múltiplos documentos `.docx`. Ele utiliza uma única planilha Excel (`dados_documentos.xlsx`) como fonte de dados para preencher dinamicamente variáveis (*placeholders*) em diversos templates Word, utilizando a biblioteca `docxtpl`.

O sistema inclui uma funcionalidade de **sincronização automática**, garantindo que as variáveis em todos os templates estejam sempre mapeadas como colunas na sua planilha de dados.

## 🚀 Como Iniciar o Projeto

Siga estes passos para configurar e executar o automatizador em seu ambiente.

### 1\. Pré-requisitos

Certifique-se de ter o Python instalado (versão 3.6 ou superior).

### 2\. Configuração e Instalação de Dependências

É **obrigatório** utilizar um ambiente virtual (`venv`) para gerenciar as dependências do projeto de forma isolada.

```bash
# 1. Crie o ambiente virtual
python -m venv .venv

# 2. Ative o ambiente virtual
# No Windows (PowerShell):
.venv\Scripts\Activate

# Em Linux/macOS:
source .venv/bin/activate

# 3. Instale as dependências listadas no requirements.txt
pip install -r requirements.txt
```

### 3\. Estrutura de Diretórios

O script foi projetado para operar com a seguinte estrutura de pastas. Todos os diretórios são criados automaticamente se não existirem, exceto `dados` e `modelos`.

```
AUTOMATIZADORDOCUMENTOS/
├── .venv/                              # Ambiente virtual (ignorar no Git)
├── dados/
│   └── dados_documentos.xlsx           # ⬅️ FONTE DE DADOS PRINCIPAL
├── documentos_gerados/                 # ⬅️ PASTA DE SAÍDA (Documentos finais)
├── modelos/                            # ⬅️ PASTA RAIZ dos templates Word
│   └── modelosEdital/                  # Exemplo de Subpasta
│   │   └── CARTA PROPOSTA.docx         # Template
├── gerarDocumentos.py                  # Script principal da automação
└── requirements.txt                    # Lista de dependências Python
```

### 4\. Configuração da Planilha (`dados_documentos.xlsx`)

A primeira linha da planilha deve conter os cabeçalhos que correspondem aos *placeholders* nos seus templates Word (`{{NOME_DA_VARIAVEL}}`).

**Colunas Essenciais de Controle:**

As seguintes colunas são **obrigatórias** e determinam o comportamento do sistema e o nome do arquivo de saída:

| Nome no Excel | Coluna no Código | Função | Exemplo de Valor |
| :--- | :--- | :--- | :--- |
| **NOME\_DO\_MODELO** | `COLUNA_TEMPLATE` | **Caminho relativo** do template a ser usado, a partir da pasta `modelos/`. Suporta subpastas. | `modelosEdital/CARTA PROPOSTA.docx` |
| **CLIENTE** | `COLUNA_NOME_CLIENTE` | Nome da entidade (Usado na nomenclatura do arquivo de saída). | `Policia Militar de MG` |
| **DOCUMENTO** | `COLUNA_NOME_DOCUMENTO` | Tipo ou Título do documento (Usado na nomenclatura do arquivo de saída). | `CARTA_PROPOSTA_2025` |
| **NUMERO\_PREGAO** | `COLUNA_NUMERO_PREGAO` | Código de referência opcional (Usado na nomenclatura do arquivo de saída). | `9003_2025` |

**Colunas de Dados:**

  * Quaisquer outras colunas na sua planilha serão usadas como contexto para preencher os *placeholders* correspondentes nos templates (ex: `VALOR_PROPOSTA`, `DATA_ASSINATURA`).

> **REMOÇÃO DE LINHAS VAZIAS:** O script ignora e não processa automaticamente as linhas onde a coluna **`NOME_DO_MODELO`** estiver vazia, garantindo que apenas registros com um template definido sejam processados.

## ⚙️ Funcionalidades e Execução

### Sincronização Automática de Colunas

Antes de gerar os documentos, o script varre todos os arquivos `.docx` na pasta `modelos/` e compara suas variáveis (*placeholders*) com as colunas existentes na `dados_documentos.xlsx`.

  * **Se uma nova variável for encontrada:** A coluna correspondente é **adicionada automaticamente** à planilha e preenchida com o valor padrão `N/A`. O script, então, para e solicita que o usuário preencha o novo campo no Excel antes de executar novamente.
  * Isso garante que nunca haja variáveis não preenchidas (*missing placeholders*) durante a renderização.

### Execução do Script

Execute o script principal diretamente do terminal (com o ambiente virtual ativado):

```bash
python gerarDocumentos.py
```

O console exibirá o progresso, indicando quais documentos estão sendo gerados e tratando quaisquer erros (como templates não encontrados ou problemas de renderização) de forma robusta.

### Padronização do Nome do Arquivo de Saída

Para manter a organização, o nome do arquivo final é construído de forma padronizada:

**`<DOCUMENTO>_<CLIENTE>_<NUMERO_PREGAO>.docx`**

**Tratamento de Caracteres:**
Uma função de limpeza é aplicada a cada segmento (`DOCUMENTO`, `CLIENTE`, `NUMERO_PREGAO`) para remover caracteres inválidos em nomes de arquivo (como `/`, `\` e `.` ), substituindo-os por *underscore* (`_`).

**Omissão de Campos Vazios:**
Se o campo **`NUMERO_PREGAO`** for o valor padrão vazio (`N/A`), ele será automaticamente omitido da nomenclatura final, mantendo o nome do arquivo limpo e conciso.
