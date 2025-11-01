import pandas as pd
from docxtpl import DocxTemplate
import os

# --- Configurações ---
NOME_TEMPLATE = 'CARTA PROPOSTA.docx'
NOME_PLANILHA = 'dados_contratos.xlsx'
PASTA_DADOS = 'dados'
PASTA_SAIDA = 'contratos_gerados'

def gerar_documentos():
    # 1. Carregar os dados usando Pandas
    caminho_planilha = os.path.join(PASTA_DADOS, NOME_PLANILHA)
    
    try:
        # Lê a planilha, preenchendo células vazias com string vazia
        df = pd.read_excel(caminho_planilha).fillna('')
    except FileNotFoundError:
        print(f"ERRO: Arquivo {caminho_planilha} não encontrado.")
        return

    # 2. Carregar o template
    try:
        doc_template = DocxTemplate(NOME_TEMPLATE)
    except FileNotFoundError:
        print(f"ERRO: Template {NOME_TEMPLATE} não encontrado na raiz do projeto.")
        return

    # Cria a pasta de saída se ela não existir
    os.makedirs(PASTA_SAIDA, exist_ok=True)
    
    contador = 0

    # 3. Processar cada linha de dados
    # A função df.to_dict('records') transforma o DataFrame em uma lista de dicionários
    for dados_contrato in df.to_dict('records'):
        
        # O DocxTemplate precisa de um dicionário (contexto) para a substituição
        context = {k: v for k, v in dados_contrato.items()}
        
        # A biblioteca DocxTemplate se encarrega de:
        # 1. Encontrar todos os {{placeholders}} no template.
        # 2. Substituí-los pelos valores do 'context'.
        # 3. *PRESERVAR* toda a formatação original do texto ao redor e do próprio placeholder.
        doc_template.render(context)

        # 4. Salvar o novo contrato
        # Usa o valor da coluna 'SOCIO' para nomear o arquivo
        nome_arquivo = f"Contrato_{dados_contrato.get('SOCIO', contador)}.docx"
        caminho_saida = os.path.join(PASTA_SAIDA, nome_arquivo)
        
        doc_template.save(caminho_saida)
        
        contador += 1
        print(f"✅ Contrato gerado: {nome_arquivo}")

    print(f"\n--- Automação Concluída ---")
    print(f"{contador} contratos gerados na pasta '{PASTA_SAIDA}'")


if __name__ == "__main__":
    gerar_documentos()