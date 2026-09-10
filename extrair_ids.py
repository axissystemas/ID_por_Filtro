import os
import sys
import pandas as pd

def extrair_ids(arquivo_excel="Chamados N3 SGBOM.xlsx", arquivo_saida_csv="IDs_agrupados_por_categoria.csv"):
    # Verifica se o arquivo existe no diretório atual
    diretorio_atual = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
    caminho_excel = os.path.join(diretorio_atual, arquivo_excel)
    
    if not os.path.exists(caminho_excel):
        # Tenta procurar qualquer arquivo .xlsx no diretório se o nome exato não for encontrado
        arquivos_xlsx = [f for f in os.listdir(diretorio_atual) if f.endswith('.xlsx') and not f.startswith('~$')]
        if arquivos_xlsx:
            caminho_excel = os.path.join(diretorio_atual, arquivos_xlsx[0])
            print(f"Arquivo '{arquivo_excel}' não encontrado diretamente. Usando: {arquivos_xlsx[0]}")
        else:
            print(f"Erro: Arquivo Excel não encontrado em: {caminho_excel}")
            return

    print(f"Lendo planilha: {os.path.basename(caminho_excel)}...")
    df = pd.read_excel(caminho_excel)

    if df.shape[1] < 3:
        print("Erro: A planilha deve conter pelo menos 3 colunas (Coluna A = ID, Coluna C = Categoria).")
        return

    # Extrai coluna A (índice 0) e coluna C (índice 2)
    nome_col_id = df.columns[0]
    nome_col_cat = df.columns[2]
    
    print(f"Coluna A detectada: '{nome_col_id}'")
    print(f"Coluna C detectada: '{nome_col_cat}'")

    # Tratamento e limpeza dos dados
    df_dados = pd.DataFrame({
        'ID_Original': df.iloc[:, 0].astype(str).str.strip(),
        # Remove espaços no meio do número (ex: '15 653' -> '15653') e sufixos decimais se houver
        'ID_Limpo': df.iloc[:, 0].astype(str).str.strip().str.replace(' ', '', regex=False).str.replace(r'\.0$', '', regex=True),
        'Categoria': df.iloc[:, 2].astype(str).str.strip()
    })

    # Remove linhas vazias ou nulas
    df_dados = df_dados[df_dados['Categoria'].str.lower() != 'nan']
    df_dados = df_dados[df_dados['ID_Limpo'].str.lower() != 'nan']
    df_dados = df_dados[df_dados['Categoria'] != '']
    df_dados = df_dados[df_dados['ID_Limpo'] != '']

    total_registros = len(df_dados)
    print(f"Total de chamados válidos processados: {total_registros}")

    # 1. Agrupamento por Categoria com IDs separados por vírgula
    # Para sistemas/filtros, IDs limpos separados por vírgula
    agrupado = df_dados.groupby('Categoria', sort=False).agg(
        Quantidade=('ID_Limpo', 'count'),
        IDs=('ID_Limpo', lambda ids: ', '.join(ids))
    ).reset_index()

    # Ordena pelas categorias com mais chamados
    agrupado = agrupado.sort_values(by='Quantidade', ascending=False).reset_index(drop=True)

    # Caminho do CSV principal
    caminho_saida = os.path.join(diretorio_atual, arquivo_saida_csv)
    # Salva com utf-8-sig para abrir perfeitamente com acentuação no Excel
    agrupado.to_csv(caminho_saida, index=False, sep=',', encoding='utf-8-sig')
    print(f"\n[OK] CSV agrupado gerado com sucesso em:")
    print(f"     -> {caminho_saida}")

    # 2. Também gera um CSV detalhado linha por linha (Categoria, ID) para flexibilidade
    caminho_detalhado = os.path.join(diretorio_atual, "IDs_por_categoria_detalhado.csv")
    df_dados[['Categoria', 'ID_Limpo']].rename(columns={'ID_Limpo': 'ID'}).to_csv(
        caminho_detalhado, index=False, sep=',', encoding='utf-8-sig'
    )
    print(f"[OK] CSV detalhado (linha por linha) gerado em:")
    print(f"     -> {caminho_detalhado}")

    # Exibe prévia das primeiras categorias no terminal
    print("\n--- Prévia dos resultados agrupados ---")
    for _, row in agrupado.head(5).iterrows():
        ids_amostra = row['IDs'][:60] + ('...' if len(row['IDs']) > 60 else '')
        print(f"- {row['Categoria']} ({row['Quantidade']} chamados): {ids_amostra}")

if __name__ == '__main__':
    extrair_ids()
