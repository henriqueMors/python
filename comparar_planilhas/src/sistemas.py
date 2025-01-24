import pandas as pd

# Carregar o arquivo Excel
arquivo = "python/comparar_planilhas/SISSER.xlsx"
comparativo = pd.read_excel(arquivo, sheet_name="comparativo")
atualizar = pd.read_excel(arquivo, sheet_name="atualizar")
atualizado = pd.read_excel(arquivo, sheet_name="atualizado")

# Remover espaços nos nomes das colunas
comparativo.columns = comparativo.columns.str.strip()
atualizar.columns = atualizar.columns.str.strip()
atualizado.columns = atualizado.columns.str.strip()

# Garantir que estamos lidando com os códigos corretamente
codigos_comparativo = comparativo["codigo"]

# Lista de colunas a serem comparadas (C até AY)
colunas_para_comparar = atualizar.columns[2:]  # Colunas da 3ª em diante (C até AY)

# DataFrame para armazenar as diferenças
diferencas = []

# Comparar as colunas para os códigos em "comparativo"
for codigo in codigos_comparativo:
    # Filtrar as linhas correspondentes ao código em cada planilha
    linha_atualizar = atualizar[atualizar["codigo"] == codigo]
    linha_atualizado = atualizado[atualizado["codigo"] == codigo]

    # Verificar se ambos os códigos existem nas planilhas
    if not linha_atualizar.empty and not linha_atualizado.empty:
        for coluna in colunas_para_comparar:
            valor_atualizar = linha_atualizar[coluna].values[0]  # Valor na planilha "atualizar"
            valor_atualizado = linha_atualizado[coluna].values[0]  # Valor na planilha "atualizado"

            # Comparar os valores
            if valor_atualizar != valor_atualizado:
                # Adicionar a diferença ao DataFrame
                diferencas.append({
                    "codigo": codigo,
                    "coluna": coluna,
                    "valor_atualizado": valor_atualizado,
                    "valor_atualizar": valor_atualizar
                })

# Converter a lista de diferenças em um DataFrame
para_atualizar = pd.DataFrame(diferencas)

# Salvar o resultado na nova planilha "paraAtualizar"
with pd.ExcelWriter(arquivo, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    para_atualizar.to_excel(writer, sheet_name="paraAtualizar", index=False)

print(f"Planilha 'paraAtualizar' criada com {len(para_atualizar)} diferenças encontradas!")
