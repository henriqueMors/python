import pandas as pd

# Carregar o arquivo Excel
arquivo = "python/comparar_planilhas/SISSER.xlsx"
atualizar = pd.read_excel(arquivo, sheet_name="atualizar")
atualizado = pd.read_excel(arquivo, sheet_name="atualizado")

# Mostrar as colunas para verificar os nomes
print("Colunas disponíveis na planilha 'atualizar':", atualizar.columns)

# Remover espaços nos nomes das colunas
atualizar.columns = atualizar.columns.str.strip()
atualizado.columns = atualizado.columns.str.strip()

# Acessar a coluna correta
nomes_atualizar = atualizar["codigo"]
nomes_atualizado = atualizado["codigo"]

# Encontrar os nomes em comum
nomes_em_comum = nomes_atualizar[nomes_atualizar.isin(nomes_atualizado)]

# Criar a planilha "comparativo"
comparativo = pd.DataFrame(nomes_em_comum, columns=["codigo"])

# Encontrar os nomes que estão em "atualizado" mas não em "atualizar"
nomes_nao_localizados = nomes_atualizado[~nomes_atualizado.isin(nomes_atualizar)]

# Criar a planilha "nlocalizado"
nlocalizado = pd.DataFrame(nomes_nao_localizados, columns=["codigo"])

# Salvar o resultado no Excel (sobrescrevendo ou criando novas abas)
with pd.ExcelWriter(arquivo, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    comparativo.to_excel(writer, sheet_name="comparativo", index=False)
    nlocalizado.to_excel(writer, sheet_name="nlocalizado", index=False)

print(f"Comparação concluída! Planilha 'comparativo' salva com {len(comparativo)} nomes.")
print(f"Planilha 'nlocalizado' criada com {len(nlocalizado)} nomes que estão em 'atualizado', mas não em 'atualizar'.")
