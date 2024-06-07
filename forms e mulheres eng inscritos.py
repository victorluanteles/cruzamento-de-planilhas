import pandas as pd

# Caminho para o arquivo Excel
caminho_arquivo1 = 'C:\\Users\\victor.oliveira\\Downloads\\Mulheres na Engenharia 2024 (2).xlsx'
caminho_arquivo2 = 'C:\\Users\\victor.oliveira\\Downloads\\Base de Dados Inscritos_Stellantis Mulheres na Engenharia 2024 (4).xlsx'

# Lendo o arquivo Excel e transformando em um DataFrame
dados_forms = pd.read_excel(caminho_arquivo1)
dados_inscritos = pd.read_excel(caminho_arquivo2)
def formatar_valor(valor):
    # Remover pontos e traços
    valor_sem_pontos = str(valor).replace(".", "").replace("-", "").replace("+", "")
    # Converter para número
    valor_numerico = float(valor_sem_pontos)
    # Formatar como string com zeros à esquerda
    valor_formatado = f"{int(valor_numerico):011d}"
    return valor_formatado

dados_forms.rename(columns={"Qual o seu nome completo?": 'NOMEFORMS'}, inplace=True)
dados_forms.rename(columns={"Qual seu CPF? ": 'CPFFORMS'}, inplace=True)
dados_forms.rename(columns={"Valorizamos o direito de cada indivíduo de expressar sua identidade de maneira autêntica. Embora estejamos interessados em conhecê-lo(a) melhor, é importante destacar que quaisquer informações forneci": 'Genero'}, inplace=True)

dados_forms['CPF_formatado1'] = dados_forms['CPFFORMS'].apply(formatar_valor)
dados_forms['NOME_arrumado1'] = dados_forms['NOMEFORMS'].str.strip().str.upper()

dados_inscritos['CPF_formatado2'] = dados_inscritos['CPF'].apply(formatar_valor)
dados_inscritos['NOME_arrumado2'] = dados_inscritos['NOME'].str.strip().str.upper()

dados_forms.rename(columns={"Qual o complemento? ": 'Complemento'}, inplace=True)
dados_forms.rename(columns={"Estado de residência?": 'Estado'}, inplace=True)

colunas_desejadas_transporte = ["NOMEFORMS", "CPF_formatado1", "NOME_arrumado1", "Hora de conclusão",
                                "Qual o seu país?", "Endereço de residência?", "Complemento","Cidade de residência?","Estado", "E qual o número da sua residência?","E o seu bairro de residência?","Qual a sua data de nascimento?","Você possui algum tipo de deficiência?","Se sim, descreva sua deficiência.","Qual seu grau de escolaridade?", "CPFFORMS","Qual o nome completo da mãe?","Qual seu documento de identidade?","Genero","Qual o seu estado civil?"]

dados_selecionados_transporte = dados_forms[colunas_desejadas_transporte]

colunas_desejadas_inscritos = ["CPF","NOME","CPF_formatado2","NOME_arrumado2","CURSO","DG | AVALIAÇÃO","EMAIL","ETNIA","CEP","TELEFONE"]
dados_selecionados_inscritos = dados_inscritos[colunas_desejadas_inscritos]

dados_completos_join_nome = pd.merge(dados_selecionados_transporte, dados_selecionados_inscritos, left_on='NOME_arrumado1', right_on='NOME_arrumado2', how='left')

dados_completos = pd.merge(dados_completos_join_nome, dados_selecionados_inscritos, left_on='CPF_formatado1', right_on='CPF_formatado2', how='left')

print("teste")
print(dados_completos)

nome_arquivo_filtrado = "teste_mulheres.xlsx"
dados_completos.to_excel(nome_arquivo_filtrado, index=False)


dados_completos['NOME_z'] = dados_completos['NOME_x'].fillna(dados_completos['NOME_y'])
dados_completos['CURSO_z'] = dados_completos['CURSO_x'].fillna(dados_completos['CURSO_y'])
dados_completos['DG | AVALIAÇÃO_z'] = dados_completos['DG | AVALIAÇÃO_x'].fillna(dados_completos['DG | AVALIAÇÃO_y'])
dados_completos['EMAIL_z'] = dados_completos['EMAIL_x'].fillna(dados_completos['EMAIL_y'])
dados_completos['ETINIA_z'] = dados_completos['ETNIA_x'].fillna(dados_completos['ETNIA_y'])
dados_completos['CEP_z'] = dados_completos['CEP_x'].fillna(dados_completos['CEP_y'])
dados_completos['TELEFONE_z'] = dados_completos['TELEFONE_x'].fillna(dados_completos['TELEFONE_y'])

colunas_desejadas = ["NOMEFORMS","EMAIL_z","Qual o seu país?", "Endereço de residência?", "Complemento", "Estado", "E qual o número da sua residência?","E o seu bairro de residência?","Cidade de residência?","CEP_z","Qual a sua data de nascimento?","ETINIA_z","Você possui algum tipo de deficiência?","Se sim, descreva sua deficiência.","Qual seu grau de escolaridade?", "CPFFORMS","Qual o nome completo da mãe?","Qual seu documento de identidade?","TELEFONE_z","Genero","Qual o seu estado civil?","Hora de conclusão","NOME_z","EMAIL_z","DG | AVALIAÇÃO_z"]
dados_selecionados = dados_completos[colunas_desejadas]

#nome_arquivo1 = "extracao.xlsx"
#dados_inscritos.to_excel(nome_arquivo1, index=False)
#print(f"Dados salvos em '{nome_arquivo1}'")

# Extrair apenas os valores "Aprovado(a)" ou "Stand by" do campo "DG | AVALIAÇÃO_z"
dados_filtrados = dados_selecionados[dados_selecionados['DG | AVALIAÇÃO_z'].isin(['Aprovado(a)', 'Stand by'])]

# Salvar os dados filtrados em um novo arquivo Excel
nome_arquivo_filtrado = "extracaoexcel\\dados_filtrados_mulheres.xlsx"
dados_filtrados.to_excel(nome_arquivo_filtrado, index=False)

print(f"Dados filtrados salvos em '{nome_arquivo_filtrado}'")

nome_arquivo3 = "extracaoexcel\\extracao_completos_forms_mulheres.xlsx"
dados_selecionados.to_excel(nome_arquivo3, index=False)
print(f"Dados salvos em '{nome_arquivo3}'")

