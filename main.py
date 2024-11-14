import pandas as pd
import os

# Caminho para a área de trabalho
desktop_path = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")

# Caminhos dos arquivos
comissao_path = os.path.join(desktop_path, "relatorio_comissao_vendedores.xlsx")
extrato_path = os.path.join(desktop_path, "relatorio_extrato.xlsx")
repasses_path = os.path.join(desktop_path, "profissionais_nao_eme.xlsx")
amplimed_path = os.path.join(desktop_path, "Amplimed - Gestão de Clínicas.csv")
contas_pagar_path = os.path.join(desktop_path, "relatorio_contas_pagar.xlsx")
funcionarios_path = os.path.join(desktop_path, "relatorio_funcionarios.xlsx")
comissao_modificado_path = os.path.join(desktop_path, "relatorio_comissao_vendedores_modificado.xlsx")

# Função para logar contagem de profissionais
def log_professional_counts(df, profissional_column, source_name):
    if profissional_column in df.columns:
        counts = df[profissional_column].value_counts(dropna=True)
        print(f"\nContagem de registros por profissional em {source_name}:")
        for profissional, count in counts.items():
            print(f"  {profissional}: {count} registros")
    else:
        print(f"\nA coluna '{profissional_column}' não foi encontrada em {source_name}.")

# Carregar lista de funcionários
df_funcionarios = pd.read_excel(funcionarios_path, skiprows=1)
funcionarios_list = [nome.strip().upper() for nome in df_funcionarios['Nome'].dropna().unique()]
print(f"Total de funcionários carregados: {len(funcionarios_list)}")

# Processar o relatório de comissão e criar o arquivo modificado
df_comissao = pd.read_excel(comissao_path, skiprows=1)
df_comissao.loc[:, 'Código'] = df_comissao['Código'].str.strip()

profissional = None
df_comissao.loc[:, 'Profissional'] = None

for index in range(len(df_comissao) - 1, -1, -1):
    if isinstance(df_comissao.at[index, 'Código'], str) and not df_comissao.at[index, 'Código'].isdigit():
        profissional = df_comissao.at[index, 'Código']
    elif profissional:
        df_comissao.loc[index, 'Profissional'] = profissional

df_comissao = df_comissao[df_comissao['Código'].apply(lambda x: str(x).isdigit())]
log_professional_counts(df_comissao, 'Profissional', 'relatorio_comissao_vendedores')

# Carregar o extrato e realizar a junção
df_extrato = pd.read_excel(extrato_path, skiprows=1)
if 'Descrição' in df_extrato.columns and 'Forma de pagamento' in df_extrato.columns:
    df_completo = pd.merge(df_comissao, df_extrato[['Descrição', 'Forma de pagamento']], on='Descrição', how='left')
    df_completo.to_excel(comissao_modificado_path, index=False)
    print("Arquivo modificado salvo em:", comissao_modificado_path)

# Função para extrair colunas por índices com validação
def extract_columns_by_index(df, indices):
    max_index = df.shape[1] - 1
    valid_indices = [idx for idx in indices if idx <= max_index]

    if not valid_indices:
        print(f"Erro: Índices fora dos limites. Total de colunas disponíveis: {max_index + 1}")
        return pd.DataFrame()

    try:
        return df.iloc[:, valid_indices]
    except Exception as e:
        print(f"Erro ao extrair colunas: {e}")
        return pd.DataFrame()

# Ler o arquivo Amplimed com delimitador correto
df_amplimed = pd.read_csv(amplimed_path, sep=';', encoding='utf-8')
print(f"Estrutura do Amplimed: {df_amplimed.shape}")
amplimed_data = extract_columns_by_index(df_amplimed, [4, 10, 12, 16, 18, 25, 26])
if not amplimed_data.empty:
    amplimed_data.columns = ['Data de Atendimento', 'Nome do Procedimento', 'Convênio', 'Nome do Paciente', 'Nome do Profissional', 'Modo de Pagamento', 'Valor Total']

# Ler o arquivo Profissionais Não EME ajustando leitura
df_prof_nao_eme = pd.read_excel(repasses_path)
print(f"Estrutura do Profissionais Não EME: {df_prof_nao_eme.shape}")
prof_nao_eme_data = extract_columns_by_index(df_prof_nao_eme, [1, 4, 5, 3, 11, 9])
if not prof_nao_eme_data.empty:
    prof_nao_eme_data.columns = ['Data de Atendimento', 'Nome do Procedimento', 'Convênio', 'Nome do Paciente', 'Nome do Profissional', 'Valor Total']

# Carregar e extrair dados do Contas a Pagar
df_contas_pagar = pd.read_excel(contas_pagar_path, skiprows=1)
print(f"Estrutura do Contas a Pagar: {df_contas_pagar.shape}")
contas_pagar_data = extract_columns_by_index(df_contas_pagar, [10, 4, 5, 3, 1, 18])
if not contas_pagar_data.empty:
    contas_pagar_data.columns = ['Data de Atendimento', 'Nome do Procedimento', 'Modo de Pagamento', 'Nome do Paciente', 'Nome do Profissional', 'Valor Total']

# Carregar o relatório de comissão modificado
df_comissao_modificado = pd.read_excel(comissao_modificado_path)
print(f"Estrutura do Comissão Modificado: {df_comissao_modificado.shape}")
comissao_modificado_data = extract_columns_by_index(df_comissao_modificado, [4, 3, 2, 1, 9, 10, 7])
if not comissao_modificado_data.empty:
    comissao_modificado_data.columns = ['Data de Atendimento', 'Nome do Procedimento', 'Convênio', 'Nome do Paciente', 'Nome do Profissional', 'Modo de Pagamento', 'Valor Total']

# Consolidar todos os dados extraídos
consolidated_data = pd.concat([amplimed_data, prof_nao_eme_data, contas_pagar_data, comissao_modificado_data], ignore_index=True)

# Filtrar apenas os profissionais na lista de funcionários
consolidated_data.loc[:, 'Nome do Profissional'] = consolidated_data['Nome do Profissional'].str.strip().str.upper()
filtered_data = consolidated_data[consolidated_data['Nome do Profissional'].isin(funcionarios_list)]

# Limpeza e formatação dos dados
filtered_data.loc[:, 'Convênio'] = filtered_data['Convênio'].fillna('Particular')
filtered_data = filtered_data.dropna(subset=['Data de Atendimento'])
filtered_data.loc[:, 'Data de Atendimento'] = pd.to_datetime(filtered_data['Data de Atendimento'], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y')
filtered_data.loc[:, 'Valor Total'] = filtered_data['Valor Total']

# Salvar o relatório consolidado
output_path = os.path.join(desktop_path, 'Relatorio_Consolidado.xlsx')
filtered_data.to_excel(output_path, index=False)
print("Relatório consolidado salvo em:", output_path)
