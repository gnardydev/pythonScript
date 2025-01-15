import pandas as pd
import os
import re

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
regranegocio_path = os.path.join(desktop_path, "Regras de negocio.xlsx")

# Função para logar contagem de profissionais
def log_professional_counts(df, profissional_column, source_name):
    if profissional_column in df.columns:
        counts = df[profissional_column].value_counts(dropna=True)
        print(f"\nContagem de registros por profissional em {source_name}:")
        for profissional, count in counts.items():
            print(f"  {profissional}: {count} registros")
    else:
        print(f"\nA coluna '{profissional_column}' não foi encontrada em {source_name}.")

# Adicionar logs detalhados para valores
def log_values(title, df, columns):
    print(f"\n{title}:")
    print(df[columns].head(10))

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
    amplimed_data = amplimed_data.copy()
    amplimed_data.loc[:, 'Fonte'] = 'Amplimed'

# Ler o arquivo Profissionais Não EME ajustando leitura
df_prof_nao_eme = pd.read_excel(repasses_path)
print(f"Estrutura do Profissionais Não EME: {df_prof_nao_eme.shape}")
prof_nao_eme_data = extract_columns_by_index(df_prof_nao_eme, [1, 4, 5, 3, 11, 9])
if not prof_nao_eme_data.empty:
    prof_nao_eme_data.columns = ['Data de Atendimento', 'Nome do Procedimento', 'Convênio', 'Nome do Paciente', 'Nome do Profissional', 'Valor Total']
    prof_nao_eme_data = prof_nao_eme_data.copy()
    prof_nao_eme_data.loc[:, 'Fonte'] = 'Profissionais Não EME'

# Carregar e extrair dados do Contas a Pagar
df_contas_pagar = pd.read_excel(contas_pagar_path, skiprows=1)
print(f"Estrutura do Contas a Pagar: {df_contas_pagar.shape}")
contas_pagar_data = extract_columns_by_index(df_contas_pagar, [10, 4, 5, 3, 1, 18])
if not contas_pagar_data.empty:
    contas_pagar_data.columns = ['Data de Atendimento', 'Nome do Procedimento', 'Modo de Pagamento', 'Nome do Paciente', 'Nome do Profissional', 'Valor Total']
    contas_pagar_data = contas_pagar_data.copy()
    contas_pagar_data.loc[:, 'Fonte'] = 'Contas a Pagar'

# Carregar o relatório de comissão modificado
df_comissao_modificado = pd.read_excel(comissao_modificado_path)
print(f"Estrutura do Comissão Modificado: {df_comissao_modificado.shape}")
comissao_modificado_data = extract_columns_by_index(df_comissao_modificado, [4, 3, 2, 1, 9, 10, 7])
if not comissao_modificado_data.empty:
    comissao_modificado_data.columns = ['Data de Atendimento', 'Nome do Procedimento', 'Convênio', 'Nome do Paciente', 'Nome do Profissional', 'Modo de Pagamento', 'Valor Total']
    comissao_modificado_data = comissao_modificado_data.copy()
    comissao_modificado_data.loc[:, 'Fonte'] = 'Comissão Modificada'

# Consolidar todos os dados extraídos
consolidated_data = pd.concat([amplimed_data, prof_nao_eme_data, contas_pagar_data, comissao_modificado_data], ignore_index=True)

# Log da consolidação inicial
log_values("Dados consolidados iniciais", consolidated_data, ['Fonte', 'Nome do Profissional', 'Valor Total'])

# Filtrar apenas os profissionais na lista de funcionários
consolidated_data.loc[:, 'Nome do Profissional'] = consolidated_data['Nome do Profissional'].str.strip().str.upper()
filtered_data = consolidated_data[consolidated_data['Nome do Profissional'].isin(funcionarios_list)]

# Limpeza e formatação dos dados
filtered_data.loc[:, 'Convênio'] = filtered_data['Convênio'].fillna('Particular')
filtered_data = filtered_data.dropna(subset=['Data de Atendimento'])
filtered_data.loc[:, 'Data de Atendimento'] = pd.to_datetime(filtered_data['Data de Atendimento'], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y')

# Converter 'Valor Total' de moeda para numérico
# Função para limpar e converter valores monetários corretamente
# Função para limpar e converter valores monetários corretamente
# Função para limpar e converter valores monetários corretamente por fonte
def clean_currency(value, fonte):
    try:
        value = str(value).strip()  # Remove espaços em branco

        if fonte == 'Amplimed':
            # Lógica baseada em taxas_data['Taxa']
            return float(
                value
                .replace('R$', '')  # Remove "R$"
                .replace('.', '')   # Remove separadores de milhar
                .replace(',', '.')  # Converte vírgula para ponto decimal
            )
        elif fonte == 'Profissionais Não EME':
            # Tratamento específico para Profissionais Não EME (formato brasileiro)
            if ',' in value and '.' in value:
                value = value.replace('.', '').replace(',', '.')
            elif ',' in value:
                value = value.replace(',', '.')
            return float(value)
        else:
            # Tratamento padrão para outras fontes
            if ',' in value and '.' in value:
                value = value.replace('.', '').replace(',', '.')
            elif ',' in value:
                value = value.replace(',', '.')
            return float(value)
    except ValueError:
        return 0  # Retorna 0 para valores inválidos

# Aplicar a função de limpeza na coluna 'Valor Total' com base na fonte
filtered_data['Valor Total'] = filtered_data.apply(
    lambda row: clean_currency(row['Valor Total'], row['Fonte']), axis=1
)

# Preenchendo valores nulos com 0 após limpeza
filtered_data['Valor Total'] = filtered_data['Valor Total'].fillna(0)

# Log para validar os valores corrigidos
print("\nLog - Valores corrigidos por fonte:")
print(filtered_data[['Fonte', 'Valor Total']].groupby('Fonte').head(5))


# Log após a conversão de moeda para numérico
log_values("Dados após conversão de moeda para numérico", filtered_data, ['Fonte', 'Nome do Profissional', 'Valor Total'])

print("\nLog Inicial - Valores de 'Valor Total' ao carregar 'Profissionais Não EME':")
print(filtered_data[filtered_data['Fonte'] == 'Profissionais Não EME']['Valor Total'].head(10))


# Carregar as tabelas de regras de negócio
taxas_data = pd.read_excel(regranegocio_path, sheet_name='Taxas')
rateio_data = pd.read_excel(regranegocio_path, sheet_name='Rateio por profissional')

# Converter coluna "Taxa" para numérico
taxas_data['Taxa'] = (
    taxas_data['Taxa']
    .astype(str)
    .str.replace('R$', '', regex=False)
    .str.replace('.', '', regex=False)
    .str.replace(',', '.', regex=False)
    .astype(float)
    .fillna(0)
)

# Criar nova coluna para valor subtraído
filtered_data['Valor Subtraído'] = filtered_data['Valor Total']

# Aplicar taxas proporcionais a parcelamentos e atualizar o valor subtraído para fontes não "Profissionais Não EME"
for _, row in taxas_data.iterrows():
    procedimento = row['Procedimento']
    taxa = row['Taxa']
    print(f"Aplicando taxa {taxa} para o procedimento {procedimento}")
    mask = (filtered_data['Nome do Procedimento'] == procedimento) & (filtered_data['Fonte'] != 'Profissionais Não EME') & (filtered_data['Fonte'] != 'Contas a Pagar')

    for idx in filtered_data[mask].index:
        convenio = filtered_data.at[idx, 'Convênio']
        parcel_match = re.search(r'\((\d+)/(\d+)\)', convenio)

        if parcel_match:
            parcela_atual = int(parcel_match.group(1))
            total_parcelas = int(parcel_match.group(2))
            taxa_parcela = round(taxa / total_parcelas, 2)  # Garantir 2 casas decimais
            filtered_data.at[idx, 'Valor Subtraído'] -= taxa_parcela
        else:
            filtered_data.at[idx, 'Valor Subtraído'] -= round(taxa, 2)  # Garantir 2 casas decimais

    filtered_data.loc[mask, 'Taxa Aplicada'] = round(taxa, 2)

# Log das taxas aplicadas
log_values("Dados após aplicação de taxas", filtered_data, ['Fonte', 'Nome do Profissional', 'Valor Subtraído', 'Taxa Aplicada'])


# Aplicar rateio com base no Valor Subtraído somente se a Fonte não for "Profissionais Não EME"
rateio_dict = dict(zip(rateio_data['Profissionais'].str.strip().str.upper(), rateio_data['%Profissionais']))
print(f"Rateio disponível para profissionais: {rateio_dict}")


#filtered_data.loc[filtered_data['Fonte'] == 'Profissionais Não EME', 'Valor Total'] = (
#    filtered_data.loc[filtered_data['Fonte'] == 'Profissionais Não EME', 'Valor Total']
 #   .astype(str)                           # Garante que os valores sejam strings
  #  .str.replace('.', '', regex=False)     # Remove separadores de milhar
   # .str.replace(',', '.', regex=False)    # Converte vírgulas para pontos
    #.astype(float)                         # Converte para número decimal
#)

def calcular_valores(row):
    if row['Fonte'] == 'Contas a Pagar':
        # Para Contas a Pagar, o valor do profissional será igual ao valor total, e valor da clínica será zero
        valor_profissional = row['Valor Total']
        valor_clinica = 0
    elif row['Fonte'] == 'Profissionais Não EME':
        # Para Profissionais Não EME, aplicar a divisão proporcional
        valor_total = row['Valor Total']
        valor_profissional = round(valor_total * 72.5 / 100, 2)  # 72,5%
        valor_clinica = round(valor_total * 27.5 / 100, 2)       # 27,5%
        # = row['Valor Total']
        #valor_clinica = 0  # A clínica não retém nenhum valor nesse caso
    else:
        # Para outros casos, aplicar o rateio e dedução de taxas
        valor_profissional = round(row['Valor Subtraído'] * rateio_dict.get(row['Nome do Profissional'], 0), 2)
        valor_clinica = round(row['Valor Subtraído'] - valor_profissional, 2)
    return pd.Series([valor_profissional, valor_clinica])



print("\nLog Final - Valores de 'Valor Total' após as transformações:")
print(filtered_data[filtered_data['Fonte'] == 'Profissionais Não EME']['Valor Total'].head(10))


# Aplicar o cálculo no DataFrame
filtered_data[['Valor Profissional', 'Valor Clínica']] = filtered_data.apply(calcular_valores, axis=1)

# Formatar valores para duas casas decimais
filtered_data[['Valor Total', 'Valor Subtraído', 'Taxa Aplicada', 'Valor Profissional', 'Valor Clínica']] = \
    filtered_data[['Valor Total', 'Valor Subtraído', 'Taxa Aplicada', 'Valor Profissional', 'Valor Clínica']].round(2)

# Reordenar as colunas no relatório final
columns_order = [
    'Data de Atendimento', 'Nome do Procedimento', 'Convênio', 'Nome do Paciente',
    'Nome do Profissional', 'Modo de Pagamento', 'Fonte', 'Valor Total',
    'Taxa Aplicada', 'Valor Subtraído', 'Valor Profissional', 'Valor Clínica'
]
filtered_data = filtered_data[columns_order]

# Criar uma tabela sumarizada com o nome do profissional e soma do valor profissional
summary_table = filtered_data.groupby('Nome do Profissional')['Valor Profissional'].sum().reset_index()
summary_table = summary_table.sort_values(by='Valor Profissional', ascending=False)

# Filtrar dados da fonte "Profissionais Não EME"
prof_nao_eme_log = filtered_data[filtered_data['Fonte'] == 'Profissionais Não EME']

# Exibir os dados no console para validação
print("\nLog - Dados da fonte 'Profissionais Não EME':")
print(prof_nao_eme_log[['Nome do Profissional', 'Valor Total', 'Valor Profissional', 'Valor Clínica']].head(10))


# Salvar o relatório consolidado e a tabela sumarizada
output_path = os.path.join(desktop_path, 'Relatorio_Consolidado_Com_Regras.xlsx')
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    filtered_data.to_excel(writer, sheet_name='Detalhado', index=False,float_format='%.2f')
    summary_table.to_excel(writer, sheet_name='Sumarizado', index=False,float_format='%.2f')

print("Relatório consolidado com regras de negócio salvo em:", output_path)
