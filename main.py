import pandas as pd
import os
import re
import sys
import tkinter as tk
from tkinter import messagebox

# Criar janela root oculta para mensagens
root = tk.Tk()
root.withdraw()

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

print("=" * 80)
print("INICIANDO PROCESSAMENTO DOS RELATÓRIOS")
print("=" * 80)


# Função para mostrar erro com mensagem
def mostrar_erro(titulo, mensagem):
    messagebox.showerror(titulo, mensagem)
    sys.exit(1)


# Função para verificar existência de arquivos
def verificar_arquivos():
    print("\n[1/10] Verificando existência dos arquivos necessários...")
    arquivos = {
        "Relatório de Comissão": comissao_path,
        "Relatório de Extrato": extrato_path,
        "Profissionais Não EME": repasses_path,
        "Amplimed": amplimed_path,
        "Contas a Pagar": contas_pagar_path,
        "Funcionários": funcionarios_path,
        "Regras de Negócio": regranegocio_path
    }

    arquivos_faltantes = []
    for nome, caminho in arquivos.items():
        if os.path.exists(caminho):
            print(f"   ✓ {nome}: Encontrado")
        else:
            print(f"   ✗ {nome}: NÃO ENCONTRADO")
            arquivos_faltantes.append(nome)

    if arquivos_faltantes:
        mensagem = "Os seguintes arquivos não foram encontrados:\n\n"
        for arquivo in arquivos_faltantes:
            mensagem += f"• {arquivo}\n"
        mensagem += "\nPor favor, verifique se todos os arquivos estão na pasta correta."
        mostrar_erro("Arquivos Não Encontrados", mensagem)
    print("   ✓ Todos os arquivos encontrados!")


verificar_arquivos()


# Função para logar contagem de profissionais
def log_professional_counts(df, profissional_column, source_name):
    if profissional_column in df.columns:
        counts = df[profissional_column].value_counts(dropna=True)
        print(f"\n   Contagem de registros por profissional em {source_name}:")
        for profissional, count in counts.items():
            print(f"      • {profissional}: {count} registros")
    else:
        print(f"\n   ⚠️  AVISO: Coluna '{profissional_column}' não encontrada em {source_name}")


# Adicionar logs detalhados para valores
def log_values(title, df, columns):
    print(f"\n   {title}:")
    print(df[columns].head(5).to_string(index=False))


# Carregar lista de funcionários
print("\n[2/10] Carregando lista de funcionários...")
try:
    df_funcionarios = pd.read_excel(funcionarios_path, skiprows=1)
    funcionarios_list = [nome.strip().upper() for nome in df_funcionarios['Nome'].dropna().unique()]
    print(f"   ✓ Total de funcionários carregados: {len(funcionarios_list)}")
except Exception as e:
    mostrar_erro("Erro ao Carregar Funcionários", f"Não foi possível carregar a lista de funcionários:\n\n{str(e)}")

# Processar o relatório de comissão
print("\n[3/10] Processando relatório de comissão...")
try:
    df_comissao = pd.read_excel(comissao_path, skiprows=1)
    print(f"   ✓ Arquivo carregado: {len(df_comissao)} linhas")

    df_comissao.loc[:, 'Código'] = df_comissao['Código'].str.strip()

    profissional = None
    df_comissao.loc[:, 'Profissional'] = None

    for index in range(len(df_comissao) - 1, -1, -1):
        if isinstance(df_comissao.at[index, 'Código'], str) and not df_comissao.at[index, 'Código'].isdigit():
            profissional = df_comissao.at[index, 'Código']
        elif profissional:
            df_comissao.loc[index, 'Profissional'] = profissional

    df_comissao = df_comissao[df_comissao['Código'].apply(lambda x: str(x).isdigit())]
    print(f"   ✓ Registros válidos após filtro: {len(df_comissao)}")
    log_professional_counts(df_comissao, 'Profissional', 'relatorio_comissao_vendedores')
except Exception as e:
    mostrar_erro("Erro ao Processar Comissão", f"Não foi possível processar o relatório de comissão:\n\n{str(e)}")

# Carregar extrato e fazer merge
print("\n[4/10] Mesclando dados com extrato...")
try:
    df_extrato = pd.read_excel(extrato_path, skiprows=1)
    print(f"   ✓ Extrato carregado: {len(df_extrato)} linhas")

    if 'Descrição' in df_extrato.columns and 'Forma de pagamento' in df_extrato.columns:
        df_completo = pd.merge(df_comissao, df_extrato[['Descrição', 'Forma de pagamento']], on='Descrição', how='left')
        df_completo.to_excel(comissao_modificado_path, index=False)
        print(f"   ✓ Arquivo modificado salvo com {len(df_completo)} registros")
    else:
        print("   ⚠️  AVISO: Colunas 'Descrição' ou 'Forma de pagamento' não encontradas no extrato")
except Exception as e:
    mostrar_erro("Erro ao Processar Extrato", f"Não foi possível processar o extrato:\n\n{str(e)}")


# Função para extrair colunas por índices com validação
def extract_columns_by_index(df, indices, nome_arquivo):
    max_index = df.shape[1] - 1
    valid_indices = [idx for idx in indices if idx <= max_index]

    if len(valid_indices) != len(indices):
        indices_invalidos = [idx for idx in indices if idx > max_index]
        print(f"   ⚠️  AVISO ({nome_arquivo}): Índices fora dos limites: {indices_invalidos}")
        print(f"      Total de colunas disponíveis: {max_index + 1}")

    if not valid_indices:
        print(f"   ❌ ERRO ({nome_arquivo}): Nenhum índice válido!")
        return pd.DataFrame()

    try:
        return df.iloc[:, valid_indices]
    except Exception as e:
        print(f"   ❌ ERRO ao extrair colunas de {nome_arquivo}: {e}")
        return pd.DataFrame()


# Ler Amplimed com tratamento de coluna vazia inicial
print("\n[5/10] Processando arquivo Amplimed...")
try:
    df_amplimed_raw = pd.read_csv(amplimed_path, sep=';', encoding='utf-8')
    print(f"   ✓ Arquivo carregado: {df_amplimed_raw.shape[0]} linhas, {df_amplimed_raw.shape[1]} colunas")

    # Verificar e remover coluna vazia inicial
    if df_amplimed_raw.columns[0] == '' or pd.isna(df_amplimed_raw.columns[0]) or df_amplimed_raw.columns[
        0].strip() == '':
        print("   ⚠️  Detectada coluna vazia inicial - removendo automaticamente...")
        df_amplimed = df_amplimed_raw.iloc[:, 1:]  # Remove primeira coluna
        print(f"   ✓ Nova estrutura: {df_amplimed.shape[1]} colunas")
    else:
        df_amplimed = df_amplimed_raw
        print("   ✓ Nenhuma coluna vazia detectada no início")

    # Ajustar índices considerando a remoção da coluna vazia
    amplimed_data = extract_columns_by_index(df_amplimed, [3, 10, 12, 16, 18, 25, 26], "Amplimed")

    if not amplimed_data.empty:
        amplimed_data.columns = ['Data de Atendimento', 'Nome do Procedimento', 'Convênio',
                                 'Nome do Paciente', 'Nome do Profissional', 'Modo de Pagamento', 'Valor Total']
        amplimed_data = amplimed_data.copy()
        amplimed_data.loc[:, 'Fonte'] = 'Amplimed'
        print(f"   ✓ Dados extraídos: {len(amplimed_data)} registros")
    else:
        print("   ❌ ERRO: Falha ao extrair dados do Amplimed")
except Exception as e:
    mostrar_erro("Erro ao Processar Amplimed", f"Não foi possível processar o arquivo Amplimed:\n\n{str(e)}")

# Ler Profissionais Não EME
print("\n[6/10] Processando Profissionais Não EME...")
try:
    df_prof_nao_eme = pd.read_excel(repasses_path)
    print(f"   ✓ Arquivo carregado: {df_prof_nao_eme.shape[0]} linhas, {df_prof_nao_eme.shape[1]} colunas")

    prof_nao_eme_data = extract_columns_by_index(df_prof_nao_eme, [1, 4, 5, 3, 11, 9], "Profissionais Não EME")

    if not prof_nao_eme_data.empty:
        prof_nao_eme_data.columns = ['Data de Atendimento', 'Nome do Procedimento', 'Convênio',
                                     'Nome do Paciente', 'Nome do Profissional', 'Valor Total']
        prof_nao_eme_data = prof_nao_eme_data.copy()
        prof_nao_eme_data.loc[:, 'Fonte'] = 'Profissionais Não EME'
        print(f"   ✓ Dados extraídos: {len(prof_nao_eme_data)} registros")
except Exception as e:
    mostrar_erro("Erro ao Processar Profissionais Não EME", f"Não foi possível processar os profissionais não EME:\n\n{str(e)}")

# Carregar Contas a Pagar com validação de colunas
print("\n[7/10] Processando Contas a Pagar...")
try:
    df_contas_pagar_raw = pd.read_excel(contas_pagar_path, skiprows=1)
    print(f"   ✓ Arquivo carregado: {df_contas_pagar_raw.shape[0]} linhas, {df_contas_pagar_raw.shape[1]} colunas")

    # Verificar se as colunas Q e R existem (índices 16 e 17)
    colunas_necessarias = ['Taxa do banco', 'Taxa da operadora']
    colunas_faltantes = []

    # Verificar pela posição (colunas Q=16 e R=17)
    if df_contas_pagar_raw.shape[1] < 18:
        print(f"   ⚠️  AVISO: Arquivo tem apenas {df_contas_pagar_raw.shape[1]} colunas")
        print("      Adicionando colunas faltantes: 'Taxa do banco' e 'Taxa da operadora'")

        # Adicionar colunas faltantes com valor 0
        while df_contas_pagar_raw.shape[1] < 17:
            df_contas_pagar_raw[f'Coluna_Extra_{df_contas_pagar_raw.shape[1]}'] = 0

        df_contas_pagar_raw['Taxa do banco'] = 0
        df_contas_pagar_raw['Taxa da operadora'] = 0
        print(f"   ✓ Colunas adicionadas automaticamente (valores = 0)")
    else:
        print("   ✓ Todas as colunas necessárias presentes")

    df_contas_pagar = df_contas_pagar_raw

    contas_pagar_data = extract_columns_by_index(df_contas_pagar, [10, 4, 5, 3, 1, 18], "Contas a Pagar")

    if not contas_pagar_data.empty:
        contas_pagar_data.columns = ['Data de Atendimento', 'Nome do Procedimento', 'Modo de Pagamento',
                                     'Nome do Paciente', 'Nome do Profissional', 'Valor Total']
        contas_pagar_data = contas_pagar_data.copy()
        contas_pagar_data.loc[:, 'Fonte'] = 'Contas a Pagar'
        print(f"   ✓ Dados extraídos: {len(contas_pagar_data)} registros")
except Exception as e:
    mostrar_erro("Erro ao Processar Contas a Pagar", f"Não foi possível processar as contas a pagar:\n\n{str(e)}")

# Carregar comissão modificado
print("\n[8/10] Processando Comissão Modificada...")
try:
    df_comissao_modificado = pd.read_excel(comissao_modificado_path)
    print(
        f"   ✓ Arquivo carregado: {df_comissao_modificado.shape[0]} linhas, {df_comissao_modificado.shape[1]} colunas")

    comissao_modificado_data = extract_columns_by_index(df_comissao_modificado, [4, 3, 2, 1, 9, 10, 7],
                                                        "Comissão Modificada")

    if not comissao_modificado_data.empty:
        comissao_modificado_data.columns = ['Data de Atendimento', 'Nome do Procedimento', 'Convênio',
                                            'Nome do Paciente', 'Nome do Profissional', 'Modo de Pagamento',
                                            'Valor Total']
        comissao_modificado_data = comissao_modificado_data.copy()
        comissao_modificado_data.loc[:, 'Fonte'] = 'Comissão Modificada'
        print(f"   ✓ Dados extraídos: {len(comissao_modificado_data)} registros")
except Exception as e:
    mostrar_erro("Erro ao Processar Comissão Modificada", f"Não foi possível processar a comissão modificada:\n\n{str(e)}")

# Consolidar dados
print("\n[9/10] Consolidando todos os dados...")
dataframes_validos = [df for df in [amplimed_data, prof_nao_eme_data, contas_pagar_data, comissao_modificado_data] if
                      not df.empty]

if not dataframes_validos:
    print("   ❌ ERRO: Nenhum dado válido para consolidar!")
    sys.exit(1)

consolidated_data = pd.concat(dataframes_validos, ignore_index=True)
print(f"   ✓ Total de registros consolidados: {len(consolidated_data)}")

# Exibir contagem por fonte
print("\n   Registros por fonte:")
for fonte, count in consolidated_data['Fonte'].value_counts().items():
    print(f"      • {fonte}: {count} registros")

# Filtrar funcionários
consolidated_data.loc[:, 'Nome do Profissional'] = consolidated_data['Nome do Profissional'].str.strip().str.upper()
filtered_data = consolidated_data[consolidated_data['Nome do Profissional'].isin(funcionarios_list)]
print(f"\n   ✓ Registros após filtro de funcionários: {len(filtered_data)}")

# Limpeza de dados
filtered_data.loc[:, 'Convênio'] = filtered_data['Convênio'].fillna('Particular')
filtered_data = filtered_data.dropna(subset=['Data de Atendimento'])
filtered_data.loc[:, 'Data de Atendimento'] = pd.to_datetime(filtered_data['Data de Atendimento'], dayfirst=True,
                                                             errors='coerce').dt.strftime('%d/%m/%Y')


# Função de limpeza de moeda
def clean_currency(value, fonte):
    try:
        value = str(value).strip()

        if fonte == 'Amplimed':
            return float(value.replace('R$', '').replace('.', '').replace(',', '.'))
        elif fonte == 'Profissionais Não EME':
            if ',' in value and '.' in value:
                value = value.replace('.', '').replace(',', '.')
            elif ',' in value:
                value = value.replace(',', '.')
            return float(value)
        else:
            if ',' in value and '.' in value:
                value = value.replace('.', '').replace(',', '.')
            elif ',' in value:
                value = value.replace(',', '.')
            return float(value)
    except ValueError:
        return 0


filtered_data['Valor Total'] = filtered_data.apply(
    lambda row: clean_currency(row['Valor Total'], row['Fonte']), axis=1
)
filtered_data['Valor Total'] = filtered_data['Valor Total'].fillna(0)

print("\n   ✓ Valores monetários convertidos com sucesso")

# Carregar regras de negócio
print("\n[10/10] Aplicando regras de negócio...")
try:
    taxas_data = pd.read_excel(regranegocio_path, sheet_name='Taxas')
    rateio_data = pd.read_excel(regranegocio_path, sheet_name='Rateio por profissional')
    print(f"   ✓ Regras carregadas: {len(taxas_data)} taxas, {len(rateio_data)} rateios")

    # Converter taxa
    taxas_data['Taxa'] = (
        taxas_data['Taxa']
        .astype(str)
        .str.replace('R$', '', regex=False)
        .str.replace('.', '', regex=False)
        .str.replace(',', '.', regex=False)
        .astype(float)
        .fillna(0)
    )

    # Criar coluna valor subtraído
    filtered_data['Valor Subtraído'] = filtered_data['Valor Total']

    # Aplicar taxas
    for _, row in taxas_data.iterrows():
        procedimento = row['Procedimento']
        taxa = row['Taxa']
        mask = (filtered_data['Nome do Procedimento'] == procedimento) & \
               (filtered_data['Fonte'] != 'Profissionais Não EME') & \
               (filtered_data['Fonte'] != 'Contas a Pagar')

        for idx in filtered_data[mask].index:
            convenio = filtered_data.at[idx, 'Convênio']
            parcel_match = re.search(r'\((\d+)/(\d+)\)', convenio)

            if parcel_match:
                parcela_atual = int(parcel_match.group(1))
                total_parcelas = int(parcel_match.group(2))
                taxa_parcela = round(taxa / total_parcelas, 2)
                filtered_data.at[idx, 'Valor Subtraído'] -= taxa_parcela
            else:
                filtered_data.at[idx, 'Valor Subtraído'] -= round(taxa, 2)

        filtered_data.loc[mask, 'Taxa Aplicada'] = round(taxa, 2)

    print("   ✓ Taxas aplicadas com sucesso")

    # Aplicar rateio
    rateio_dict = dict(zip(rateio_data['Profissionais'].str.strip().str.upper(), rateio_data['%Profissionais']))


    def calcular_valores(row):
        if row['Fonte'] == 'Contas a Pagar':
            valor_profissional = row['Valor Total']
            valor_clinica = 0
        elif row['Fonte'] == 'Profissionais Não EME':
            valor_total = row['Valor Total']
            valor_profissional = round(valor_total * 72.5 / 100, 2)
            valor_clinica = round(valor_total * 27.5 / 100, 2)
        else:
            valor_profissional = round(row['Valor Subtraído'] * rateio_dict.get(row['Nome do Profissional'], 0), 2)
            valor_clinica = round(row['Valor Subtraído'] - valor_profissional, 2)
        return pd.Series([valor_profissional, valor_clinica])


    filtered_data[['Valor Profissional', 'Valor Clínica']] = filtered_data.apply(calcular_valores, axis=1)
    print("   ✓ Rateios calculados com sucesso")

except Exception as e:
    mostrar_erro("Erro ao Aplicar Regras de Negócio", f"Não foi possível aplicar as regras de negócio:\n\n{str(e)}")

# Formatar valores
filtered_data[['Valor Total', 'Valor Subtraído', 'Taxa Aplicada', 'Valor Profissional', 'Valor Clínica']] = \
    filtered_data[['Valor Total', 'Valor Subtraído', 'Taxa Aplicada', 'Valor Profissional', 'Valor Clínica']].round(2)

# Reordenar colunas
columns_order = [
    'Data de Atendimento', 'Nome do Procedimento', 'Convênio', 'Nome do Paciente',
    'Nome do Profissional', 'Modo de Pagamento', 'Fonte', 'Valor Total',
    'Taxa Aplicada', 'Valor Subtraído', 'Valor Profissional', 'Valor Clínica'
]
filtered_data = filtered_data[columns_order]

# Criar tabela sumarizada
summary_table = filtered_data.groupby('Nome do Profissional')['Valor Profissional'].sum().reset_index()
summary_table = summary_table.sort_values(by='Valor Profissional', ascending=False)

# Salvar relatório
output_path = os.path.join(desktop_path, 'Relatorio_Consolidado_Com_Regras.xlsx')
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    filtered_data.to_excel(writer, sheet_name='Detalhado', index=False, float_format='%.2f')
    summary_table.to_excel(writer, sheet_name='Sumarizado', index=False, float_format='%.2f')

print("\n" + "=" * 80)
print("✓ PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
print("=" * 80)
print(f"\n📄 Relatório salvo em: {output_path}")
print(f"📊 Total de registros processados: {len(filtered_data)}")
print(f"👥 Total de profissionais: {len(summary_table)}")
print("\n" + "=" * 80)