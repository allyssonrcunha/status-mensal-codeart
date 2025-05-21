import dash
from dash import dcc, html, dash_table, Input, Output, State, callback
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from dash.dash_table.Format import Format, Scheme, Group
import dash_bootstrap_components as dbc
import time
from datetime import datetime
import base64
import os
import re
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO
import random

load_dotenv()

# Variáveis de cache global
CACHE_PROJETOS = None
CACHE_CODENAUTAS = None
CACHE_ACOES = None
LAST_CACHE_UPDATE = None
CACHE_DURATION = 60 * 10  # 10 minutos em segundos

EXCEL_FILE_PATH = 'Revisão Projetos - Geral.xlsx'  # Nome do arquivo que deve estar na mesma pasta do script

# Funções para salvar e carregar dados localmente como fallback
def save_data_to_local(df, name):
    """Salva um DataFrame como arquivo CSV local"""
    try:
        filename = f"{name}_backup.csv"
        df.to_csv(filename, index=False)
        print(f"Dados de {name} salvos localmente em {filename}")
        return True
    except Exception as e:
        print(f"Erro ao salvar dados localmente: {e}")
        return False

def load_data_from_local(name):
    """Carrega um DataFrame de um arquivo CSV local"""
    try:
        filename = f"{name}_backup.csv"
        if os.path.exists(filename):
            df = pd.read_csv(filename)
            print(f"Dados de {name} carregados localmente de {filename}")
            
            # Converter colunas de data se for o arquivo de ações
            if name == "acoes":
                date_cols = ['Data de Cadastro', 'Data Limite', 'Data de Conclusão']
                for col in date_cols:
                    if col in df.columns:
                        df[col] = pd.to_datetime(df[col], errors='coerce')
            
            return df
        else:
            print(f"Arquivo local {filename} não encontrado")
            return None
    except Exception as e:
        print(f"Erro ao carregar dados localmente: {e}")
        return None

# Função auxiliar para retry com backoff exponencial
def retry_with_backoff(func, max_retries=3, initial_delay=1):
    retries = 0
    while retries < max_retries:
        try:
            return func()
        except Exception as e:
            if "429" in str(e):  # Erro de quota excedida
                wait_time = initial_delay * (2 ** retries) + random.uniform(0, 1)
                print(f"Quota excedida. Aguardando {wait_time:.2f} segundos antes de tentar novamente...")
                time.sleep(wait_time)
                retries += 1
            else:
                # Para outros erros, apenas repassar a exceção
                raise e
    
    # Se chegou aqui, todas as tentativas falharam
    raise Exception(f"Falha após {max_retries} tentativas")

# Configuração do Google Sheets
def connect_google_sheets():
    try:
        # Definir escopo de acesso ao Google Drive e Sheets
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        
        # Caminho para o arquivo de credenciais
        credentials_path = 'credentials/google-credentials.json'
        
        # Verificar se o arquivo de credenciais existe
        if not os.path.exists(credentials_path):
            print(f"ERRO: Arquivo de credenciais não encontrado em {credentials_path}")
            print("Por favor, verifique se você colocou o arquivo google-credentials.json no diretório 'credentials'")
            return None
        
        # Carregar credenciais do arquivo JSON
        credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
        
        # Autorizar o cliente gspread com as credenciais
        client = gspread.authorize(credentials)
        
        # Abrir a planilha pelo nome
        try:
            spreadsheet = client.open('Revisão Projetos - Geral')
            return spreadsheet
        except gspread.exceptions.SpreadsheetNotFound:
            print("ERRO: Planilha 'Revisão Projetos - Geral' não encontrada no Google Drive")
            print("Verifique se o nome da planilha está correto e se a conta de serviço tem acesso a ela")
            return None
        
    except Exception as e:
        print(f"ERRO ao conectar com Google Sheets: {e}")
        if 'invalid_grant' in str(e).lower():
            print("Possível problema com as credenciais ou token expirado.")
        return None

# Carregar dados da planilha do Google Sheets
def load_data_from_sheets():
    global CACHE_PROJETOS, LAST_CACHE_UPDATE
    
    # Verificar se existe cache válido
    if CACHE_PROJETOS is not None and LAST_CACHE_UPDATE is not None:
        elapsed_time = time.time() - LAST_CACHE_UPDATE
        if elapsed_time < CACHE_DURATION:
            print(f"Usando dados em cache de Projetos (cache de {elapsed_time:.1f} segundos)")
            return CACHE_PROJETOS
    
    try:
        print("Carregando dados da aba Projetos...")
        
        def fetch_data():
            spreadsheet = connect_google_sheets()
            if not spreadsheet:
                print("Aviso: Usando DataFrame vazio para Projetos devido a falha na conexão com Google Sheets.")
                return pd.DataFrame()
                
            # Carregar aba Projetos
            sheet = spreadsheet.worksheet('Projetos')
            data = sheet.get_all_records()
            
            if not data:
                print("Aviso: Planilha Projetos está vazia.")
                return pd.DataFrame()
                
            df_projetos = pd.DataFrame(data)
            print(f"✅ Dados carregados com sucesso: {len(df_projetos)} projetos encontrados.")
            print(f"Colunas originais na planilha: {df_projetos.columns.tolist()}")
            return df_projetos
        
        # Usar retry com backoff exponencial
        df_projetos = retry_with_backoff(fetch_data)
        
        # Atualizar cache
        CACHE_PROJETOS = df_projetos
        LAST_CACHE_UPDATE = time.time()
        
        return df_projetos
            
    except Exception as e:
        print(f"Erro ao carregar dados de Projetos: {e}")
        # Se houver erro mas existir cache, usar dados do cache mesmo se expirado
        if CACHE_PROJETOS is not None:
            print("Usando dados em cache de Projetos devido a erro na atualização")
            return CACHE_PROJETOS
        return pd.DataFrame()

# Carregar dados dos Codenautas
def load_codenautas_from_sheets():
    global CACHE_CODENAUTAS, LAST_CACHE_UPDATE
    
    # Verificar se existe cache válido
    if CACHE_CODENAUTAS is not None and LAST_CACHE_UPDATE is not None:
        elapsed_time = time.time() - LAST_CACHE_UPDATE
        if elapsed_time < CACHE_DURATION:
            print(f"Usando dados em cache de Codenautas (cache de {elapsed_time:.1f} segundos)")
            return CACHE_CODENAUTAS
    
    try:
        print("Carregando dados da aba Codenautas...")
        
        def fetch_data():
            spreadsheet = connect_google_sheets()
            if not spreadsheet:
                print("Aviso: Usando DataFrame vazio para Codenautas devido a falha na conexão com Google Sheets.")
                return pd.DataFrame()
                
            # Carregar aba Codenautas
            sheet = spreadsheet.worksheet('Codenautas')
            data = sheet.get_all_records()
            
            if not data:
                print("Aviso: Planilha Codenautas está vazia.")
                return pd.DataFrame()
                
            df_codenautas = pd.DataFrame(data)
            print(f"✅ Dados carregados com sucesso: {len(df_codenautas)} codenautas encontrados.")
            return df_codenautas
        
        # Usar retry com backoff exponencial
        df_codenautas = retry_with_backoff(fetch_data)
        
        # Atualizar cache
        CACHE_CODENAUTAS = df_codenautas
        LAST_CACHE_UPDATE = time.time()
        
        return df_codenautas
            
    except Exception as e:
        print(f"Erro ao carregar dados de Codenautas: {e}")
        if CACHE_CODENAUTAS is not None:
            print("Usando dados em cache de Codenautas devido a erro na atualização")
            return CACHE_CODENAUTAS
        return pd.DataFrame()

# Carregar dados das Ações
def load_acoes_from_sheets():
    global CACHE_ACOES, LAST_CACHE_UPDATE
    
    # Verificar se existe cache válido
    if CACHE_ACOES is not None and LAST_CACHE_UPDATE is not None:
        elapsed_time = time.time() - LAST_CACHE_UPDATE
        if elapsed_time < CACHE_DURATION:
            print(f"Usando dados em cache de Ações (cache de {elapsed_time:.1f} segundos)")
            return CACHE_ACOES
    
    try:
        print("Carregando dados da aba Ações...")
        
        def fetch_data():
            spreadsheet = connect_google_sheets()
            if not spreadsheet:
                print("Aviso: Usando DataFrame vazio para Ações devido a falha na conexão com Google Sheets.")
                return pd.DataFrame()
                
            # Carregar aba Ações
            sheet = spreadsheet.worksheet('Ações')
            data = sheet.get_all_records()
            
            if not data:
                print("Aviso: Planilha Ações está vazia.")
                return pd.DataFrame()
                
            df_acoes = pd.DataFrame(data)
            
            # Converter colunas de data para datetime
            date_cols = ['Data de Cadastro', 'Data Limite', 'Data de Conclusão']
            for col in date_cols:
                if col in df_acoes.columns:
                    df_acoes[col] = pd.to_datetime(df_acoes[col], errors='coerce')
            
            print(f"✅ Dados carregados com sucesso: {len(df_acoes)} ações encontradas.")
            return df_acoes
        
        # Usar retry com backoff exponencial
        df_acoes = retry_with_backoff(fetch_data)
        
        # Se o dataframe estiver vazio, tentar carregar do backup local
        if df_acoes.empty:
            print("Tentando carregar ações do backup local...")
            df_local = load_data_from_local("acoes")
            if df_local is not None and not df_local.empty:
                df_acoes = df_local
                print(f"Carregadas {len(df_acoes)} ações do backup local.")
        
        # Atualizar cache
        CACHE_ACOES = df_acoes
        LAST_CACHE_UPDATE = time.time()
        
        # Salvar cópia local para backup
        if not df_acoes.empty:
            save_data_to_local(df_acoes, "acoes")
        
        return df_acoes
        
    except Exception as e:
        print(f"Erro ao carregar dados de Ações: {e}")
        if CACHE_ACOES is not None:
            print("Usando dados em cache de Ações devido a erro na atualização")
            return CACHE_ACOES
        
        # Tentar carregar do backup local
        print("Tentando carregar ações do backup local após erro...")
        df_local = load_data_from_local("acoes")
        if df_local is not None:
            return df_local
            
        return pd.DataFrame()

# Função para atualizar dados das Ações
def update_acoes_in_sheets(df_acoes):
    try:
        spreadsheet = connect_google_sheets()
        if not spreadsheet:
            return False
            
        # Preparar dados para upload
        # Converter datas para string no formato YYYY-MM-DD
        df_to_upload = df_acoes.copy()
        date_cols = ['Data de Cadastro', 'Data Limite', 'Data de Conclusão']
        for col in date_cols:
            if col in df_to_upload.columns:
                df_to_upload[col] = df_to_upload[col].apply(
                    lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) and hasattr(x, 'strftime') else ''
                )
        
        # Converter DataFrame para lista de listas
        values = [df_to_upload.columns.tolist()]  # Cabeçalho
        values.extend(df_to_upload.values.tolist())  # Dados
        
        # Atualizar planilha
        acoes_sheet = spreadsheet.worksheet('Ações')
        
        # Limpar dados existentes (exceto cabeçalho)
        existing_data = acoes_sheet.get_all_values()
        if len(existing_data) > 1:  # Se tiver mais do que só o cabeçalho
            acoes_sheet.batch_clear(["A2:Z" + str(len(existing_data))])
        
        # Inserir novos dados
        if len(values) > 1:  # Se tiver dados (além do cabeçalho)
            acoes_sheet.update('A1', values)
            print(f"Planilha Ações atualizada com {len(df_to_upload)} registros")
        else:
            # Manter apenas o cabeçalho se não houver dados
            acoes_sheet.update('A1', [values[0]])
            print("Planilha Ações atualizada apenas com o cabeçalho (sem dados)")
            
        # Atualizar cache
        global CACHE_ACOES, LAST_CACHE_UPDATE
        CACHE_ACOES = df_acoes
        LAST_CACHE_UPDATE = time.time()
        
        return True
    
    except Exception as e:
        print(f"Erro ao atualizar planilha de Ações: {e}")
        import traceback
        traceback.print_exc()
        return False

# Função para atualizar dados de Projetos
def update_projetos_in_sheets(df_projetos):
    try:
        spreadsheet = connect_google_sheets()
        if not spreadsheet:
            return False
            
        # Preparar dados para upload
        values = [df_projetos.columns.tolist()]  # Cabeçalho
        values.extend(df_projetos.values.tolist())  # Dados
        
        # Atualizar planilha
        sheet = spreadsheet.worksheet('Projetos')
        sheet.clear()  # Limpar dados existentes
        sheet.update('A1', values)  # Atualizar com novos dados
        
        print("Dados de Projetos atualizados com sucesso no Google Sheets!")
        return True
    except Exception as e:
        print(f"Erro ao atualizar dados de Projetos: {e}")
        return False

# Configuração global do tema dos gráficos Plotly
import plotly.io as pio

# Definir o tema padrão para todos os gráficos
pio.templates["codeart_theme"] = go.layout.Template(
    layout=dict(
        font=dict(
            family="Outfit, 'Segoe UI', 'Roboto', sans-serif",
            color="#303E47"
        ),
        title=dict(
            font=dict(
                family="'All Round Gothic', 'Arial Rounded MT Bold', sans-serif",
                color="#303E47"
            )
        ),
        plot_bgcolor='#ffffff',
        paper_bgcolor='#ffffff',
        colorway=['#6CC0ED', '#FED600', '#416072', '#303E47', '#28a745', '#dc3545'],
        legend=dict(
            font=dict(
                family="Outfit, 'Segoe UI', 'Roboto', sans-serif",
                color="#303E47"
            )
        ),
        xaxis=dict(
            gridcolor='#f2f2f2',
            zerolinecolor='#f2f2f2',
            title=dict(
                font=dict(
                    family="Outfit, 'Segoe UI', 'Roboto', sans-serif",
                    color="#303E47"
                )
            ),
            tickfont=dict(
                family="Outfit, 'Segoe UI', 'Roboto', sans-serif",
                color="#303E47"
            )
        ),
        yaxis=dict(
            gridcolor='#f2f2f2',
            zerolinecolor='#f2f2f2',
            title=dict(
                font=dict(
                    family="Outfit, 'Segoe UI', 'Roboto', sans-serif",
                    color="#303E47"
                )
            ),
            tickfont=dict(
                family="Outfit, 'Segoe UI', 'Roboto', sans-serif",
                color="#303E47"
            )
        )
    )
)

# Definir como template padrão
pio.templates.default = "codeart_theme"

# Cores e estilos da marca Codeart
codeart_colors = {
    'blue_sky': '#6CC0ED',  # Azul claro
    'yellow': '#FED600',    # Amarelo
    'dark_blue': '#416072', # Azul escuro
    'dark_gray': '#303E47', # Cinza escuro
    'white': '#FFFFFF',     # Branco
    'success': '#28a745',   # Verde para sucesso
    'danger': '#dc3545',    # Vermelho para alerta
    'charcoal_blue': '#172B36', # Azul muito escuro
    'cloud': '#F7F9FA',     # Cinza muito claro
    'deep_sea': '#3A84A7',  # Azul médio
    'background': '#F4F7FA', # Fundo cinza claro
    'text': '#333333',      # Texto principal
    'card_bg': '#FFFFFF'    # Fundo de cartões
}

# Paleta de cores para gráficos
codeart_chart_palette = [
    codeart_colors['blue_sky'], 
    codeart_colors['yellow'],
    codeart_colors['dark_blue'],
    codeart_colors['dark_gray'],
    codeart_colors['success'],
    codeart_colors['danger']
]

# Estilos de tipografia
font_styles = {
    'title_font': "'All Round Gothic', 'Segoe UI', 'Arial Rounded MT Bold', sans-serif",
    'body_font': "'Outfit', 'Segoe UI', 'Roboto', sans-serif"
}

# Carregar a logo da Codeart
try:
    image_filename = 'logo-codeart-solutions.png'
    if os.path.exists(image_filename):
        encoded_image = base64.b64encode(open(image_filename, 'rb').read())
        logo_src = f'data:image/png;base64,{encoded_image.decode()}'
    else:
        logo_src = None
        print(f"Aviso: Logo não encontrada em {image_filename}")
except Exception as e:
    logo_src = None
    print(f"Erro ao carregar logo: {e}")

# Estilos personalizados
custom_style = {
    'body': {
        'fontFamily': "Outfit, 'Segoe UI', 'Roboto', sans-serif",
        'margin': '0',
        'padding': '0',
        'backgroundColor': '#f8f9fa'
    },
    'header': {
        'backgroundColor': '#ffffff',
        'boxShadow': '0 2px 4px rgba(0,0,0,0.1)',
        'padding': '16px 24px',
        'display': 'flex',
        'justifyContent': 'space-between',
        'alignItems': 'center',
        'marginBottom': '24px'
    },
    'logo': {
        'height': '40px',
        'marginRight': '16px'
    },
    'title': {
        'fontFamily': "'All Round Gothic', 'Arial Rounded MT Bold', sans-serif",
        'color': codeart_colors['dark_gray'],
        'margin': '0',
        'fontSize': '1.8rem'
    },
    'last_update_style': {
        'color': '#6c757d',
        'fontSize': '0.9rem',
        'marginLeft': '10px'
    },
    'metric-card': {
        'backgroundColor': '#ffffff',
        'borderRadius': '8px',
        'boxShadow': '0 2px 4px rgba(0,0,0,0.05)',
        'padding': '16px',
        'textAlign': 'center',
        'marginBottom': '16px'
    },
    'chart-container': {
        'backgroundColor': '#ffffff',
        'borderRadius': '8px',
        'boxShadow': '0 2px 4px rgba(0,0,0,0.05)',
        'padding': '16px',
        'marginBottom': '24px'
    }
}

# Função para processar dados
def process_data(df_projetos):
    if df_projetos.empty:
        print("AVISO: DataFrame vazio recebido em process_data")
        return df_projetos # Retorna DataFrame vazio se não houver dados
    
    # Imprimir informações para debug
    print(f"Processando dados da planilha: {len(df_projetos)} linhas, {len(df_projetos.columns)} colunas")
    print(f"Colunas originais: {df_projetos.columns.tolist()}")

    # Mapeamento de nomes de colunas da planilha para os nomes esperados pelo código
    column_mapping = {
        'Mês': 'Mês',
        'Projeto': 'Projeto', 
        'GP Responsável': 'GP Responsável',
        'Status': 'Status',
        'Segmento': 'Segmento',
        'Tipo': 'Tipo',
        'Coordenação': 'Coordenação',
        'Financeiro': 'Financeiro',
        'Horas Previstas (Contrato)': 'Previsão',  # Nome real da coluna
        'Real': 'Real',
        'Saldo Acumulado': 'Saldo Acumulado',
        'Atraso em dias': 'Atraso em dias ',  # Note o espaço extra no final
        'NPS ': 'NPS ',  # Note o espaço extra no final
        'Observações': 'Observacoes',  # Sem acento no código
        'Decisões': 'Decisões'
    }
    
    # Renomear colunas conforme o mapeamento
    df_renamed = df_projetos.copy()
    
    # Verificar se temos conflitos de nome (colunas que existem no dataframe e também no mapeamento)
    # Isso acontece quando tentamos mapear 'Horas Previstas (Contrato)' para 'Previsão', mas 'Previsão' já existe
    problematic_columns = []
    for original, expected in column_mapping.items():
        if original in df_renamed.columns and expected in df_renamed.columns and original != expected:
            problematic_columns.append((original, expected))
    
    # Resolver conflitos renomeando colunas originais temporariamente
    for original, expected in problematic_columns:
        print(f"Resolvendo conflito: '{original}' e '{expected}' existem simultaneamente.")
        df_renamed = df_renamed.rename(columns={expected: f"{expected}_temp"})
    
    # Agora podemos aplicar o mapeamento com segurança
    for original, expected in column_mapping.items():
        if original in df_renamed.columns and original != expected:
            df_renamed = df_renamed.rename(columns={original: expected})
    
    # Verificar colunas obrigatórias e criar se não existirem
    required_columns = ['Mês', 'Projeto', 'GP Responsável', 'Status', 'Segmento', 'Tipo', 
                        'Coordenação', 'Financeiro', 'Previsão', 'Real', 'Saldo Acumulado', 
                        'Atraso em dias ', 'NPS ', 'Observacoes', 'Decisões']
    
    for col in required_columns:
        if col not in df_renamed.columns:
            if col == 'Observacoes':
                if 'Observações' in df_renamed.columns:
                    df_renamed = df_renamed.rename(columns={'Observações': 'Observacoes'})
                else:
                    df_renamed[col] = ""
            elif col == 'NPS ':
                if 'NPS' in df_renamed.columns:
                    df_renamed = df_renamed.rename(columns={'NPS': 'NPS '})
                else:
                    df_renamed[col] = ""
            elif col == 'Atraso em dias ':
                if 'Atraso em dias' in df_renamed.columns:
                    df_renamed = df_renamed.rename(columns={'Atraso em dias': 'Atraso em dias '})
                else:
                    df_renamed[col] = 0
            elif col == 'Previsão':
                # Verificar diversas opções possíveis para a coluna de horas previstas
                for possible_col in ['Horas Previstas (Contrato)', 'Previsão', 'Previsto']:
                    if possible_col in df_renamed.columns:
                        df_renamed = df_renamed.rename(columns={possible_col: 'Previsão'})
                        break
                else:
                    df_renamed[col] = 0
            else:
                df_renamed[col] = "Não Informado" if col in ['GP Responsável', 'Status', 'Segmento', 'Tipo', 'Coordenação', 'Financeiro'] else 0
    
    # 1. Formatar coluna 'Mês' para exibição e filtro (ex: Abr/2025)
    try:
        df_renamed['Mês_datetime'] = pd.to_datetime(df_renamed['Mês'], errors='coerce')
        # CORREÇÃO: O ponto (.) antes de 'Out' foi substituído por dois pontos (:)
        month_map_pt = {1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun',
                        7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'}
        df_renamed['MesAnoFormatado'] = df_renamed['Mês_datetime'].apply(
            lambda x: f"{month_map_pt[x.month]}/{x.year}" if pd.notnull(x) and hasattr(x, 'month') and hasattr(x, 'year') else str(df_renamed.loc[x.name if hasattr(x, 'name') else -1, 'Mês'] if hasattr(x, 'name') else 'Data Inválida')
        )
        # Fallback for any original strings that couldn't be parsed but should be kept
        df_renamed['MesAnoFormatado'] = np.where(df_renamed['Mês_datetime'].isna(), df_renamed['Mês'].astype(str), df_renamed['MesAnoFormatado'])
        
        # Adicionar coluna de ano-mês para agrupamento nos gráficos de evolução
        df_renamed['Ano_Mes'] = df_renamed['Mês_datetime'].dt.strftime('%Y-%m')
    except Exception as e:
        print(f"Erro ao formatar coluna 'Mês': {e}. Usando como string.")
        df_renamed['MesAnoFormatado'] = df_renamed['Mês'].astype(str)
        df_renamed['Ano_Mes'] = df_renamed['Mês'].astype(str)

    # 2. Converter outras colunas de filtro para string para consistência
    for col in ['GP Responsável', 'Status', 'Segmento', 'Tipo', 'Coordenação', 'Financeiro']:
        if col in df_renamed.columns:
            df_renamed[col] = df_renamed[col].astype(str).fillna('Não Informado') # Tratar NaNs nos filtros
        else:
            df_renamed[col] = 'Não Informado' # Adicionar coluna se não existir

    # 3. Identificar projetos críticos baseado na coluna "Decisões"
    if 'Decisões' in df_renamed.columns:
        df_renamed['Prioridade'] = np.where(df_renamed['Decisões'].astype(str).str.contains('Crítico', na=False), 'Crítico', 'Normal')
    else:
        df_renamed['Prioridade'] = 'Normal' # Se a coluna não existir, nenhum é crítico

    # 4. Combinar NPS e emoji em uma única coluna
    def nps_com_emoji(nps_value):
        if pd.isna(nps_value) or str(nps_value).strip() == '' or str(nps_value).lower() == 'nan': # Tratar NaN, string vazia e 'nan'
            return ""
        elif nps_value == "Promotor":
            return "Promotor 😀"  # Emoji feliz verde
        elif nps_value == "Neutro":
            return "Neutro 😐"  # Emoji neutro amarelo
        elif nps_value == "Detrator":
            return "Detrator 😡"  # Emoji triste vermelho
        else:
            return str(nps_value) # Retorna o valor original se for diferente dos esperados

    if 'NPS ' in df_renamed.columns:
        df_renamed['NPS_Combinado'] = df_renamed['NPS '].apply(nps_com_emoji)
    else:
        df_renamed['NPS_Combinado'] = "" # Se a coluna não existir, criar vazia

    # 5. Garantir que colunas para cálculo sejam numéricas e preencher NaNs
    cols_to_convert_to_numeric = ['Atraso em dias ', 'Previsão', 'Real', 'Saldo Acumulado', 'Horas Previstas (Contrato)', 'Horas Mês']
    for col in cols_to_convert_to_numeric:
        try:
            if col in df_renamed.columns:
                # Converter para numérico, tratando erros como NaN
                df_renamed[col] = pd.to_numeric(df_renamed[col], errors='coerce')
                
                # Corrigir valores com formatação incorreta
                if col in ['Previsão', 'Real', 'Saldo Acumulado', 'Horas Previstas (Contrato)', 'Horas Mês']:
                    # Verificar se há valores extremamente altos
                    valor_medio = df_renamed[col].median()  # Usar mediana em vez de média para ser menos afetado por outliers
                    valores_suspeitosos = df_renamed[col] > valor_medio * 10  # Valores 10x acima da mediana
                    
                    # Corrigir apenas se houver poucos valores suspeitos (evitar corrigir dados válidos)
                    if valores_suspeitosos.sum() > 0 and valores_suspeitosos.sum() < len(df_renamed) * 0.2:
                        print(f"Encontrados {valores_suspeitosos.sum()} valores suspeitosamente altos na coluna '{col}'. Mediana: {valor_medio:.2f}")
                        
                        # Verificar quais projetos estão afetados
                        projetos_afetados = df_renamed.loc[valores_suspeitosos, 'Projeto'].tolist()
                        print(f"Projetos afetados: {projetos_afetados}")
                        
                        # Determinando o fator de correção com base na magnitude dos valores
                        valores_altos = df_renamed.loc[valores_suspeitosos, col]
                        
                        if valores_altos.max() > 1000000:
                            fator_correcao = 10000
                        elif valores_altos.max() > 100000:
                            fator_correcao = 1000
                        elif valores_altos.max() > 10000:
                            fator_correcao = 100
                        else:
                            fator_correcao = 10
                            
                        print(f"Aplicando fator de correção de {fator_correcao} para valores altos na coluna {col}")
                        
                        # Aplicar a correção
                        df_renamed.loc[valores_suspeitosos, col] = df_renamed.loc[valores_suspeitosos, col] / fator_correcao
                        
                        # Verificar novos valores
                        print(f"Após correção '{col}': Min={df_renamed[col].min():.2f}, Max={df_renamed[col].max():.2f}, Média={df_renamed[col].mean():.2f}")
                
                # Preencher valores nulos com zero
                df_renamed[col] = df_renamed[col].fillna(0)
            else:
                print(f"AVISO: Coluna '{col}' não encontrada, criando com zeros.")
                df_renamed[col] = 0
        except Exception as e:
            print(f"ERRO ao processar coluna '{col}': {e}")
            # Tentar recuperar a coluna em caso de erro
            df_renamed[col] = 0

    # 6. Garantir que a coluna Observacoes exista e tratar NaNs
    if 'Observacoes' not in df_renamed.columns:
        if 'Observações' in df_renamed.columns:
            print(f"INFO: Renomeando coluna 'Observações' para 'Observacoes'")
            df_renamed = df_renamed.rename(columns={'Observações': 'Observacoes'})
        else:
            print(f"AVISO: Coluna 'Observações' não encontrada, criando 'Observacoes' vazia")
            df_renamed['Observacoes'] = ""
    else:
        print(f"INFO: Coluna 'Observacoes' já existe no DataFrame")
    
    # Debug: Verificar valores da coluna Observacoes
    if 'Observacoes' in df_renamed.columns:
        df_renamed['Observacoes'] = df_renamed['Observacoes'].fillna("") # Preencher NaNs com string vazia
        non_empty = (df_renamed['Observacoes'] != '').sum()
        print(f"INFO: A coluna 'Observacoes' tem {non_empty} valores não vazios de {len(df_renamed)} registros")

    # 7. Extrair nome do cliente do nome do projeto
    if 'Projeto' in df_renamed.columns:
        # Função para extrair o nome do cliente do nome do projeto
        def extract_client_name(project_name):
            if pd.isna(project_name) or project_name == "":
                return "Não informado"
            
            # Padrão: Nome do cliente antes do primeiro "|" ou o nome completo se não houver "|"
            parts = str(project_name).split('|')
            client_name = parts[0].strip()
            return client_name
        
        df_renamed['Cliente'] = df_renamed['Projeto'].apply(extract_client_name)

    # Print para debug
    print(f"Colunas disponíveis após processamento: {df_renamed.columns.tolist()}")
    print(f"Linhas após processamento: {len(df_renamed)}")
    
    # Verificar métricas básicas
    try:
        print("\nMétricas básicas após processamento:")
        for col in ['Previsão', 'Real', 'Saldo Acumulado']:
            if col in df_renamed.columns:
                print(f"  {col}: Min={df_renamed[col].min()}, Max={df_renamed[col].max()}, Média={df_renamed[col].mean():.2f}")
    except Exception as e:
        print(f"Erro ao calcular métricas básicas: {e}")
    
    return df_renamed

# Função para processar dados das ações
def process_acoes(df_acoes):
    if df_acoes.empty:
        return df_acoes
        
    # Garantir que Status e Prioridade tenham valores padrão
    df_acoes['Status'] = df_acoes['Status'].fillna('Pendente')
    df_acoes['Prioridade'] = df_acoes['Prioridade'].fillna('Média')
    
    # Calcular dias até a data limite (para ações pendentes)
    hoje = pd.Timestamp.now().normalize()
    
    # Verificar se há valores NaT/None e tratá-los antes da operação
    df_acoes['Dias Restantes'] = pd.NA
    mask_data_limite_valida = ~df_acoes['Data Limite'].isna()
    
    # Aplicar o cálculo apenas onde há datas válidas
    if mask_data_limite_valida.any():
        # Garantir que a data limite seja um objeto datetime antes do cálculo
        # Converter qualquer string para datetime se necessário
        if df_acoes.loc[mask_data_limite_valida, 'Data Limite'].dtype == 'object':
            df_acoes.loc[mask_data_limite_valida, 'Data Limite'] = pd.to_datetime(
                df_acoes.loc[mask_data_limite_valida, 'Data Limite'], 
                errors='coerce'
            )
            # Atualizar a máscara para considerar apenas valores válidos após a conversão
            mask_data_limite_valida = ~df_acoes['Data Limite'].isna()
        
        # Calcular dias restantes para cada linha individualmente para evitar operações com arrays
        for idx in df_acoes[mask_data_limite_valida].index:
            try:
                data_limite = pd.to_datetime(df_acoes.at[idx, 'Data Limite'])
                if pd.notna(data_limite):
                    df_acoes.at[idx, 'Dias Restantes'] = (data_limite - hoje).days
            except Exception as e:
                print(f"Erro ao calcular dias restantes para índice {idx}: {e}")
                df_acoes.at[idx, 'Dias Restantes'] = pd.NA
    
    # Marcar ações atrasadas (com status pendente e data limite passada)
    df_acoes['Atrasada'] = (
        (df_acoes['Dias Restantes'].notna()) & 
        (df_acoes['Dias Restantes'] < 0) & 
        (df_acoes['Status'] != 'Concluída')
    ).astype(int)
    
    # Calcular tempo de conclusão para ações concluídas
    df_acoes['Tempo de Conclusão'] = pd.NA
    mask_concluida = (df_acoes['Status'] == 'Concluída') & ~df_acoes['Data de Conclusão'].isna() & ~df_acoes['Data de Cadastro'].isna()
    
    if mask_concluida.any():
        # Garantir que ambas as datas sejam objetos datetime
        if df_acoes.loc[mask_concluida, 'Data de Conclusão'].dtype == 'object':
            df_acoes.loc[mask_concluida, 'Data de Conclusão'] = pd.to_datetime(
                df_acoes.loc[mask_concluida, 'Data de Conclusão'], 
                errors='coerce'
            )
        
        if df_acoes.loc[mask_concluida, 'Data de Cadastro'].dtype == 'object':
            df_acoes.loc[mask_concluida, 'Data de Cadastro'] = pd.to_datetime(
                df_acoes.loc[mask_concluida, 'Data de Cadastro'], 
                errors='coerce'
            )
        
        # Atualizar máscara após conversões
        mask_concluida = (df_acoes['Status'] == 'Concluída') & ~df_acoes['Data de Conclusão'].isna() & ~df_acoes['Data de Cadastro'].isna()
        
        # Calcular tempo de conclusão linha por linha
        for idx in df_acoes[mask_concluida].index:
            try:
                data_conclusao = pd.to_datetime(df_acoes.at[idx, 'Data de Conclusão'])
                data_cadastro = pd.to_datetime(df_acoes.at[idx, 'Data de Cadastro'])
                if pd.notna(data_conclusao) and pd.notna(data_cadastro):
                    df_acoes.at[idx, 'Tempo de Conclusão'] = (data_conclusao - data_cadastro).days
            except Exception as e:
                print(f"Erro ao calcular tempo de conclusão para índice {idx}: {e}")
                df_acoes.at[idx, 'Tempo de Conclusão'] = pd.NA
    
    return df_acoes

# Inicializar o aplicativo Dash com tema Bootstrap e definir o título da página
app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP, 'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.1/css/all.min.css'],
    title="Status mensal | Codeart",
    suppress_callback_exceptions=True  # Adicionado para suprimir exceções de callback
)
server = app.server

# Carregar e processar dados iniciais
print("Iniciando carregamento de dados...")
df_projetos_initial = load_data_from_sheets()
if df_projetos_initial.empty:
    # Tentar carregar do backup local
    print("Tentando carregar dados de projetos do backup local...")
    df_local = load_data_from_local("projetos")
    if df_local is not None:
        df_projetos_initial = df_local

print("Processando dados...")
df_projetos_initial = process_data(df_projetos_initial)
print("Carregando dados de codenautas...")
df_codenautas_initial = load_codenautas_from_sheets()
if df_codenautas_initial.empty:
    # Tentar carregar do backup local
    print("Tentando carregar dados de codenautas do backup local...")
    df_local = load_data_from_local("codenautas")
    if df_local is not None:
        df_codenautas_initial = df_local

print("Carregando dados de ações...")
df_acoes_initial = load_acoes_from_sheets()
if df_acoes_initial.empty:
    # Tentar carregar do backup local
    print("Tentando carregar dados de ações do backup local...")
    df_local = load_data_from_local("acoes")
    if df_local is not None:
        df_acoes_initial = df_local

print("Processando dados de ações...")
df_acoes_initial = process_acoes(df_acoes_initial)
print("Dados carregados com sucesso!")

# Obter listas para os filtros iniciais
def get_filter_options(df):
    if df.empty:
         return [], [], [], [], [], [], []  # Adicionado um [] extra para o status financeiro
    meses_anos = sorted(df['MesAnoFormatado'].unique())
    gestoras = sorted(df['GP Responsável'].unique())
    status_list = sorted(df['Status'].unique())
    segmentos = sorted(df['Segmento'].unique())
    tipos = sorted(df['Tipo'].unique())
    coordenacoes = sorted(df['Coordenação'].unique())
    # Adicionar status financeiro
    financeiro_list = sorted(df['Financeiro'].astype(str).unique())

    return meses_anos, gestoras, status_list, segmentos, tipos, coordenacoes, financeiro_list

meses_anos_initial, gestoras_initial, status_list_initial, segmentos_initial, tipos_initial, coordenacoes_initial, financeiro_list_initial = get_filter_options(df_projetos_initial)

# Layout simples para teste
app.layout = html.Div(style=custom_style['body'], children=[
    # Cabeçalho com logo, título e botão de atualização
    html.Div(style=custom_style['header'], children=[
        html.Div([ # Container para logo e título
            html.Img(src=logo_src, style=custom_style['logo']) if logo_src else None,
            html.H1("Status Mensal Codeart", style=custom_style['title'])
        ], style={'display': 'flex', 'align-items': 'center'}),
        html.Div([ # Container para botão e hora
            dbc.Button(
                [html.I(className="fas fa-sync-alt me-2"), " Atualizar Dados"],
                id="refresh-data-button",
                color="primary", 
                className="me-2",
                style={'backgroundColor': codeart_colors['blue_sky'], 'borderColor': codeart_colors['blue_sky']}
            ),
            html.Span(id="last-update-time", style=custom_style['last_update_style'])
        ], style={'display': 'flex', 'align-items': 'center'})
    ]),

    # Container principal
    dbc.Container([
        # Sistema de abas
        dbc.Tabs([
            # Aba de Projetos
            dbc.Tab(label="Projetos", tab_id="tab-projetos", children=[
                # Métricas 
                dbc.Row([
                    dbc.Col([html.Div([html.Div(id="total-projetos", className="metric-value"), html.Div("Total de Projetos", className="metric-label")], style=custom_style['metric-card'])], width=3),
                    dbc.Col([html.Div([html.Div(id="total-clientes", className="metric-value"), html.Div("Total de Clientes", className="metric-label")], style=custom_style['metric-card'])], width=3),
                    dbc.Col([html.Div([html.Div(id="projetos-atrasados", className="metric-value"), html.Div("Projetos Atrasados", className="metric-label")], style=custom_style['metric-card'])], width=3),
                    dbc.Col([html.Div([html.Div(id="projetos-criticos", className="metric-value"), html.Div("Projetos Críticos", className="metric-label")], style=custom_style['metric-card'])], width=3),
                ], className="mb-4"),
                
                # Filtros
                dbc.Row([
                    dbc.Col([
                        html.H5("Filtros", className="mb-2"),
                        dbc.Row([
                            dbc.Col([
                                html.Label("Mês/Ano"),
                                dcc.Dropdown(
                                    id="mes-ano-filter",
                                    options=[{"label": mes, "value": mes} for mes in meses_anos_initial],
                                    multi=True,
                                    placeholder="Selecione o mês/ano"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Gestora"),
                                dcc.Dropdown(
                                    id="gestora-filter",
                                    options=[{"label": gp, "value": gp} for gp in gestoras_initial],
                                    multi=True,
                                    placeholder="Selecione a gestora"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Status"),
                                dcc.Dropdown(
                                    id="status-filter",
                                    options=[{"label": status, "value": status} for status in status_list_initial],
                                    multi=True,
                                    placeholder="Selecione o status"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Coordenação"),
                                dcc.Dropdown(
                                    id="coordenacao-filter",
                                    options=[{"label": coord, "value": coord} for coord in coordenacoes_initial],
                                    multi=True,
                                    placeholder="Selecione a coordenação"
                                )
                            ], width=3),
                        ]),
                        dbc.Row([
                            dbc.Col([
                                html.Label("Segmento"),
                                dcc.Dropdown(
                                    id="segmento-filter",
                                    options=[{"label": seg, "value": seg} for seg in segmentos_initial],
                                    multi=True,
                                    placeholder="Selecione o segmento"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Tipo"),
                                dcc.Dropdown(
                                    id="tipo-filter",
                                    options=[{"label": tipo, "value": tipo} for tipo in tipos_initial],
                                    multi=True,
                                    placeholder="Selecione o tipo"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Financeiro"),
                                dcc.Dropdown(
                                    id="financeiro-filter",
                                    options=[{"label": fin, "value": fin} for fin in financeiro_list_initial],
                                    multi=True,
                                    placeholder="Selecione o status financeiro"
                                )
                            ], width=3),
                            dbc.Col(width=3),
                        ]),
                        dbc.Row([
                            dbc.Col([
                                html.Div([
                                    dbc.Button("Aplicar Filtros", id="apply-project-filters", color="primary", className="me-2", style={'backgroundColor': codeart_colors['dark_blue']}),
                                    dbc.Button("Limpar Filtros", id="reset-project-filters", color="secondary")
                                ], className="mt-3 mb-4")
                            ], width=12),
                        ])
                    ])
                ]),
                
                # Gráficos
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="status-chart")], style=custom_style['chart-container'])], width=6),
                    dbc.Col([html.Div([dcc.Graph(id="financeiro-chart")], style=custom_style['chart-container'])], width=6),
                ]),
                
                # Gráfico de NPS e gráfico de Projetos por Gestora (em tela cheia)
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="nps-chart")], style=custom_style['chart-container'])], width=6),
                    dbc.Col([html.Div([dcc.Graph(id="segmento-chart")], style=custom_style['chart-container'])], width=6),
                ]),
                
                # Gráfico de Projetos por Gestora
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="projetos-gp-chart")], style=custom_style['chart-container'])], width=12),
                ]),
                
                # Gráficos adicionais
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="horas-chart")], style=custom_style['chart-container'])], width=12),
                ]),
                
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="saldo-chart")], style=custom_style['chart-container'])], width=12),
                ]),
                
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="atraso-coordenacao-chart")], style=custom_style['chart-container'])], width=12),
                ]),
                
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="evolucao-quitados-chart")], style=custom_style['chart-container'])], width=12),
                ]),
                
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="evolucao-atrasados-chart")], style=custom_style['chart-container'])], width=12),
                ]),
                
                # Tabela de dados
                dbc.Row([
                    dbc.Col([
                        html.Div([
                            html.H3("Lista de Projetos", className="mt-4 mb-3 d-inline-block"),
                            html.Div([
                                dbc.Button(
                                    [html.I(className="fas fa-file-export me-2"), " Exportar Dados"],
                                    id="export-table-button",
                                    color="primary",
                                    size="sm",
                                    className="ms-3"
                                ),
                                dcc.Download(id="download-dataframe-csv")
                            ], className="d-inline-block")
                        ])
                    ], width=12)
                ]),
                
                html.Div([dash_table.DataTable(
                    id="projetos-table",
                    columns=[
                        # Nova coluna para o ícone de ação
                        {"name": "", "id": "action_icon", "type": "text"},
                        {"name": "Projeto", "id": "Projeto"},
                        {"name": "Cliente", "id": "Cliente"},
                        {"name": "Gestora", "id": "GP Responsável"},
                        {"name": "Coordenação", "id": "Coordenação"},
                        {"name": "Segmento", "id": "Segmento"},
                        {"name": "Tipo", "id": "Tipo"},
                        {"name": "Status", "id": "Status"},
                        {"name": "Horas Previstas", "id": "Previsão", "type": "numeric", "format": Format(precision=1, scheme=Scheme.fixed)},
                        {"name": "Horas Realizadas", "id": "Real", "type": "numeric", "format": Format(precision=1, scheme=Scheme.fixed)},
                        {"name": "Saldo", "id": "Saldo Acumulado", "type": "numeric", "format": Format(precision=1, scheme=Scheme.fixed)},
                        {"name": "Atraso (dias)", "id": "Atraso em dias ", "type": "numeric", "format": Format(precision=0, scheme=Scheme.fixed)},
                        {"name": "NPS", "id": "NPS_Combinado"},
                        {"name": "Financeiro", "id": "Financeiro"},
                        {"name": "Observações", "id": "Observacoes"}
                    ],
                    style_table={'overflowX': 'auto'},
                    style_header={
                        'backgroundColor': codeart_colors['charcoal_blue'],
                        'color': codeart_colors['cloud'],
                        'fontWeight': 'bold',
                        'textAlign': 'center',
                        'fontFamily': font_styles['title_font'],
                        'border': f'1px solid {codeart_colors["deep_sea"]}'
                    },
                    style_cell={
                        'textAlign': 'left',
                        'padding': '10px',
                        'fontSize': '14px',
                        'fontFamily': font_styles['body_font'],
                        'color': codeart_colors['charcoal_blue']
                    },
                    style_data_conditional=[
                        # Estilo para linhas alternadas
                        {
                            'if': {'row_index': 'odd'},
                            'backgroundColor': codeart_colors['cloud']
                        },
                        # Estilo para projetos críticos
                        {
                            'if': {
                                'filter_query': '{Prioridade} = "Crítico"'
                            },
                            'backgroundColor': '#FFF3CD',  # Amarelo claro para destaque
                            'fontWeight': 'bold'
                        },
                        # Estilo para projetos atrasados
                        {
                            'if': {
                                'filter_query': '{Status} = "Atrasado"'
                            },
                            'color': codeart_colors['danger']
                        },
                        # Estilo para saldo negativo
                        {
                            'if': {
                                'filter_query': '{Saldo Acumulado} < 0',
                                'column_id': 'Saldo Acumulado'
                            },
                            'color': codeart_colors['danger']
                        },
                        # Estilo para saldo positivo
                        {
                            'if': {
                                'filter_query': '{Saldo Acumulado} > 0',
                                'column_id': 'Saldo Acumulado'
                            },
                            'color': codeart_colors['success']
                        },
                        # Estilo para a coluna de ícone de ação
                        {
                            'if': {'column_id': 'action_icon'},
                            'textAlign': 'center',
                            'width': '40px',
                            'cursor': 'pointer',
                            'color': codeart_colors['blue_sky'],
                            'fontWeight': 'bold',
                            'fontSize': '18px'
                        }
                    ],
                    page_size=20,  # Aumentado de 10 para 20 linhas
                    sort_action='native',
                    filter_action='native',
                    sort_mode='multi',
                    style_as_list_view=True,
                    css=[{"selector": ".dash-cell div.dash-cell-value", "rule": "display: inline; white-space: inherit; overflow: inherit; text-overflow: inherit;"}],
                    cell_selectable=True,
                    row_selectable=False,
                    selected_cells=[]
                )], style={'overflowX': 'auto'}),
            ]),
            
            # Aba de Ações
            dbc.Tab(label="Ações", tab_id="tab-acoes", children=[
                # Métricas de ações
                dbc.Row([
                    dbc.Col([html.Div([html.Div(id="total-acoes", className="metric-value"), html.Div("Total de Ações", className="metric-label")], style=custom_style['metric-card'])], width=3),
                    dbc.Col([html.Div([html.Div(id="acoes-pendentes", className="metric-value"), html.Div("Ações Pendentes", className="metric-label")], style=custom_style['metric-card'])], width=3),
                    dbc.Col([html.Div([html.Div(id="acoes-concluidas", className="metric-value"), html.Div("Ações Concluídas", className="metric-label")], style=custom_style['metric-card'])], width=3),
                    dbc.Col([html.Div([html.Div(id="acoes-atrasadas", className="metric-value"), html.Div("Ações Atrasadas", className="metric-label")], style=custom_style['metric-card'])], width=3),
                ], className="mb-4"),
                
                # Filtros para ações
                dbc.Row([
                    dbc.Col([
                        html.H5("Filtros", className="mb-2"),
                        dbc.Row([
                            dbc.Col([
                                html.Label("Mês/Ano"),
                                dcc.Dropdown(
                                    id="mes-ano-filter-acoes",
                                    options=[{"label": mes, "value": mes} for mes in meses_anos_initial],
                                    multi=True,
                                    placeholder="Selecione o mês/ano"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Responsável"),
                                dcc.Dropdown(
                                    id="responsavel-filter-acoes",
                                    options=[],  # Será preenchido pelo callback
                                    multi=True,
                                    placeholder="Selecione o responsável"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Status"),
                                dcc.Dropdown(
                                    id="status-filter-acoes",
                                    options=[
                                        {"label": "Pendente", "value": "Pendente"},
                                        {"label": "Em Andamento", "value": "Em Andamento"},
                                        {"label": "Concluída", "value": "Concluída"}
                                    ],
                                    multi=True,
                                    placeholder="Selecione o status"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Prioridade"),
                                dcc.Dropdown(
                                    id="prioridade-filter-acoes",
                                    options=[
                                        {"label": "Baixa", "value": "Baixa"},
                                        {"label": "Média", "value": "Média"},
                                        {"label": "Alta", "value": "Alta"}
                                    ],
                                    multi=True,
                                    placeholder="Selecione a prioridade"
                                )
                            ], width=3),
                        ]),
                        dbc.Row([
                            dbc.Col([
                                html.Div([
                                    dbc.Button("Aplicar Filtros", id="apply-acoes-filters", color="primary", className="me-2", style={'backgroundColor': codeart_colors['dark_blue']}),
                                    dbc.Button("Limpar Filtros", id="reset-acoes-filters", color="secondary", className="me-2"),
                                    dbc.Button([html.I(className="fas fa-plus me-2"), " Nova Ação"], id="nova-acao-btn", color="success", style={'backgroundColor': codeart_colors['success']})
                                ], className="mt-3 mb-4")
                            ], width=12)
                        ])
                    ])
                ]),
                
                # Gráficos para ações
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="status-acoes-chart")], style=custom_style['chart-container'])], width=6),
                    dbc.Col([html.Div([dcc.Graph(id="prioridade-acoes-chart")], style=custom_style['chart-container'])], width=6),
                ]),
                
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="responsaveis-acoes-chart")], style=custom_style['chart-container'])], width=6),
                    dbc.Col([html.Div([dcc.Graph(id="evolucao-acoes-chart")], style=custom_style['chart-container'])], width=6),
                ]),
                
                # Tabela de ações
                dbc.Row([
                    dbc.Col([
                        html.Div([
                            html.H3("Lista de Ações", className="mt-4 mb-3"),
                            html.Div([
                                dbc.Button(
                                    [html.I(className="fas fa-download me-2"), " Exportar"],
                                    id="export-acoes-button",
                                    color="success",
                                    className="mb-3",
                                    style={'backgroundColor': codeart_colors['success']}
                                ),
                                dash_table.DataTable(
                                    id="acoes-table",
                                    columns=[
                                        {"name": "ID", "id": "ID da Ação"},
                                        {"name": "Projeto", "id": "Projeto"},
                                        {"name": "Descrição", "id": "Descrição da Ação"},
                                        {"name": "Responsáveis", "id": "Responsáveis"},
                                        {"name": "Data Limite", "id": "Data Limite"},
                                        {"name": "Status", "id": "Status"},
                                        {"name": "Prioridade", "id": "Prioridade"}
                                    ],
                                    page_size=10,
                                    style_table={'overflowX': 'auto'},
                                    style_cell={
                                        'textAlign': 'left',
                                        'padding': '8px'
                                    },
                                    style_header={
                                        'backgroundColor': codeart_colors['dark_gray'],
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    style_data_conditional=[
                                        {
                                            'if': {'row_index': 'odd'},
                                            'backgroundColor': '#f8f9fa'
                                        }
                                    ],
                                ),
                                dcc.Download(id="download-acoes-xlsx")
                            ], style={'overflowX': 'auto'})
                        ])
                    ], width=12)
                ]),
            ]),
        ], id="tabs", active_tab="tab-projetos"),
        
        # Modais para Ações
        dbc.Modal(
            [
                dbc.ModalHeader("Cadastrar Nova Ação"),
                dbc.ModalBody([
                    dbc.Row([
                        dbc.Col([
                            html.Label("Projeto"),
                            dcc.Dropdown(id="modal-projeto", options=[])
                        ], width=6),
                        dbc.Col([
                            html.Label("Mês de Referência"),
                            dcc.Dropdown(id="modal-mes-referencia", options=[])
                        ], width=6),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Prioridade"),
                            dcc.Dropdown(
                                id="modal-prioridade",
                                options=[
                                    {"label": "Baixa", "value": "Baixa"},
                                    {"label": "Média", "value": "Média"},
                                    {"label": "Alta", "value": "Alta"}
                                ],
                                value="Média"
                            )
                        ], width=6),
                        dbc.Col([
                            html.Label("Responsáveis"),
                            dcc.Dropdown(id="modal-responsaveis", options=[], multi=True)
                        ], width=6),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Descrição da Ação"),
                            dbc.Textarea(id="modal-descricao", rows=3)
                        ], width=12),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Data Limite"),
                            dcc.DatePickerSingle(id="modal-data-limite", date=None)
                        ], width=4),
                        dbc.Col([
                            html.Label("Status"),
                            dcc.Dropdown(
                                id="modal-status",
                                options=[
                                    {"label": "Pendente", "value": "Pendente"},
                                    {"label": "Em Andamento", "value": "Em Andamento"},
                                    {"label": "Concluída", "value": "Concluída"}
                                ],
                                value="Pendente"
                            )
                        ], width=4),
                        dbc.Col([
                            html.Label("Data de Conclusão"),
                            dcc.DatePickerSingle(id="modal-data-conclusao", date=None)
                        ], width=4),
                    ]),
                    dbc.Alert("Preencha todos os campos obrigatórios", id="modal-alert-text", color="danger", is_open=False)
                ]),
                dbc.ModalFooter([
                    dbc.Button("Cancelar", id="modal-cancel", color="secondary", className="me-2"),
                    dbc.Button("Salvar", id="modal-save", color="primary", style={'backgroundColor': codeart_colors['dark_blue']})
                ]),
            ],
            id="modal-cadastro-acao",
            size="lg",
            is_open=False,
        ),
        
        dbc.Modal(
            [
                dbc.ModalHeader("Editar Ação"),
                dbc.ModalBody([
                    dbc.Row([
                        dbc.Col([
                            html.Label("ID da Ação"),
                            dbc.Input(id="modal-edit-id", readonly=True)
                        ], width=6),
                        dbc.Col([
                            html.Label("Projeto"),
                            dcc.Dropdown(id="modal-edit-projeto", options=[])
                        ], width=6),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Mês de Referência"),
                            dcc.Dropdown(id="modal-edit-mes-referencia", options=[])
                        ], width=6),
                        dbc.Col([
                            html.Label("Prioridade"),
                            dcc.Dropdown(
                                id="modal-edit-prioridade",
                                options=[
                                    {"label": "Baixa", "value": "Baixa"},
                                    {"label": "Média", "value": "Média"},
                                    {"label": "Alta", "value": "Alta"}
                                ]
                            )
                        ], width=6),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Descrição da Ação"),
                            dbc.Textarea(id="modal-edit-descricao", rows=3)
                        ], width=12),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Responsáveis"),
                            dcc.Dropdown(id="modal-edit-responsaveis", options=[], multi=True)
                        ], width=12),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Data Limite"),
                            dcc.DatePickerSingle(id="modal-edit-data-limite", date=None)
                        ], width=4),
                        dbc.Col([
                            html.Label("Status"),
                            dcc.Dropdown(
                                id="modal-edit-status",
                                options=[
                                    {"label": "Pendente", "value": "Pendente"},
                                    {"label": "Em Andamento", "value": "Em Andamento"},
                                    {"label": "Concluída", "value": "Concluída"}
                                ]
                            )
                        ], width=4),
                        dbc.Col([
                            html.Label("Data de Conclusão"),
                            dcc.DatePickerSingle(id="modal-edit-data-conclusao", date=None)
                        ], width=4),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Observações de Conclusão"),
                            dbc.Textarea(id="modal-edit-observacoes", rows=2)
                        ], width=12),
                    ]),
                    dbc.Alert("Preencha todos os campos obrigatórios", id="modal-edit-alert-text", color="danger", is_open=False)
                ]),
                dbc.ModalFooter([
                    dbc.Button("Cancelar", id="modal-edit-cancel", color="secondary", className="me-2"),
                    dbc.Button("Salvar", id="modal-edit-save", color="primary", style={'backgroundColor': codeart_colors['dark_blue']})
                ]),
            ],
            id="modal-edicao-acao",
            size="lg",
            is_open=False,
        ),
        
        # Modal para Nova Ação separado
        dbc.Modal(
            [
                dbc.ModalHeader("Cadastrar Nova Ação"),
                dbc.ModalBody([
                    dbc.Row([
                        dbc.Col([
                            html.Label("Projeto"),
                            dcc.Dropdown(id="modal-acao-projeto", options=[])
                        ], width=6),
                        dbc.Col([
                            html.Label("Mês de Referência"),
                            dcc.Dropdown(id="modal-acao-mes-referencia", options=[])
                        ], width=6),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Prioridade"),
                            dcc.Dropdown(
                                id="modal-acao-prioridade",
                                options=[
                                    {"label": "Baixa", "value": "Baixa"},
                                    {"label": "Média", "value": "Média"},
                                    {"label": "Alta", "value": "Alta"}
                                ],
                                value="Média"
                            )
                        ], width=6),
                        dbc.Col([
                            html.Label("Responsáveis"),
                            dcc.Dropdown(id="modal-acao-responsaveis", options=[], multi=True)
                        ], width=6),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Descrição da Ação"),
                            dbc.Textarea(id="modal-acao-descricao", rows=3)
                        ], width=12),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Data Limite"),
                            dcc.DatePickerSingle(id="modal-acao-data-limite", date=None)
                        ], width=4),
                        dbc.Col([
                            html.Label("Status"),
                            dcc.Dropdown(
                                id="modal-acao-status",
                                options=[
                                    {"label": "Pendente", "value": "Pendente"},
                                    {"label": "Em Andamento", "value": "Em Andamento"},
                                    {"label": "Concluída", "value": "Concluída"}
                                ],
                                value="Pendente"
                            )
                        ], width=4),
                        dbc.Col([
                            html.Label("Data de Conclusão"),
                            dcc.DatePickerSingle(id="modal-acao-data-conclusao", date=None)
                        ], width=4),
                    ]),
                    dbc.Alert("Preencha todos os campos obrigatórios", id="modal-acao-alert-text", color="danger", is_open=False)
                ]),
                dbc.ModalFooter([
                    dbc.Button("Cancelar", id="modal-acao-cancel", color="secondary", className="me-2"),
                    dbc.Button("Salvar", id="modal-acao-save", color="primary", style={'backgroundColor': codeart_colors['dark_blue']})
                ]),
            ],
            id="modal-nova-acao",
            size="lg",
            is_open=False,
        ),
        
        # Stores para dados
        dcc.Store(id="raw-data-store", data=df_projetos_initial.to_dict('records')),
        dcc.Store(id="codenautas-store", data=df_codenautas_initial.to_dict('records')),
        dcc.Store(id="acoes-store", data=df_acoes_initial.to_dict('records')),
        dcc.Store(id="filter-options-store", data={
            "meses_anos": meses_anos_initial,
            "gestoras": gestoras_initial,
            "status_list": status_list_initial,
            "segmentos": segmentos_initial,
            "tipos": tipos_initial,
            "coordenacoes": coordenacoes_initial,
            "financeiro_list": financeiro_list_initial
        }),
        dcc.Store(id="active-tab-store", data="tab-projetos"),
        dcc.Store(id="selected-gestora-store", data=None),
        dcc.Store(id="selected-status-store", data=None),
        dcc.Store(id="selected-financeiro-store", data=None),
        dcc.Store(id="selected-nps-store", data=None),
        dcc.Store(id="selected-project-store", data=None),
        
        # Componente invisível para evitar erro de callback
        html.Div(dcc.Graph(id="segmento-chart"), style={'display': 'none'})
    ], fluid=True)
])

# Callback para atualizar a hora da última atualização
@app.callback(
    Output("last-update-time", "children"),
    Input("refresh-data-button", "n_clicks"),
    prevent_initial_call=True
)
def update_time(n_clicks):
    if n_clicks:
        now = datetime.now()
        return f"Última atualização: {now.strftime('%d/%m/%Y %H:%M:%S')}"
    return ""

# Callback para atualizar os dados quando o botão de atualização é clicado
@app.callback(
    [
        Output("raw-data-store", "data"),
        Output("codenautas-store", "data"),
        Output("acoes-store", "data")
    ],
    Input("refresh-data-button", "n_clicks"),
    prevent_initial_call=True
)
def refresh_data(n_clicks):
    global CACHE_PROJETOS, CACHE_CODENAUTAS, CACHE_ACOES, LAST_CACHE_UPDATE
    
    if n_clicks:
        # Forçar atualização definindo o último cache como muito antigo
        LAST_CACHE_UPDATE = 0
        
        # Recarregar dados do Google Sheets
        df_projetos_refreshed = load_data_from_sheets()
        df_projetos_refreshed = process_data(df_projetos_refreshed)
        
        # Salvar cópia local
        if not df_projetos_refreshed.empty:
            save_data_to_local(df_projetos_refreshed, "projetos")
        elif CACHE_PROJETOS is None:
            # Tentar carregar do backup local
            df_local = load_data_from_local("projetos")
            if df_local is not None:
                df_projetos_refreshed = process_data(df_local)
        
        # Recarregar dados dos codenautas
        df_codenautas_refreshed = load_codenautas_from_sheets()
        
        # Salvar cópia local
        if not df_codenautas_refreshed.empty:
            save_data_to_local(df_codenautas_refreshed, "codenautas")
        elif CACHE_CODENAUTAS is None:
            # Tentar carregar do backup local
            df_local = load_data_from_local("codenautas")
            if df_local is not None:
                df_codenautas_refreshed = df_local
        
        # Recarregar dados das ações
        df_acoes_refreshed = load_acoes_from_sheets()
        df_acoes_refreshed = process_acoes(df_acoes_refreshed)
        
        # Salvar cópia local
        if not df_acoes_refreshed.empty:
            save_data_to_local(df_acoes_refreshed, "acoes")
        elif CACHE_ACOES is None:
            # Tentar carregar do backup local
            df_local = load_data_from_local("acoes")
            if df_local is not None:
                df_acoes_refreshed = process_acoes(df_local)
        
        return df_projetos_refreshed.to_dict('records'), df_codenautas_refreshed.to_dict('records'), df_acoes_refreshed.to_dict('records')
    
    # Se não houver clique, retornar os dados atuais
    return dash.no_update, dash.no_update, dash.no_update

# Callback para atualizar métricas e gráficos
@app.callback(
    [
        Output("total-projetos", "children"),
        Output("total-clientes", "children"),
        Output("projetos-atrasados", "children"),
        Output("projetos-criticos", "children"),
        Output("status-chart", "figure"),
        Output("projetos-table", "data"),
        Output("financeiro-chart", "figure"),
        Output("nps-chart", "figure"),
        Output("segmento-chart", "figure"),
        Output("projetos-gp-chart", "figure"),
        # Novos gráficos adicionados
        Output("horas-chart", "figure"),
        Output("saldo-chart", "figure"),
        Output("atraso-coordenacao-chart", "figure"),
        Output("evolucao-quitados-chart", "figure"),
        Output("evolucao-atrasados-chart", "figure")
    ],
    Input("raw-data-store", "data")
)
def update_dashboard(data):
    # Converter dados para DataFrame
    df = pd.DataFrame(data) if data else pd.DataFrame()
    
    # Se o DataFrame estiver vazio, retornar valores vazios
    if df.empty:
        empty_fig = go.Figure().update_layout(title="Sem dados disponíveis")
        return "0", "0", "0", "0", empty_fig, [], empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig
    
    # Calcular métricas
    total_projetos = len(df)
    total_clientes = len(df['Cliente'].unique()) if 'Cliente' in df.columns else 0
    projetos_atrasados = len(df[df['Status'] == 'Atrasado']) if 'Status' in df.columns else 0
    projetos_criticos = len(df[df['Prioridade'] == 'Crítico']) if 'Prioridade' in df.columns else 0
    
    # Criar gráfico de status
    status_counts = df['Status'].value_counts().reset_index() if 'Status' in df.columns else pd.DataFrame(columns=['Status', 'Quantidade'])
    status_counts.columns = ['Status', 'Quantidade']
    
    status_fig = px.pie(
        status_counts, names='Status', values='Quantidade',
        title='Distribuição por Status',
        color_discrete_sequence=codeart_chart_palette,
    )
    status_fig.update_traces(textposition='inside', textinfo='percent+label')
    
    # Criar gráfico de financeiro
    financeiro_counts = df['Financeiro'].value_counts().reset_index() if 'Financeiro' in df.columns else pd.DataFrame(columns=['Financeiro', 'Quantidade'])
    financeiro_counts.columns = ['Financeiro', 'Quantidade']
    
    financeiro_fig = px.pie(
        financeiro_counts, names='Financeiro', values='Quantidade',
        title='Distribuição por Status Financeiro',
        color_discrete_sequence=codeart_chart_palette,
    )
    financeiro_fig.update_traces(textposition='inside', textinfo='percent+label')
    
    # Criar gráfico de NPS
    nps_counts = df['NPS '].value_counts().reset_index() if 'NPS ' in df.columns else pd.DataFrame(columns=['NPS', 'Quantidade'])
    nps_counts.columns = ['NPS', 'Quantidade']
    
    nps_fig = px.pie(
        nps_counts, names='NPS', values='Quantidade',
        title='Distribuição por NPS',
        color_discrete_sequence=codeart_chart_palette,
    )
    nps_fig.update_traces(textposition='inside', textinfo='percent+label')
    
    # Criar gráfico de Segmento
    segmento_counts = df['Segmento'].value_counts().reset_index() if 'Segmento' in df.columns else pd.DataFrame(columns=['Segmento', 'Quantidade'])
    segmento_counts.columns = ['Segmento', 'Quantidade']
    segmento_counts = segmento_counts.sort_values('Quantidade', ascending=False)
    
    # Verificar se há dados
    if len(segmento_counts) > 0:
        segmento_fig = px.bar(
            segmento_counts, x='Segmento', y='Quantidade',
            title='Distribuição por Segmento',
            color_discrete_sequence=[codeart_colors['blue_sky']],
            text_auto=True
        )
        segmento_fig.update_traces(textposition='outside')
    else:
        # Criar figura vazia
        segmento_fig = go.Figure()
        segmento_fig.update_layout(
            title="Sem dados de segmento disponíveis",
            xaxis_title="Segmento",
            yaxis_title="Quantidade"
        )
    
    # Criar gráfico de GP Responsável
    gp_counts = df['GP Responsável'].value_counts().reset_index() if 'GP Responsável' in df.columns else pd.DataFrame(columns=['GP Responsável', 'Quantidade'])
    gp_counts.columns = ['GP Responsável', 'Quantidade']
    
    gp_fig = px.bar(
        gp_counts, x='GP Responsável', y='Quantidade',
        title='Projetos por Gestora',
        color_discrete_sequence=[codeart_colors['blue_sky']],
        text_auto=True
    )
    gp_fig.update_traces(textposition='outside')
    
    # NOVOS GRÁFICOS
    
    # Gráfico de Horas Previstas vs Realizadas
    horas_fig = go.Figure()
    if 'Previsão' in df.columns and 'Real' in df.columns and 'Projeto' in df.columns:
        # Selecionar top 10 projetos por horas previstas
        top_projetos = df.sort_values('Previsão', ascending=False).head(10)
        
        horas_fig = go.Figure()
        horas_fig.add_trace(go.Bar(
            x=top_projetos['Projeto'],
            y=top_projetos['Previsão'],
            name='Horas Previstas',
            marker_color=codeart_colors['blue_sky'],
            text=top_projetos['Previsão'].round(1),
            textposition='outside'
        ))
        horas_fig.add_trace(go.Bar(
            x=top_projetos['Projeto'],
            y=top_projetos['Real'],
            name='Horas Realizadas',
            marker_color=codeart_colors['dark_blue'],
            text=top_projetos['Real'].round(1),
            textposition='outside'
        ))
        
        horas_fig.update_layout(
            title='Top 10 Projetos: Horas Previstas vs Realizadas',
            barmode='group',
            xaxis_tickangle=-45
        )
    else:
        horas_fig.update_layout(title="Sem dados de horas")
    
    # Gráfico de Saldo de Horas
    saldo_fig = go.Figure()
    if 'Saldo Acumulado' in df.columns and 'Projeto' in df.columns:
        # Filtrar projetos com saldo não zero
        df_saldo = df[df['Saldo Acumulado'] != 0].copy()
        
        if not df_saldo.empty:
            # Ordenar por saldo (do menor para o maior)
            df_saldo = df_saldo.sort_values('Saldo Acumulado')
            
            # Limitar a 15 projetos para melhor visualização
            if len(df_saldo) > 15:
                df_saldo = pd.concat([df_saldo.head(7), df_saldo.tail(8)])
            
            # Definir cores baseadas no saldo
            colors = ['#dc3545' if x < 0 else '#28a745' for x in df_saldo['Saldo Acumulado']]
            
            saldo_fig = go.Figure(data=[go.Bar(
                x=df_saldo['Projeto'],
                y=df_saldo['Saldo Acumulado'],
                marker_color=colors,
                text=df_saldo['Saldo Acumulado'].round(1),
                textposition='outside'
            )])
            
            saldo_fig.update_layout(
                title='Saldo de Horas por Projeto',
                xaxis_tickangle=-45
            )
        else:
            saldo_fig.update_layout(title="Sem projetos com saldo diferente de zero")
    else:
        saldo_fig.update_layout(title="Sem dados de saldo")
    
    # Gráfico de Atraso por Coordenação
    atraso_coord_fig = go.Figure()
    if 'Coordenação' in df.columns and 'Status' in df.columns:
        # Agrupar por coordenação e contar projetos atrasados
        atraso_coord_data = df.groupby('Coordenação').apply(
            lambda x: pd.Series({
                'Total Projetos': len(x),
                'Projetos Atrasados': len(x[x['Status'] == 'Atrasado'])
            })
        ).reset_index()
        
        if not atraso_coord_data.empty:
            # Calcular percentual de projetos atrasados
            atraso_coord_data['Percentual'] = (atraso_coord_data['Projetos Atrasados'] / atraso_coord_data['Total Projetos'] * 100).round(1)
            
            atraso_coord_fig = go.Figure()
            atraso_coord_fig.add_trace(go.Bar(
                x=atraso_coord_data['Coordenação'],
                y=atraso_coord_data['Projetos Atrasados'],
                name='Projetos Atrasados',
                marker_color=codeart_colors['danger'],
                text=atraso_coord_data['Projetos Atrasados'],
                textposition='outside'
            ))
            
            atraso_coord_fig.add_trace(go.Bar(
                x=atraso_coord_data['Coordenação'],
                y=atraso_coord_data['Total Projetos'] - atraso_coord_data['Projetos Atrasados'],
                name='Projetos no Prazo',
                marker_color=codeart_colors['success'],
                text=atraso_coord_data['Total Projetos'] - atraso_coord_data['Projetos Atrasados'],
                textposition='outside'
            ))
            
            atraso_coord_fig.update_layout(
                title='Projetos Atrasados por Coordenação',
                barmode='stack',
                xaxis_tickangle=-45
            )
        else:
            atraso_coord_fig.update_layout(title="Sem dados de atraso por coordenação")
    else:
        atraso_coord_fig.update_layout(title="Sem dados de coordenação ou status")
    
    # Gráfico de Evolução de Projetos Quitados
    evolucao_quitados_fig = go.Figure()
    if 'Financeiro' in df.columns and 'MesAnoFormatado' in df.columns:
        # Contar projetos quitados por mês
        quitados_por_mes = df.groupby('MesAnoFormatado').apply(
            lambda x: len(x[x['Financeiro'] == 'Quitado'])
        ).reset_index()
        quitados_por_mes.columns = ['MesAnoFormatado', 'Projetos Quitados']
        
        if not quitados_por_mes.empty:
            evolucao_quitados_fig = px.line(
                quitados_por_mes, x='MesAnoFormatado', y='Projetos Quitados',
                title='Evolução de Projetos Quitados',
                markers=True,
                color_discrete_sequence=[codeart_colors['success']]
            )
            
            evolucao_quitados_fig.update_layout(xaxis_tickangle=-45)
        else:
            evolucao_quitados_fig.update_layout(title="Sem dados de projetos quitados")
    else:
        evolucao_quitados_fig.update_layout(title="Sem dados de financeiro ou período")
    
    # Gráfico de Evolução de Projetos Atrasados
    evolucao_atrasados_fig = go.Figure()
    if 'Status' in df.columns and 'MesAnoFormatado' in df.columns:
        # Contar projetos atrasados por mês
        atrasados_por_mes = df.groupby('MesAnoFormatado').apply(
            lambda x: len(x[x['Status'] == 'Atrasado'])
        ).reset_index()
        atrasados_por_mes.columns = ['MesAnoFormatado', 'Projetos Atrasados']
        
        if not atrasados_por_mes.empty:
            evolucao_atrasados_fig = px.line(
                atrasados_por_mes, x='MesAnoFormatado', y='Projetos Atrasados',
                title='Evolução de Projetos Atrasados',
                markers=True,
                color_discrete_sequence=[codeart_colors['danger']]
            )
            
            evolucao_atrasados_fig.update_layout(xaxis_tickangle=-45)
        else:
            evolucao_atrasados_fig.update_layout(title="Sem dados de projetos atrasados")
    else:
        evolucao_atrasados_fig.update_layout(title="Sem dados de status ou período")
    
    # Retornar dados da tabela
    table_data = df.to_dict('records')
    
    return str(total_projetos), str(total_clientes), str(projetos_atrasados), str(projetos_criticos), status_fig, table_data, financeiro_fig, nps_fig, segmento_fig, gp_fig, horas_fig, saldo_fig, atraso_coord_fig, evolucao_quitados_fig, evolucao_atrasados_fig

# Armazenar a aba ativa
@app.callback(
    Output("active-tab-store", "data"),
    Input("tabs", "active_tab")
)
def store_active_tab(active_tab):
    return active_tab

# Callback para preencher as opções do dropdown de responsáveis nas ações
@app.callback(
    Output("responsavel-filter-acoes", "options"),
    Input("codenautas-store", "data")
)
def update_responsaveis_filter_options(codenautas_data):
    if not codenautas_data:
        return []
    df_codenautas = pd.DataFrame(codenautas_data)
    if df_codenautas.empty or 'Nome' not in df_codenautas.columns:
        return []
    options = [{"label": nome, "value": nome} for nome in sorted(df_codenautas['Nome'].unique())]
    return options

# Callback para preencher as opções do dropdown de responsáveis no modal
@app.callback(
    Output("modal-responsaveis", "options"),
    Input("codenautas-store", "data")
)
def update_responsaveis_options(codenautas_data):
    if not codenautas_data:
        return []
    df_codenautas = pd.DataFrame(codenautas_data)
    if df_codenautas.empty or 'Nome' not in df_codenautas.columns:
        return []
    options = [{"label": nome, "value": nome} for nome in sorted(df_codenautas['Nome'].unique())]
    return options

# Callback para preencher as opções do dropdown de responsáveis no modal de edição
@app.callback(
    Output("modal-edit-responsaveis", "options"),
    Input("codenautas-store", "data")
)
def update_edit_responsaveis_options(codenautas_data):
    if not codenautas_data:
        return []
    df_codenautas = pd.DataFrame(codenautas_data)
    if df_codenautas.empty or 'Nome' not in df_codenautas.columns:
        return []
    options = [{"label": nome, "value": nome} for nome in sorted(df_codenautas['Nome'].unique())]
    return options

# Callback para preencher as opções do dropdown de responsáveis no modal de nova ação
@app.callback(
    Output("modal-acao-responsaveis", "options"),
    Input("codenautas-store", "data")
)
def update_acao_responsaveis_options(codenautas_data):
    if not codenautas_data:
        return []
    df_codenautas = pd.DataFrame(codenautas_data)
    if df_codenautas.empty or 'Nome' not in df_codenautas.columns:
        return []
    options = [{"label": nome, "value": nome} for nome in sorted(df_codenautas['Nome'].unique())]
    return options

# Callback para atualizar os dropdowns de filtro com base nas opções disponíveis
@app.callback(
    [
        Output("coordenacao-filter", "options"),
        Output("mes-ano-filter", "options"),
        Output("gestora-filter", "options"),
        Output("status-filter", "options"),
        Output("financeiro-filter", "options"),
        Output("segmento-filter", "options"),
        Output("tipo-filter", "options"),
        Output("modal-mes-referencia", "options")
    ],
    Input("filter-options-store", "data")
)
def update_dropdown_options(filter_options_data):
    if not filter_options_data:
        return [], [], [], [], [], [], [], []
    
    coordenacoes = [{"label": coord, "value": coord} for coord in filter_options_data.get("coordenacoes", [])]
    meses_anos = [{"label": mes, "value": mes} for mes in filter_options_data.get("meses_anos", [])]
    gestoras = [{"label": gp, "value": gp} for gp in filter_options_data.get("gestoras", [])]
    status_list = [{"label": status, "value": status} for status in filter_options_data.get("status_list", [])]
    financeiro_list = [{"label": fin, "value": fin} for fin in filter_options_data.get("financeiro_list", [])]
    segmentos = [{"label": seg, "value": seg} for seg in filter_options_data.get("segmentos", [])]
    tipos = [{"label": tipo, "value": tipo} for tipo in filter_options_data.get("tipos", [])]
    
    return coordenacoes, meses_anos, gestoras, status_list, financeiro_list, segmentos, tipos, meses_anos

# Callback para limpar filtros de projetos
@app.callback(
    [
        Output("coordenacao-filter", "value"),
        Output("mes-ano-filter", "value"),
        Output("gestora-filter", "value"),
        Output("status-filter", "value"),
        Output("financeiro-filter", "value"),
        Output("segmento-filter", "value"),
        Output("tipo-filter", "value"),
        # Limpar também os stores de seleção de gráficos
        Output("selected-gestora-store", "data", allow_duplicate=True),
        Output("selected-status-store", "data", allow_duplicate=True),
        Output("selected-financeiro-store", "data", allow_duplicate=True),
        Output("selected-nps-store", "data", allow_duplicate=True)
    ],
    Input("reset-project-filters", "n_clicks"),
    State("filter-options-store", "data"),
    prevent_initial_call=True
)
def reset_filters(n_clicks, filter_options_data):
    return None, None, None, None, None, None, None, None, None, None, None

# Callback para limpar filtros de ações
@app.callback(
    [
        Output("mes-ano-filter-acoes", "value"),
        Output("responsavel-filter-acoes", "value"),
        Output("status-filter-acoes", "value"),
        Output("prioridade-filter-acoes", "value")
    ],
    Input("reset-acoes-filters", "n_clicks"),
    prevent_initial_call=True
)
def reset_acoes_filters(n_clicks):
    return None, None, None, None

# Callback para preencher as opções do dropdown de projetos no modal
@app.callback(
    Output("modal-projeto", "options"),
    Input("raw-data-store", "data")
)
def update_projetos_options(data):
    if not data:
        return []
    df = pd.DataFrame(data)
    if df.empty or 'Projeto' not in df.columns:
        return []
    options = [{"label": projeto, "value": projeto} for projeto in sorted(df['Projeto'].unique())]
    return options

# Callback para preencher as opções do dropdown de projetos no modal de edição
@app.callback(
    Output("modal-edit-projeto", "options"),
    Input("raw-data-store", "data")
)
def update_edit_projetos_options(data):
    if not data:
        return []
    df = pd.DataFrame(data)
    if df.empty or 'Projeto' not in df.columns:
        return []
    options = [{"label": projeto, "value": projeto} for projeto in sorted(df['Projeto'].unique())]
    return options

# Callback para preencher as opções do dropdown de projetos no modal de nova ação
@app.callback(
    Output("modal-acao-projeto", "options"),
    Input("raw-data-store", "data")
)
def update_acao_projetos_options(data):
    if not data:
        return []
    df = pd.DataFrame(data)
    if df.empty or 'Projeto' not in df.columns:
        return []
    options = [{"label": projeto, "value": projeto} for projeto in sorted(df['Projeto'].unique())]
    return options

# Callback para preencher as opções do dropdown de mês de referência no modal de nova ação
@app.callback(
    Output("modal-acao-mes-referencia", "options"),
    Input("filter-options-store", "data")
)
def update_acao_mes_referencia_options(filter_options_data):
    if not filter_options_data:
        return []
    meses_anos = [{"label": mes, "value": mes} for mes in filter_options_data.get("meses_anos", [])]
    return meses_anos

# Callback para atualizar métricas e gráficos da aba Ações
@app.callback(
    [
        Output("total-acoes", "children"),
        Output("acoes-pendentes", "children"),
        Output("acoes-concluidas", "children"),
        Output("acoes-atrasadas", "children"),
        Output("status-acoes-chart", "figure"),
        Output("prioridade-acoes-chart", "figure"),
        Output("responsaveis-acoes-chart", "figure"),
        Output("evolucao-acoes-chart", "figure"),
        Output("acoes-table", "data")
    ],
    [
        Input("apply-acoes-filters", "n_clicks"),
        Input("reset-acoes-filters", "n_clicks"),
        Input("active-tab-store", "data"),
        Input("acoes-store", "data")
    ],
    [
        State("mes-ano-filter-acoes", "value"),
        State("responsavel-filter-acoes", "value"),
        State("status-filter-acoes", "value"),
        State("prioridade-filter-acoes", "value")
    ]
)
def update_acoes_dashboard(n_clicks_apply, n_clicks_reset, active_tab, acoes_data, mes_ano, responsavel, status, prioridade):
    # Criar figura vazia para usar como padrão
    empty_fig = go.Figure().update_layout(title="Sem dados disponíveis")
    
    # Se não houver dados, retornar valores vazios
    if not acoes_data:
        return "0", "0", "0", "0", empty_fig, empty_fig, empty_fig, empty_fig, []
    
    # Converter para DataFrame
    df_acoes = pd.DataFrame(acoes_data)
    
    # Se o DataFrame estiver vazio, retornar valores vazios
    if df_acoes.empty:
        return "0", "0", "0", "0", empty_fig, empty_fig, empty_fig, empty_fig, []
    
    # Verificar se estamos na aba de ações
    if active_tab != "tab-acoes":
        # Se não estamos na aba de ações, apenas retornamos os valores atuais
        # para evitar cálculos desnecessários
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update
    
    # Aplicar filtros se o botão de aplicar filtros foi clicado e há filtros selecionados
    ctx = dash.callback_context
    if ctx.triggered:
        button_id = ctx.triggered[0]['prop_id'].split('.')[0]
        if button_id == "apply-acoes-filters" and (mes_ano or responsavel or status or prioridade):
            filtered_df = df_acoes.copy()
            
            # Filtrar por mês/ano
            if mes_ano:
                if not isinstance(mes_ano, list):
                    mes_ano = [mes_ano]
                filtered_df = filtered_df[filtered_df['Mês de Referência'].isin(mes_ano)]
            
            # Filtrar por responsável
            if responsavel:
                if not isinstance(responsavel, list):
                    responsavel = [responsavel]
                # Considerando que um responsável pode estar em uma lista separada por vírgulas ou como string única
                mask = filtered_df['Responsáveis'].apply(
                    lambda x: any(resp in str(x).split(',') for resp in responsavel)
                )
                filtered_df = filtered_df[mask]
            
            # Filtrar por status
            if status:
                if not isinstance(status, list):
                    status = [status]
                filtered_df = filtered_df[filtered_df['Status'].isin(status)]
            
            # Filtrar por prioridade
            if prioridade:
                if not isinstance(prioridade, list):
                    prioridade = [prioridade]
                filtered_df = filtered_df[filtered_df['Prioridade'].isin(prioridade)]
        else:
            filtered_df = df_acoes
    else:
        filtered_df = df_acoes
    
    # Calcular métricas
    total_acoes = len(filtered_df)
    pendentes = len(filtered_df[filtered_df['Status'] != 'Concluída'])
    concluidas = len(filtered_df[filtered_df['Status'] == 'Concluída'])
    atrasadas = filtered_df['Atrasada'].sum() if 'Atrasada' in filtered_df.columns else 0
    
    # Criar gráfico de status
    status_counts = filtered_df['Status'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Quantidade']
    
    status_fig = px.pie(
        status_counts, names='Status', values='Quantidade',
        title='Distribuição por Status',
        color_discrete_sequence=codeart_chart_palette,
    )
    status_fig.update_traces(textposition='inside', textinfo='percent+label')
    
    # Criar gráfico de prioridade
    prioridade_counts = filtered_df['Prioridade'].value_counts().reset_index()
    prioridade_counts.columns = ['Prioridade', 'Quantidade']
    
    prioridade_fig = px.pie(
        prioridade_counts, names='Prioridade', values='Quantidade',
        title='Distribuição por Prioridade',
        color_discrete_sequence=codeart_chart_palette,
    )
    prioridade_fig.update_traces(textposition='inside', textinfo='percent+label')
    
    # Criar gráfico de responsáveis
    # Precisamos primeiro expandir a coluna de responsáveis que pode conter múltiplos valores
    responsaveis_expandidos = []
    for idx, row in filtered_df.iterrows():
        if pd.notna(row['Responsáveis']):
            for resp in str(row['Responsáveis']).split(','):
                resp = resp.strip()
                if resp:
                    responsaveis_expandidos.append(resp)
    
    if responsaveis_expandidos:
        responsaveis_counts = pd.Series(responsaveis_expandidos).value_counts().reset_index()
        responsaveis_counts.columns = ['Responsável', 'Quantidade']
        
        responsaveis_fig = px.bar(
            responsaveis_counts, x='Responsável', y='Quantidade',
            title='Distribuição por Responsável',
            color_discrete_sequence=[codeart_colors['blue_sky']],
            text_auto=True
        )
        responsaveis_fig.update_traces(textposition='outside')
    else:
        responsaveis_fig = go.Figure().update_layout(title="Sem dados de responsáveis")
    
    # Criar gráfico de evolução de ações ao longo do tempo
    # Considerando a Data de Cadastro para agrupar
    if 'Data de Cadastro' in filtered_df.columns:
        # Garantir que a coluna é datetime
        filtered_df['Data de Cadastro'] = pd.to_datetime(filtered_df['Data de Cadastro'], errors='coerce')
        
        # Criar coluna de mês/ano para agrupamento
        filtered_df['Mês Cadastro'] = filtered_df['Data de Cadastro'].dt.strftime('%Y-%m')
        
        # Agrupar por mês e contar
        evolucao = filtered_df.groupby('Mês Cadastro').size().reset_index()
        evolucao.columns = ['Mês', 'Quantidade']
        
        # Ordenar cronologicamente
        evolucao = evolucao.sort_values('Mês')
        
        evolucao_fig = px.line(
            evolucao, x='Mês', y='Quantidade',
            title='Evolução de Ações Cadastradas',
            color_discrete_sequence=[codeart_colors['dark_blue']],
        )
        evolucao_fig.update_traces(mode='lines+markers')
    else:
        evolucao_fig = go.Figure().update_layout(title="Sem dados de evolução")
    
    # Retornar dados da tabela
    table_data = filtered_df.to_dict('records')
    
    return str(total_acoes), str(pendentes), str(concluidas), str(atrasadas), status_fig, prioridade_fig, responsaveis_fig, evolucao_fig, table_data

# Callback para atualizar métricas e gráficos da aba Projetos com filtros
@app.callback(
    [
        Output("total-projetos", "children", allow_duplicate=True),
        Output("total-clientes", "children", allow_duplicate=True),
        Output("projetos-atrasados", "children", allow_duplicate=True),
        Output("projetos-criticos", "children", allow_duplicate=True),
        Output("projetos-table", "data", allow_duplicate=True),
        Output("status-chart", "figure", allow_duplicate=True),
        Output("financeiro-chart", "figure", allow_duplicate=True),
        Output("nps-chart", "figure", allow_duplicate=True),
        Output("segmento-chart", "figure", allow_duplicate=True),
        Output("projetos-gp-chart", "figure", allow_duplicate=True),
        # Novos gráficos adicionados
        Output("horas-chart", "figure", allow_duplicate=True),
        Output("saldo-chart", "figure", allow_duplicate=True),
        Output("atraso-coordenacao-chart", "figure", allow_duplicate=True),
        Output("evolucao-quitados-chart", "figure", allow_duplicate=True),
        Output("evolucao-atrasados-chart", "figure", allow_duplicate=True)
    ],
    [
        Input("apply-project-filters", "n_clicks"),
        Input("reset-project-filters", "n_clicks"),
    ],
    [
        State("mes-ano-filter", "value"),
        State("gestora-filter", "value"),
        State("status-filter", "value"),
        State("segmento-filter", "value"),
        State("tipo-filter", "value"),
        State("coordenacao-filter", "value"),
        State("financeiro-filter", "value"),
        State("raw-data-store", "data"),
    ],
    prevent_initial_call=True
)
def update_dashboard_with_filters(n_clicks_apply, n_clicks_reset, mes_ano, gestora, status, segmento, tipo, coordenacao, financeiro, data):
    # Criar figura vazia para usar como padrão
    empty_fig = go.Figure().update_layout(title="Sem dados disponíveis")
    
    # Se não houver dados, retornar valores vazios
    if not data:
        return "0", "0", "0", "0", [], empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig
    
    # Converter para DataFrame
    df = pd.DataFrame(data)
    
    # Se o DataFrame estiver vazio, retornar valores vazios
    if df.empty:
        return "0", "0", "0", "0", [], empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig
    
    # Verificar qual botão foi clicado
    ctx = dash.callback_context
    button_id = ctx.triggered[0]['prop_id'].split('.')[0] if ctx.triggered else None
    
    # Se o botão de reset foi clicado, usar o DataFrame original
    if button_id == "reset-project-filters":
        filtered_df = df
    # Se o botão de aplicar filtros foi clicado, aplicar os filtros
    elif button_id == "apply-project-filters" and (mes_ano or gestora or status or segmento or tipo or coordenacao or financeiro):
        filtered_df = df.copy()
        
        # Filtrar por mês/ano
        if mes_ano:
            if not isinstance(mes_ano, list):
                mes_ano = [mes_ano]
            filtered_df = filtered_df[filtered_df['MesAnoFormatado'].isin(mes_ano)]
        
        # Filtrar por gestora
        if gestora:
            if not isinstance(gestora, list):
                gestora = [gestora]
            filtered_df = filtered_df[filtered_df['GP Responsável'].isin(gestora)]
        
        # Filtrar por status
        if status:
            if not isinstance(status, list):
                status = [status]
            filtered_df = filtered_df[filtered_df['Status'].isin(status)]
        
        # Filtrar por segmento
        if segmento:
            if not isinstance(segmento, list):
                segmento = [segmento]
            filtered_df = filtered_df[filtered_df['Segmento'].isin(segmento)]
        
        # Filtrar por tipo
        if tipo:
            if not isinstance(tipo, list):
                tipo = [tipo]
            filtered_df = filtered_df[filtered_df['Tipo'].isin(tipo)]
        
        # Filtrar por coordenação
        if coordenacao:
            if not isinstance(coordenacao, list):
                coordenacao = [coordenacao]
            filtered_df = filtered_df[filtered_df['Coordenação'].isin(coordenacao)]
        
        # Filtrar por financeiro
        if financeiro:
            if not isinstance(financeiro, list):
                financeiro = [financeiro]
            filtered_df = filtered_df[filtered_df['Financeiro'].isin(financeiro)]
    else:
        filtered_df = df
    
    # Calcular métricas
    total_projetos = len(filtered_df)
    total_clientes = len(filtered_df['Cliente'].unique())
    projetos_atrasados = len(filtered_df[filtered_df['Status'] == 'Atrasado'])
    projetos_criticos = len(filtered_df[filtered_df['Prioridade'] == 'Crítico'])
    
    # Criar gráfico de status
    status_counts = filtered_df['Status'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Quantidade']
    
    status_fig = px.pie(
        status_counts, names='Status', values='Quantidade',
        title='Distribuição por Status',
        color_discrete_sequence=codeart_chart_palette,
    )
    status_fig.update_traces(textposition='inside', textinfo='percent+label')
    
    # Criar gráfico de financeiro
    financeiro_counts = filtered_df['Financeiro'].value_counts().reset_index()
    financeiro_counts.columns = ['Financeiro', 'Quantidade']
    
    financeiro_fig = px.pie(
        financeiro_counts, names='Financeiro', values='Quantidade',
        title='Distribuição por Status Financeiro',
        color_discrete_sequence=codeart_chart_palette,
    )
    financeiro_fig.update_traces(textposition='inside', textinfo='percent+label')
    
    # Criar gráfico de NPS
    nps_counts = filtered_df['NPS '].value_counts().reset_index()
    nps_counts.columns = ['NPS', 'Quantidade']
    
    nps_fig = px.pie(
        nps_counts, names='NPS', values='Quantidade',
        title='Distribuição por NPS',
        color_discrete_sequence=codeart_chart_palette,
    )
    nps_fig.update_traces(textposition='inside', textinfo='percent+label')
    
    # Criar gráfico de Segmento
    segmento_counts = filtered_df['Segmento'].value_counts().reset_index() if 'Segmento' in filtered_df.columns else pd.DataFrame(columns=['Segmento', 'Quantidade'])
    segmento_counts.columns = ['Segmento', 'Quantidade']
    segmento_counts = segmento_counts.sort_values('Quantidade', ascending=False)
    
    segmento_fig = px.bar(
        segmento_counts, x='Segmento', y='Quantidade',
        title='Distribuição por Segmento',
        color_discrete_sequence=[codeart_colors['blue_sky']],
        text_auto=True
    )
    segmento_fig.update_traces(textposition='outside')
    
    # Criar gráfico de GP Responsável
    gp_counts = filtered_df['GP Responsável'].value_counts().reset_index()
    gp_counts.columns = ['GP Responsável', 'Quantidade']
    
    gp_fig = px.bar(
        gp_counts, x='GP Responsável', y='Quantidade',
        title='Projetos por Gestora',
        color_discrete_sequence=[codeart_colors['blue_sky']],
        text_auto=True
    )
    gp_fig.update_traces(textposition='outside')
    
    # NOVOS GRÁFICOS
    
    # Gráfico de Horas Previstas vs Realizadas
    horas_fig = go.Figure()
    if 'Previsão' in filtered_df.columns and 'Real' in filtered_df.columns and 'Projeto' in filtered_df.columns:
        # Selecionar top 10 projetos por horas previstas
        top_projetos = filtered_df.sort_values('Previsão', ascending=False).head(10)
        
        horas_fig = go.Figure()
        horas_fig.add_trace(go.Bar(
            x=top_projetos['Projeto'],
            y=top_projetos['Previsão'],
            name='Horas Previstas',
            marker_color=codeart_colors['blue_sky'],
            text=top_projetos['Previsão'].round(1),
            textposition='outside'
        ))
        horas_fig.add_trace(go.Bar(
            x=top_projetos['Projeto'],
            y=top_projetos['Real'],
            name='Horas Realizadas',
            marker_color=codeart_colors['dark_blue'],
            text=top_projetos['Real'].round(1),
            textposition='outside'
        ))
        
        horas_fig.update_layout(
            title='Top 10 Projetos: Horas Previstas vs Realizadas',
            barmode='group',
            xaxis_tickangle=-45
        )
    else:
        horas_fig.update_layout(title="Sem dados de horas")
    
    # Gráfico de Saldo de Horas
    saldo_fig = go.Figure()
    if 'Saldo Acumulado' in filtered_df.columns and 'Projeto' in filtered_df.columns:
        # Filtrar projetos com saldo não zero
        df_saldo = filtered_df[filtered_df['Saldo Acumulado'] != 0].copy()
        
        if not df_saldo.empty:
            # Ordenar por saldo (do menor para o maior)
            df_saldo = df_saldo.sort_values('Saldo Acumulado')
            
            # Limitar a 15 projetos para melhor visualização
            if len(df_saldo) > 15:
                df_saldo = pd.concat([df_saldo.head(7), df_saldo.tail(8)])
            
            # Definir cores baseadas no saldo
            colors = ['#dc3545' if x < 0 else '#28a745' for x in df_saldo['Saldo Acumulado']]
            
            saldo_fig = go.Figure(data=[go.Bar(
                x=df_saldo['Projeto'],
                y=df_saldo['Saldo Acumulado'],
                marker_color=colors,
                text=df_saldo['Saldo Acumulado'].round(1),
                textposition='outside'
            )])
            
            saldo_fig.update_layout(
                title='Saldo de Horas por Projeto',
                xaxis_tickangle=-45
            )
        else:
            saldo_fig.update_layout(title="Sem projetos com saldo diferente de zero")
    else:
        saldo_fig.update_layout(title="Sem dados de saldo")
    
    # Gráfico de Atraso por Coordenação
    atraso_coord_fig = go.Figure()
    if 'Coordenação' in filtered_df.columns and 'Status' in filtered_df.columns:
        # Agrupar por coordenação e contar projetos atrasados
        atraso_coord_data = filtered_df.groupby('Coordenação').apply(
            lambda x: pd.Series({
                'Total Projetos': len(x),
                'Projetos Atrasados': len(x[x['Status'] == 'Atrasado'])
            })
        ).reset_index()
        
        if not atraso_coord_data.empty:
            # Calcular percentual de projetos atrasados
            atraso_coord_data['Percentual'] = (atraso_coord_data['Projetos Atrasados'] / atraso_coord_data['Total Projetos'] * 100).round(1)
            
            atraso_coord_fig = go.Figure()
            atraso_coord_fig.add_trace(go.Bar(
                x=atraso_coord_data['Coordenação'],
                y=atraso_coord_data['Projetos Atrasados'],
                name='Projetos Atrasados',
                marker_color=codeart_colors['danger'],
                text=atraso_coord_data['Projetos Atrasados'],
                textposition='outside'
            ))
            
            atraso_coord_fig.add_trace(go.Bar(
                x=atraso_coord_data['Coordenação'],
                y=atraso_coord_data['Total Projetos'] - atraso_coord_data['Projetos Atrasados'],
                name='Projetos no Prazo',
                marker_color=codeart_colors['success'],
                text=atraso_coord_data['Total Projetos'] - atraso_coord_data['Projetos Atrasados'],
                textposition='outside'
            ))
            
            atraso_coord_fig.update_layout(
                title='Projetos Atrasados por Coordenação',
                barmode='stack',
                xaxis_tickangle=-45
            )
        else:
            atraso_coord_fig.update_layout(title="Sem dados de atraso por coordenação")
    else:
        atraso_coord_fig.update_layout(title="Sem dados de coordenação ou status")
    
    # Gráfico de Evolução de Projetos Quitados
    evolucao_quitados_fig = go.Figure()
    if 'Financeiro' in filtered_df.columns and 'MesAnoFormatado' in filtered_df.columns:
        # Contar projetos quitados por mês
        quitados_por_mes = filtered_df.groupby('MesAnoFormatado').apply(
            lambda x: len(x[x['Financeiro'] == 'Quitado'])
        ).reset_index()
        quitados_por_mes.columns = ['MesAnoFormatado', 'Projetos Quitados']
        
        if not quitados_por_mes.empty:
            evolucao_quitados_fig = px.line(
                quitados_por_mes, x='MesAnoFormatado', y='Projetos Quitados',
                title='Evolução de Projetos Quitados',
                markers=True,
                color_discrete_sequence=[codeart_colors['success']]
            )
            
            evolucao_quitados_fig.update_layout(xaxis_tickangle=-45)
        else:
            evolucao_quitados_fig.update_layout(title="Sem dados de projetos quitados")
    else:
        evolucao_quitados_fig.update_layout(title="Sem dados de financeiro ou período")
    
    # Gráfico de Evolução de Projetos Atrasados
    evolucao_atrasados_fig = go.Figure()
    if 'Status' in filtered_df.columns and 'MesAnoFormatado' in filtered_df.columns:
        # Contar projetos atrasados por mês
        atrasados_por_mes = filtered_df.groupby('MesAnoFormatado').apply(
            lambda x: len(x[x['Status'] == 'Atrasado'])
        ).reset_index()
        atrasados_por_mes.columns = ['MesAnoFormatado', 'Projetos Atrasados']
        
        if not atrasados_por_mes.empty:
            evolucao_atrasados_fig = px.line(
                atrasados_por_mes, x='MesAnoFormatado', y='Projetos Atrasados',
                title='Evolução de Projetos Atrasados',
                markers=True,
                color_discrete_sequence=[codeart_colors['danger']]
            )
            
            evolucao_atrasados_fig.update_layout(xaxis_tickangle=-45)
        else:
            evolucao_atrasados_fig.update_layout(title="Sem dados de projetos atrasados")
    else:
        evolucao_atrasados_fig.update_layout(title="Sem dados de status ou período")
    
    # Retornar dados da tabela
    table_data = filtered_df.to_dict('records')
    
    return str(total_projetos), str(total_clientes), str(projetos_atrasados), str(projetos_criticos), status_fig, table_data, financeiro_fig, nps_fig, segmento_fig, gp_fig, horas_fig, saldo_fig, atraso_coord_fig, evolucao_quitados_fig, evolucao_atrasados_fig

# Callback para adicionar ícone de ação na tabela
@app.callback(
    Output("projetos-table", "data", allow_duplicate=True),
    Input("raw-data-store", "data"),
    State("projetos-table", "data"),
    prevent_initial_call=True
)
def add_action_icon(raw_data, table_data):
    if not raw_data:
        return dash.no_update
    
    df = pd.DataFrame(raw_data)
    
    # Adicionar coluna de ícone de ação (usando texto simples em vez de markdown)
    df['action_icon'] = '+'
    
    return df.to_dict('records')

# Callback para capturar clique na célula do ícone de ação
@app.callback(
    [
        Output("selected-project-store", "data"),
        Output("modal-cadastro-acao", "is_open"),
        Output("modal-projeto", "value")
    ],
    Input("projetos-table", "selected_cells"),
    State("projetos-table", "data"),
    prevent_initial_call=True
)
def handle_action_icon_click(selected_cells, table_data):
    if not selected_cells or not table_data:
        return dash.no_update, dash.no_update, dash.no_update
    
    # Verificar se a célula selecionada é da coluna de ação
    if selected_cells[0]['column_id'] == 'action_icon':
        row_idx = selected_cells[0]['row']
        projeto = table_data[row_idx]['Projeto']
        return projeto, True, projeto
    
    return dash.no_update, dash.no_update, dash.no_update

# Callback para fechar o modal de cadastro de ação
@app.callback(
    [
        Output("modal-cadastro-acao", "is_open", allow_duplicate=True),
        Output("modal-alert-text", "is_open", allow_duplicate=True),
        Output("modal-alert-text", "children", allow_duplicate=True)
    ],
    Input("modal-cancel", "n_clicks"),
    prevent_initial_call=True
)
def close_modal(n_clicks):
    if n_clicks:
        return False, False, ""
    return dash.no_update, dash.no_update, dash.no_update

# Callback para salvar uma nova ação
@app.callback(
    [
        Output("modal-cadastro-acao", "is_open", allow_duplicate=True),
        Output("modal-alert-text", "is_open", allow_duplicate=True),
        Output("modal-alert-text", "children", allow_duplicate=True),
        Output("acoes-store", "data", allow_duplicate=True)
    ],
    Input("modal-save", "n_clicks"),
    [
        State("modal-projeto", "value"),
        State("modal-mes-referencia", "value"),
        State("modal-prioridade", "value"),
        State("modal-descricao", "value"),
        State("modal-responsaveis", "value"),
        State("modal-data-limite", "date"),
        State("modal-status", "value"),
        State("modal-data-conclusao", "date"),
        State("acoes-store", "data")
    ],
    prevent_initial_call=True
)
def save_action(n_clicks, projeto, mes_referencia, prioridade, descricao, responsaveis, data_limite, status, data_conclusao, acoes_data):
    if not n_clicks:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update
    
    # Validação melhorada para campos obrigatórios
    campos_vazios = []
    
    if not projeto:
        campos_vazios.append("Projeto")
        
    if not mes_referencia:
        campos_vazios.append("Mês de Referência")
        
    if not prioridade:
        campos_vazios.append("Prioridade")
        
    if not descricao:
        campos_vazios.append("Descrição")
    
    # Verificação especial para responsáveis que pode ser lista vazia, None, ou string vazia
    responsaveis_vazio = (responsaveis is None or 
                         (isinstance(responsaveis, list) and (len(responsaveis) == 0 or all(not r.strip() for r in responsaveis if isinstance(r, str)))) or
                         (isinstance(responsaveis, str) and not responsaveis.strip()))
    
    if responsaveis_vazio:
        campos_vazios.append("Responsáveis")
    
    if not data_limite:
        campos_vazios.append("Data Limite")
        
    if not status:
        campos_vazios.append("Status")
    
    # Se algum campo estiver vazio, exibir mensagem de alerta
    if campos_vazios:
        mensagem_erro = f"Por favor, preencha os seguintes campos obrigatórios: {', '.join(campos_vazios)}"
        print(f"Campos vazios no cadastro da ação: {', '.join(campos_vazios)}")
        return dash.no_update, True, mensagem_erro, dash.no_update
    
    try:
        # Converter responsáveis para string
        if isinstance(responsaveis, list):
            responsaveis_str = ', '.join(responsaveis)
        else:
            responsaveis_str = str(responsaveis)
            
        # Inicializar acoes_data como lista vazia se for None
        if acoes_data is None:
            acoes_data = []
        elif not isinstance(acoes_data, list):
            # Tentar converter para lista se não for
            try:
                acoes_data = list(acoes_data)
            except:
                acoes_data = []
                
        # Gerar ID da ação (próximo número sequencial)
        next_id = 1
        if acoes_data and len(acoes_data) > 0:
            df_acoes = pd.DataFrame(acoes_data)
            if 'ID da Ação' in df_acoes.columns:
                next_id = df_acoes['ID da Ação'].max() + 1 if not pd.isna(df_acoes['ID da Ação'].max()) else 1
            
        # Preparar nova linha
        nova_acao = {
            'ID da Ação': next_id,
            'Data de Cadastro': datetime.now().strftime('%Y-%m-%d'),
            'Mês de Referência': mes_referencia,
            'Projeto': projeto,
            'Descrição da Ação': descricao,
            'Responsáveis': responsaveis_str,
            'Data Limite': data_limite,
            'Status': status,
            'Prioridade': prioridade,
            'Data de Conclusão': data_conclusao,
            'Observações de conclusão': ""
        }
        
        # Adicionar nova ação aos dados existentes
        acoes_data.append(nova_acao)
        
        # Atualizar dados na planilha do Google Sheets
        df_acoes = pd.DataFrame(acoes_data)
        success = update_acoes_in_sheets(df_acoes)
        if success:
            print(f"Ação cadastrada com sucesso: ID {next_id}")
        else:
            print(f"AVISO: Falha ao salvar ação na planilha, mas foi salva localmente")
            save_data_to_local(df_acoes, "acoes")
        
        return False, False, "", acoes_data
    
    except Exception as e:
        print(f"Erro ao salvar ação: {e}")
        import traceback
        traceback.print_exc()
        return dash.no_update, True, f"Erro ao salvar ação: {str(e)}", dash.no_update

# Exportar a tabela de projetos
@app.callback(
    Output("download-dataframe-csv", "data"),
    Input("export-table-button", "n_clicks"),
    State("projetos-table", "data"),
    prevent_initial_call=True,
)
def export_table(n_clicks, table_data):
    if n_clicks is None or not table_data:
        return dash.no_update
        
    df = pd.DataFrame(table_data)
    # Remover a coluna de ícone de ação se existir
    if 'action_icon' in df.columns:
        df = df.drop(columns=['action_icon'])
    return dcc.send_data_frame(df.to_csv, "projetos_codeart.csv", index=False)

# Callback para edição de ação
# Se esse callback existe no código, adicione estas correções:
@app.callback(
    [
        Output("modal-edicao-acao", "is_open", allow_duplicate=True),
        Output("modal-edit-alert-text", "is_open", allow_duplicate=True),
        Output("modal-edit-alert-text", "children", allow_duplicate=True),
        Output("acoes-store", "data", allow_duplicate=True)
    ],
    Input("modal-edit-save", "n_clicks"),
    [
        State("modal-edit-id", "value"),
        State("modal-edit-projeto", "value"),
        State("modal-edit-mes-referencia", "value"),
        State("modal-edit-prioridade", "value"),
        State("modal-edit-descricao", "value"),
        State("modal-edit-responsaveis", "value"),
        State("modal-edit-data-limite", "date"),
        State("modal-edit-status", "value"),
        State("modal-edit-data-conclusao", "date"),
        State("modal-edit-observacoes", "value"),
        State("acoes-store", "data")
    ],
    prevent_initial_call=True
)
def save_action_edit(n_clicks, acao_id, projeto, mes_referencia, prioridade, descricao, responsaveis, data_limite, status, data_conclusao, observacoes, acoes_data):
    # Implementação atual permanece a mesma
    pass

# Callback para status com data de conclusão (modal)
@app.callback(
    Output("modal-edit-status", "value", allow_duplicate=True),
    Input("modal-edit-data-conclusao", "date"),
    State("modal-edit-status", "value"),
    prevent_initial_call=True
)
def update_status_on_conclusion_date_edit(data_conclusao, status_atual):
    # Implementação atual permanece a mesma
    pass

# Callback para status com data de conclusão (modal original)
@app.callback(
    Output("modal-status", "value", allow_duplicate=True),
    Input("modal-data-conclusao", "date"),
    State("modal-status", "value"),
    prevent_initial_call=True
)
def update_status_on_conclusion_date(data_conclusao, status_atual):
    # Implementação atual permanece a mesma
    pass

# Callback para status com data de conclusão (modal de ação)
@app.callback(
    Output("modal-acao-status", "value", allow_duplicate=True),
    Input("modal-acao-data-conclusao", "date"),
    State("modal-acao-status", "value"),
    prevent_initial_call=True
)
def update_status_on_conclusion_date_acao(data_conclusao, status_atual):
    # Implementação atual permanece a mesma
    pass

# Callback para salvar nova ação
@app.callback(
    [
        Output("modal-nova-acao", "is_open", allow_duplicate=True),
        Output("modal-acao-alert-text", "is_open", allow_duplicate=True),
        Output("modal-acao-alert-text", "children", allow_duplicate=True),
        Output("acoes-store", "data", allow_duplicate=True)
    ],
    Input("modal-acao-save", "n_clicks"),
    [
        State("modal-acao-projeto", "value"),
        State("modal-acao-mes-referencia", "value"),
        State("modal-acao-prioridade", "value"),
        State("modal-acao-descricao", "value"),
        State("modal-acao-responsaveis", "value"),
        State("modal-acao-data-limite", "date"),
        State("modal-acao-status", "value"),
        State("modal-acao-data-conclusao", "date"),
        State("acoes-store", "data")
    ],
    prevent_initial_call=True
)
def save_new_action(n_clicks, projeto, mes_referencia, prioridade, descricao, responsaveis, data_limite, status, data_conclusao, acoes_data):
    if not n_clicks:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update
    
    # Validação melhorada para campos obrigatórios
    campos_vazios = []
    
    if not projeto:
        campos_vazios.append("Projeto")
        
    if not mes_referencia:
        campos_vazios.append("Mês de Referência")
        
    if not prioridade:
        campos_vazios.append("Prioridade")
        
    if not descricao:
        campos_vazios.append("Descrição")
    
    # Verificação especial para responsáveis que pode ser lista vazia, None, ou string vazia
    responsaveis_vazio = (responsaveis is None or 
                         (isinstance(responsaveis, list) and (len(responsaveis) == 0 or all(not r.strip() for r in responsaveis if isinstance(r, str)))) or
                         (isinstance(responsaveis, str) and not responsaveis.strip()))
    
    if responsaveis_vazio:
        campos_vazios.append("Responsáveis")
    
    if not data_limite:
        campos_vazios.append("Data Limite")
        
    if not status:
        campos_vazios.append("Status")
    
    # Se algum campo estiver vazio, exibir mensagem de alerta
    if campos_vazios:
        mensagem_erro = f"Por favor, preencha os seguintes campos obrigatórios: {', '.join(campos_vazios)}"
        print(f"Campos vazios no cadastro da ação: {', '.join(campos_vazios)}")
        return dash.no_update, True, mensagem_erro, dash.no_update
    
    try:
        # Converter responsáveis para string
        if isinstance(responsaveis, list):
            responsaveis_str = ', '.join(responsaveis)
        else:
            responsaveis_str = str(responsaveis)
            
        # Inicializar acoes_data como lista vazia se for None
        if acoes_data is None:
            acoes_data = []
        elif not isinstance(acoes_data, list):
            # Tentar converter para lista se não for
            try:
                acoes_data = list(acoes_data)
            except:
                acoes_data = []
        
        # Gerar ID da ação (próximo número sequencial)
        next_id = 1
        if acoes_data and len(acoes_data) > 0:
            df_acoes = pd.DataFrame(acoes_data)
            if 'ID da Ação' in df_acoes.columns:
                next_id = df_acoes['ID da Ação'].max() + 1 if not pd.isna(df_acoes['ID da Ação'].max()) else 1
            
        # Preparar nova linha
        nova_acao = {
            'ID da Ação': next_id,
            'Data de Cadastro': datetime.now().strftime('%Y-%m-%d'),
            'Mês de Referência': mes_referencia,
            'Projeto': projeto,
            'Descrição da Ação': descricao,
            'Responsáveis': responsaveis_str,
            'Data Limite': data_limite,
            'Status': status,
            'Prioridade': prioridade,
            'Data de Conclusão': data_conclusao,
            'Observações de conclusão': ""
        }
        
        # Adicionar nova ação aos dados existentes
        acoes_data.append(nova_acao)
        
        # Atualizar dados na planilha do Google Sheets
        df_acoes = pd.DataFrame(acoes_data)
        success = update_acoes_in_sheets(df_acoes)
        if success:
            print(f"Ação cadastrada com sucesso: ID {next_id}")
        else:
            print(f"AVISO: Falha ao salvar ação na planilha, mas foi salva localmente")
            save_data_to_local(df_acoes, "acoes")
        
        return False, False, "", acoes_data
    
    except Exception as e:
        print(f"Erro ao salvar ação: {e}")
        import traceback
        traceback.print_exc()
        return dash.no_update, True, f"Erro ao salvar ação: {str(e)}", dash.no_update

# Callback adicional para garantir que a coluna Observacoes esteja presente nos dados da tabela
@app.callback(
    Output("projetos-table", "data", allow_duplicate=True),
    Input("projetos-table", "data"),
    prevent_initial_call=True
)
def ensure_observacoes_column(data):
    if not data:
        return dash.no_update
    
    # Converter para DataFrame para facilitar manipulação
    df = pd.DataFrame(data)
    
    # Verificar se a coluna Observacoes existe
    if 'Observacoes' not in df.columns:
        print("AVISO: Coluna 'Observacoes' não encontrada na tabela, adicionando coluna vazia.")
        df['Observacoes'] = ""
    else:
        # Garantir que valores nulos sejam tratados como strings vazias
        df['Observacoes'] = df['Observacoes'].fillna("")
        print(f"INFO: A tabela tem {len(df)} linhas e a coluna 'Observacoes' tem {(df['Observacoes'] != '').sum()} valores não vazios.")
    
    return df.to_dict('records')

# Bloco principal para executar o aplicativo
if __name__ == '__main__':
    try:
        print("\n===== Iniciando aplicativo Status Mensal Codeart =====")
        print("Verificando conexão com Google Sheets...")
        
        # Testar a conexão com o Google Sheets
        spreadsheet = connect_google_sheets()
        if spreadsheet:
            print(f"✅ Conexão com Google Sheets estabelecida com sucesso!")
            print(f"   Planilha: {spreadsheet.title}")
            print(f"   Abas disponíveis:")
            for sheet in spreadsheet.worksheets():
                print(f"   - {sheet.title}")
        else:
            print("❌ ERRO: Não foi possível conectar ao Google Sheets.")
            print("   O aplicativo funcionará com dados limitados ou vazios.")
            print("   Verifique o arquivo de credenciais e as permissões da conta de serviço.")
        
        # Em produção (como no Render), o app será executado pelo Gunicorn
        # Em desenvolvimento local, usamos o servidor integrado do Dash
        import os
        debug = os.environ.get('ENV', 'development') == 'development'
        port = int(os.environ.get('PORT', 8050))
        
        print(f"\nIniciando servidor Dash {'em modo debug' if debug else 'em produção'} na porta {port}")
        app.run(debug=debug, host='0.0.0.0', port=port)
    except Exception as e:
        print(f"❌ ERRO CRÍTICO ao iniciar o aplicativo: {e}")
        import traceback
        traceback.print_exc()

@app.callback(
    Output("modal-nova-acao", "is_open"),
    Input("nova-acao-btn", "n_clicks"),
    prevent_initial_call=True
)
def open_nova_acao_modal(n_clicks):
    if n_clicks:
        return True
    return False