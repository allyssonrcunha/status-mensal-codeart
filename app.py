import plotly.io as pio
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

# Nome do arquivo que deve estar na mesma pasta do script
EXCEL_FILE_PATH = 'Revisão Projetos - Geral.xlsx'

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
                date_cols = ['Data de Cadastro',
                             'Data Limite', 'Data de Conclusão']
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
                wait_time = initial_delay * \
                    (2 ** retries) + random.uniform(0, 1)
                print(
                    f"Quota excedida. Aguardando {wait_time:.2f} segundos antes de tentar novamente...")
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
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']

        # Caminho para o arquivo de credenciais
        credentials_path = 'google_credentials.json'

        # Verificar se o arquivo de credenciais existe
        if not os.path.exists(credentials_path):
            print(
                f"ERRO: Arquivo de credenciais não encontrado em {credentials_path}")
            print("Por favor, verifique se você colocou o arquivo google-credentials.json no diretório 'credentials'")
            return None

        # Carregar credenciais do arquivo JSON
        credentials = ServiceAccountCredentials.from_json_keyfile_name(
            credentials_path, scope)

        # Autorizar o cliente gspread com as credenciais
        client = gspread.authorize(credentials)

        # Abrir a planilha pelo nome
        try:
            spreadsheet = client.open('Revisão Projetos - Geral')
            return spreadsheet
        except gspread.exceptions.SpreadsheetNotFound:
            print(
                "ERRO: Planilha 'Revisão Projetos - Geral' não encontrada no Google Drive")
            print(
                "Verifique se o nome da planilha está correto e se a conta de serviço tem acesso a ela")
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
            print(
                f"Usando dados em cache de Projetos (cache de {elapsed_time:.1f} segundos)")
            return CACHE_PROJETOS

    try:
        print("Carregando dados da aba Projetos...")

        def fetch_data():
            spreadsheet = connect_google_sheets()
            if not spreadsheet:
                print(
                    "Aviso: Usando DataFrame vazio para Projetos devido a falha na conexão com Google Sheets.")
                return pd.DataFrame()

            # Carregar aba Projetos
            sheet = spreadsheet.worksheet('Projetos')
            data = sheet.get_all_records()

            if not data:
                print("Aviso: Planilha Projetos está vazia.")
                return pd.DataFrame()

            df_projetos = pd.DataFrame(data)
            print(
                f"✅ Dados carregados com sucesso: {len(df_projetos)} projetos encontrados.")
            print(
                f"Colunas originais na planilha: {df_projetos.columns.tolist()}")
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
            print(
                f"Usando dados em cache de Codenautas (cache de {elapsed_time:.1f} segundos)")
            return CACHE_CODENAUTAS

    try:
        print("Carregando dados da aba Codenautas...")

        def fetch_data():
            spreadsheet = connect_google_sheets()
            if not spreadsheet:
                print(
                    "Aviso: Usando DataFrame vazio para Codenautas devido a falha na conexão com Google Sheets.")
                return pd.DataFrame()

            # Carregar aba Codenautas
            sheet = spreadsheet.worksheet('Codenautas')
            data = sheet.get_all_records()

            if not data:
                print("Aviso: Planilha Codenautas está vazia.")
                return pd.DataFrame()

            df_codenautas = pd.DataFrame(data)
            print(
                f"✅ Dados carregados com sucesso: {len(df_codenautas)} codenautas encontrados.")
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

    # Forçar atualização a partir do Google Sheets (ignorar cache)
    force_update = True
    
    # Verificar se existe cache válido (só usar se não for forçar atualização)
    if not force_update and CACHE_ACOES is not None and LAST_CACHE_UPDATE is not None:
        elapsed_time = time.time() - LAST_CACHE_UPDATE
        if elapsed_time < CACHE_DURATION:
            print(
                f"Usando dados em cache de Ações (cache de {elapsed_time:.1f} segundos)")
            return CACHE_ACOES

    try:
        print("Carregando dados da aba Ações diretamente do Google Sheets...")

        def fetch_data():
            spreadsheet = connect_google_sheets()
            if not spreadsheet:
                print(
                    "Aviso: Não foi possível conectar ao Google Sheets para obter ações.")
                return pd.DataFrame()

            # Carregar aba Ações
            sheet = spreadsheet.worksheet('Ações')
            data = sheet.get_all_records()

            if not data:
                print("Aviso: Planilha Ações está vazia no Google Sheets.")
                return pd.DataFrame()

            df_acoes = pd.DataFrame(data)
            
            # Verificar e imprimir as colunas para debug
            print(f"Colunas encontradas na guia Ações: {df_acoes.columns.tolist()}")

            # Converter colunas de data para datetime
            date_cols = ['Data de Cadastro',
                         'Data Limite', 'Data de Conclusão']
            for col in date_cols:
                if col in df_acoes.columns:
                    df_acoes[col] = pd.to_datetime(
                        df_acoes[col], errors='coerce')

            print(
                f"✅ Dados carregados com sucesso do Google Sheets: {len(df_acoes)} ações encontradas.")
            return df_acoes

        # Usar retry com backoff exponencial
        df_acoes = retry_with_backoff(fetch_data)

        # Se o dataframe ainda estiver vazio após tentar carregar do Google Sheets, 
        # só então tentar carregar do backup local
        if df_acoes.empty:
            print("Nenhuma ação encontrada no Google Sheets, tentando carregar do backup local...")
            df_local = load_data_from_local("acoes")
            if df_local is not None and not df_local.empty:
                df_acoes = df_local
                print(f"Carregadas {len(df_acoes)} ações do backup local.")

        # Atualizar cache apenas se tivermos dados
        if not df_acoes.empty:
            CACHE_ACOES = df_acoes
            LAST_CACHE_UPDATE = time.time()

            # Salvar cópia local para backup
            save_data_to_local(df_acoes, "acoes")
            print(f"Backup local de ações atualizado com {len(df_acoes)} registros.")

        return df_acoes

    except Exception as e:
        print(f"Erro ao carregar dados de Ações do Google Sheets: {e}")
        import traceback
        traceback.print_exc()
        
        # Tentar usar cache existente
        if CACHE_ACOES is not None:
            print("Usando dados em cache de Ações devido a erro na atualização")
            return CACHE_ACOES

        # Tentar carregar do backup local como última opção
        print("Tentando carregar ações do backup local após erro...")
        df_local = load_data_from_local("acoes")
        if df_local is not None:
            return df_local

        return pd.DataFrame()

# Função para atualizar dados das Ações


def update_acoes_in_sheets(df_acoes):
    try:
        print("\n===== Atualizando ações no Google Sheets =====")
        print(f"Tentando atualizar {len(df_acoes)} registros de ações")
        
        spreadsheet = connect_google_sheets()
        if not spreadsheet:
            print("❌ Não foi possível conectar ao Google Sheets")
            return False

        # Preparar dados para upload - SALVAR UMA CÓPIA ANTES DAS TRANSFORMAÇÕES
        df_original = df_acoes.copy()
        
        # Converter datas para string no formato YYYY-MM-DD
        df_to_upload = df_acoes.copy()
        date_cols = ['Data de Cadastro', 'Data Limite', 'Data de Conclusão']
        
        # NOVA ABORDAGEM PARA DATAS: Tratamento mais cuidadoso
        print("\n=== APLICANDO NOVO TRATAMENTO PARA DATAS ===")
        for col in date_cols:
            if col in df_to_upload.columns:
                print(f"Processando coluna {col}...")
                
                # Verificação detalhada para Data Limite
                if col == 'Data Limite':
                    print(f"Valores em Data Limite antes da conversão: {df_to_upload[col].head(3).tolist()}")
                    # Verificar se há valores None que deveriam ser preservados
                    for idx, valor in enumerate(df_to_upload[col]):
                        if pd.isna(valor) and col == 'Data Limite':
                            print(f"ATENÇÃO: Linha {idx} tem Data Limite nula/vazia")
                
                # Converter somente valores não-nulos para strings formatadas
                # Importante: Preservar NaN/None para Data Limite e outros campos
                df_to_upload[col] = df_to_upload[col].apply(
                    lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) and hasattr(x, 'strftime') else (
                        "" if pd.isna(x) else x
                    )
                )
                
                # Garantir que todos os valores na coluna sejam strings
                df_to_upload[col] = df_to_upload[col].astype(str)
                
                # Substituir valores "nan", "None", "NaT" por string vazia
                df_to_upload[col] = df_to_upload[col].replace(["nan", "None", "NaT"], "")
                
                # Verificar tipos de dados após processamento
                print(f"Tipo de dados na coluna {col} após conversão: {df_to_upload[col].dtype}")
                print(f"Valores não vazios: {(df_to_upload[col] != '').sum()} de {len(df_to_upload)}")
        
        # Remover colunas calculadas que não devem ser enviadas para o Google Sheets
        colunas_calculadas = ['Dias Restantes', 'Atrasada', 'Tempo de Conclusão']
        for col in colunas_calculadas:
            if col in df_to_upload.columns:
                df_to_upload = df_to_upload.drop(columns=[col])
                print(f"Coluna calculada {col} removida antes do upload")

        # Verificação final da Data Limite
        if 'Data Limite' in df_to_upload.columns:
            print("\n=== VERIFICAÇÃO FINAL DA DATA LIMITE ===")
            print(f"Valores em Data Limite antes do upload:")
            print(f"  - Tipo de dados: {df_to_upload['Data Limite'].dtype}")
            print(f"  - Valores vazios: {(df_to_upload['Data Limite'] == '').sum()}")
            print(f"  - Primeiros 5 valores: {df_to_upload['Data Limite'].head(5).tolist()}")
        
        # Salvar backup local antes da atualização no Google Sheets
        save_data_to_local(df_original, "acoes_antes_upload_original")
        save_data_to_local(df_to_upload, "acoes_antes_upload_processado")
        
        # Verificar se há dados para enviar
        if df_to_upload.empty:
            print("⚠️ Não há dados para atualizar na planilha")
            return False

        # Converter DataFrame para lista de listas
        headers = df_to_upload.columns.tolist()
        rows = []
        
        for _, row in df_to_upload.iterrows():
            row_values = []
            for col in headers:
                value = row[col]
                # Garantir que não temos valores None, nan ou NaT
                if pd.isna(value) or value is None:
                    value = ""
                # Não fazer substituição automática das datas limite vazias
                row_values.append(value)
            rows.append(row_values)
            
        values = [headers] + rows
        
        print(f"Preparados {len(values)-1} registros para upload")
        print(f"Colunas para upload: {headers}")
        
        # Verificar Data Limite em valores
        data_limite_idx = headers.index('Data Limite') if 'Data Limite' in headers else -1
        if data_limite_idx >= 0:
            print(f"Verificando valores de Data Limite na lista de valores:")
            for i, row in enumerate(rows[:3]):  # Mostrar 3 primeiros exemplos
                print(f"  - Registro {i+1}: Data Limite = '{row[data_limite_idx]}'")

        # Atualizar planilha
        try:
            acoes_sheet = spreadsheet.worksheet('Ações')
            print("✅ Guia 'Ações' encontrada no Google Sheets")
            
            # Obter dados existentes antes de limpar
            existing_data = acoes_sheet.get_all_values()
            print(f"A guia tem atualmente {len(existing_data)} linhas, incluindo cabeçalho")
            
            # Usando abordagem simplificada para atualização da planilha
            print("Utilizando abordagem de atualização em duas etapas...")
            
            # 1. Atualizar apenas o cabeçalho primeiro
            acoes_sheet.update('A1', [headers])
            print("Cabeçalho atualizado")
            
            # 2. Se houver linhas de dados, atualizar os dados a partir da linha 2
            if len(rows) > 0:
                acoes_sheet.update('A2', rows)
                print(f"Dados atualizados ({len(rows)} linhas)")
            
            # Atualizar cache
            global CACHE_ACOES, LAST_CACHE_UPDATE
            CACHE_ACOES = df_original  # Usar o DataFrame original para o cache
            LAST_CACHE_UPDATE = time.time()
            
            # Salvar backup local após sucesso
            save_data_to_local(df_original, "acoes")
            print("✅ Dados de ações atualizados com sucesso no Google Sheets e no cache")
            
            return True
            
        except Exception as sheet_e:
            print(f"❌ Erro ao acessar ou atualizar a guia 'Ações': {sheet_e}")
            import traceback
            traceback.print_exc()
            return False

    except Exception as e:
        print(f"❌ Erro ao atualizar planilha de Ações: {e}")
        import traceback
        traceback.print_exc()
        return False


# Configuração global do tema dos gráficos Plotly

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
        colorway=['#6CC0ED', '#FED600', '#416072',
                  '#303E47', '#28a745', '#dc3545'],
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
    'dark_blue': '#416072',  # Azul escuro
    'dark_gray': '#303E47',  # Cinza escuro
    'white': '#FFFFFF',     # Branco
    'success': '#28a745',   # Verde para sucesso
    'danger': '#dc3545',    # Vermelho para alerta
    'charcoal_blue': '#172B36',  # Azul muito escuro
    'cloud': '#F7F9FA',     # Cinza muito claro
    'deep_sea': '#3A84A7',  # Azul médio
    'background': '#F4F7FA',  # Fundo cinza claro
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
        return df_projetos  # Retorna DataFrame vazio se não houver dados

    # Imprimir informações para debug
    print(
        f"Processando dados da planilha: {len(df_projetos)} linhas, {len(df_projetos.columns)} colunas")
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
        print(
            f"Resolvendo conflito: '{original}' e '{expected}' existem simultaneamente.")
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
                    df_renamed = df_renamed.rename(
                        columns={'Observações': 'Observacoes'})
                else:
                    df_renamed[col] = ""
            elif col == 'NPS ':
                if 'NPS' in df_renamed.columns:
                    df_renamed = df_renamed.rename(columns={'NPS': 'NPS '})
                else:
                    df_renamed[col] = ""
            elif col == 'Atraso em dias ':
                if 'Atraso em dias' in df_renamed.columns:
                    df_renamed = df_renamed.rename(
                        columns={'Atraso em dias': 'Atraso em dias '})
                else:
                    df_renamed[col] = 0
            elif col == 'Previsão':
                # Verificar diversas opções possíveis para a coluna de horas previstas
                for possible_col in ['Horas Previstas (Contrato)', 'Previsão', 'Previsto']:
                    if possible_col in df_renamed.columns:
                        df_renamed = df_renamed.rename(
                            columns={possible_col: 'Previsão'})
                        break
                else:
                    df_renamed[col] = 0
            else:
                df_renamed[col] = "Não Informado" if col in [
                    'GP Responsável', 'Status', 'Segmento', 'Tipo', 'Coordenação', 'Financeiro'] else 0

    # 1. Formatar coluna 'Mês' para exibição e filtro (ex: Abr/2025)
    try:
        df_renamed['Mês_datetime'] = pd.to_datetime(
            df_renamed['Mês'], errors='coerce')
        # CORREÇÃO: O ponto (.) antes de 'Out' foi substituído por dois pontos (:)
        month_map_pt = {1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun',
                        7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'}
        df_renamed['MesAnoFormatado'] = df_renamed['Mês_datetime'].apply(
            lambda x: f"{month_map_pt[x.month]}/{x.year}" if pd.notnull(x) and hasattr(x, 'month') and hasattr(x, 'year') else str(
                df_renamed.loc[x.name if hasattr(x, 'name') else -1, 'Mês'] if hasattr(x, 'name') else 'Data Inválida')
        )
        # Fallback for any original strings that couldn't be parsed but should be kept
        df_renamed['MesAnoFormatado'] = np.where(df_renamed['Mês_datetime'].isna(
        ), df_renamed['Mês'].astype(str), df_renamed['MesAnoFormatado'])

        # Adicionar coluna de ano-mês para agrupamento nos gráficos de evolução
        df_renamed['Ano_Mes'] = df_renamed['Mês_datetime'].dt.strftime('%Y-%m')
    except Exception as e:
        print(f"Erro ao formatar coluna 'Mês': {e}. Usando como string.")
        df_renamed['MesAnoFormatado'] = df_renamed['Mês'].astype(str)
        df_renamed['Ano_Mes'] = df_renamed['Mês'].astype(str)

    # 2. Converter outras colunas de filtro para string para consistência
    for col in ['GP Responsável', 'Status', 'Segmento', 'Tipo', 'Coordenação', 'Financeiro']:
        if col in df_renamed.columns:
            df_renamed[col] = df_renamed[col].astype(str).fillna(
                'Não Informado')  # Tratar NaNs nos filtros
        else:
            # Adicionar coluna se não existir
            df_renamed[col] = 'Não Informado'

    # 3. Identificar projetos críticos baseado na coluna "Decisões"
    if 'Decisões' in df_renamed.columns:
        df_renamed['Prioridade'] = np.where(df_renamed['Decisões'].astype(
            str).str.contains('Crítico', na=False), 'Crítico', 'Normal')
    else:
        # Se a coluna não existir, nenhum é crítico
        df_renamed['Prioridade'] = 'Normal'

    # 4. Combinar NPS e emoji em uma única coluna
    def nps_com_emoji(nps_value):
        # Tratar NaN, string vazia e 'nan'
        if pd.isna(nps_value) or str(nps_value).strip() == '' or str(nps_value).lower() == 'nan':
            return ""
        elif nps_value == "Promotor":
            return "Promotor 😀"  # Emoji feliz verde
        elif nps_value == "Neutro":
            return "Neutro 😐"  # Emoji neutro amarelo
        elif nps_value == "Detrator":
            return "Detrator 😡"  # Emoji triste vermelho
        else:
            # Retorna o valor original se for diferente dos esperados
            return str(nps_value)

    if 'NPS ' in df_renamed.columns:
        df_renamed['NPS_Combinado'] = df_renamed['NPS '].apply(nps_com_emoji)
    else:
        # Se a coluna não existir, criar vazia
        df_renamed['NPS_Combinado'] = ""

    # 5. Garantir que colunas para cálculo sejam numéricas e preencher NaNs
    cols_to_convert_to_numeric = ['Atraso em dias ', 'Previsão', 'Real',
                                  'Saldo Acumulado', 'Horas Previstas (Contrato)', 'Horas Mês']
    for col in cols_to_convert_to_numeric:
        try:
            if col in df_renamed.columns:
                # Converter para numérico, tratando erros como NaN
                df_renamed[col] = pd.to_numeric(
                    df_renamed[col], errors='coerce')

                # Corrigir valores com formatação incorreta
                if col in ['Previsão', 'Real', 'Saldo Acumulado', 'Horas Previstas (Contrato)', 'Horas Mês']:
                    # Verificar se há valores extremamente altos
                    # Usar mediana em vez de média para ser menos afetado por outliers
                    valor_medio = df_renamed[col].median()
                    # Valores 10x acima da mediana
                    valores_suspeitosos = df_renamed[col] > valor_medio * 10

                    # Corrigir apenas se houver poucos valores suspeitos (evitar corrigir dados válidos)
                    if valores_suspeitosos.sum() > 0 and valores_suspeitosos.sum() < len(df_renamed) * 0.2:
                        print(
                            f"Encontrados {valores_suspeitosos.sum()} valores suspeitosamente altos na coluna '{col}'. Mediana: {valor_medio:.2f}")

                        # Verificar quais projetos estão afetados
                        projetos_afetados = df_renamed.loc[valores_suspeitosos, 'Projeto'].tolist(
                        )
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

                        print(
                            f"Aplicando fator de correção de {fator_correcao} para valores altos na coluna {col}")

                        # Aplicar a correção
                        df_renamed.loc[valores_suspeitosos,
                                       col] = df_renamed.loc[valores_suspeitosos, col] / fator_correcao

                        # Verificar novos valores
                        print(
                            f"Após correção '{col}': Min={df_renamed[col].min():.2f}, Max={df_renamed[col].max():.2f}, Média={df_renamed[col].mean():.2f}")

                # Preencher valores nulos com zero
                df_renamed[col] = df_renamed[col].fillna(0)
            else:
                print(
                    f"AVISO: Coluna '{col}' não encontrada, criando com zeros.")
                df_renamed[col] = 0
        except Exception as e:
            print(f"ERRO ao processar coluna '{col}': {e}")
            # Tentar recuperar a coluna em caso de erro
            df_renamed[col] = 0

    # 6. Garantir que a coluna Observacoes exista e tratar NaNs
    if 'Observacoes' not in df_renamed.columns:
        if 'Observações' in df_renamed.columns:
            print(f"INFO: Renomeando coluna 'Observações' para 'Observacoes'")
            df_renamed = df_renamed.rename(
                columns={'Observações': 'Observacoes'})
        else:
            print(
                f"AVISO: Coluna 'Observações' não encontrada, criando 'Observacoes' vazia")
            df_renamed['Observacoes'] = ""
    else:
        print(f"INFO: Coluna 'Observacoes' já existe no DataFrame")

    # Debug: Verificar valores da coluna Observacoes
    if 'Observacoes' in df_renamed.columns:
        df_renamed['Observacoes'] = df_renamed['Observacoes'].fillna(
            "")  # Preencher NaNs com string vazia
        non_empty = (df_renamed['Observacoes'] != '').sum()
        print(
            f"INFO: A coluna 'Observacoes' tem {non_empty} valores não vazios de {len(df_renamed)} registros")

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

        df_renamed['Cliente'] = df_renamed['Projeto'].apply(
            extract_client_name)

    # Print para debug
    print(
        f"Colunas disponíveis após processamento: {df_renamed.columns.tolist()}")
    print(f"Linhas após processamento: {len(df_renamed)}")

    # Verificar métricas básicas
    try:
        print("\nMétricas básicas após processamento:")
        for col in ['Previsão', 'Real', 'Saldo Acumulado']:
            if col in df_renamed.columns:
                print(
                    f"  {col}: Min={df_renamed[col].min()}, Max={df_renamed[col].max()}, Média={df_renamed[col].mean():.2f}")
    except Exception as e:
        print(f"Erro ao calcular métricas básicas: {e}")

    return df_renamed

# Função para processar dados das ações


def process_acoes(df_acoes):
    if df_acoes.empty:
        return df_acoes

    # Debug: Mostra as colunas na entrada
    print(f"Processando ações com colunas: {df_acoes.columns.tolist()}")

    # Garantir que Status e Prioridade tenham valores padrão
    df_acoes['Status'] = df_acoes['Status'].fillna('Pendente')
    df_acoes['Prioridade'] = df_acoes['Prioridade'].fillna('Média')

    # Converter colunas de data para datetime
    date_cols = ['Data de Cadastro', 'Data Limite', 'Data de Conclusão']
    for col in date_cols:
        if col in df_acoes.columns:
            print(f"Convertendo coluna {col} para datetime")
            df_acoes[col] = pd.to_datetime(df_acoes[col], errors='coerce')

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
                    df_acoes.at[idx, 'Dias Restantes'] = (
                        data_limite - hoje).days
            except Exception as e:
                print(
                    f"Erro ao calcular dias restantes para índice {idx}: {e}")
                df_acoes.at[idx, 'Dias Restantes'] = pd.NA

    # Marcar ações atrasadas (com status pendente e data limite passada)
    df_acoes['Atrasada'] = (
        (df_acoes['Dias Restantes'].notna()) &
        (df_acoes['Dias Restantes'] < 0) &
        (df_acoes['Status'] != 'Concluída')
    ).astype(int)

    # Calcular tempo de conclusão para ações concluídas
    df_acoes['Tempo de Conclusão'] = pd.NA
    mask_concluida = (df_acoes['Status'] == 'Concluída') & ~df_acoes['Data de Conclusão'].isna(
    ) & ~df_acoes['Data de Cadastro'].isna()

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
        mask_concluida = (df_acoes['Status'] == 'Concluída') & ~df_acoes['Data de Conclusão'].isna(
        ) & ~df_acoes['Data de Cadastro'].isna()

        # Calcular tempo de conclusão linha por linha
        for idx in df_acoes[mask_concluida].index:
            try:
                data_conclusao = pd.to_datetime(
                    df_acoes.at[idx, 'Data de Conclusão'])
                data_cadastro = pd.to_datetime(
                    df_acoes.at[idx, 'Data de Cadastro'])
                if pd.notna(data_conclusao) and pd.notna(data_cadastro):
                    df_acoes.at[idx, 'Tempo de Conclusão'] = (
                        data_conclusao - data_cadastro).days
            except Exception as e:
                print(
                    f"Erro ao calcular tempo de conclusão para índice {idx}: {e}")
                df_acoes.at[idx, 'Tempo de Conclusão'] = pd.NA

    # Debug: Mostra as colunas na saída
    print(f"Ações processadas com colunas: {df_acoes.columns.tolist()}")
    print(f"Total de {len(df_acoes)} ações processadas")

    return df_acoes


# Inicializar o aplicativo Dash com tema Bootstrap e definir o título da página
app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP,
                          'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.1/css/all.min.css'],
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
print("Dados carregados com sucesso!")

# Obter listas para os filtros iniciais


def get_filter_options(df):
    if df.empty:
        # Adicionado um [] extra para o status financeiro
        return [], [], [], [], [], [], []
    meses_anos = sorted(df['MesAnoFormatado'].unique())
    gestoras = sorted(df['GP Responsável'].unique())
    status_list = sorted(df['Status'].unique())
    segmentos = sorted(df['Segmento'].unique())
    tipos = sorted(df['Tipo'].unique())
    coordenacoes = sorted(df['Coordenação'].unique())
    # Adicionar status financeiro
    financeiro_list = sorted(df['Financeiro'].astype(str).unique())

    return meses_anos, gestoras, status_list, segmentos, tipos, coordenacoes, financeiro_list


meses_anos_initial, gestoras_initial, status_list_initial, segmentos_initial, tipos_initial, coordenacoes_initial, financeiro_list_initial = get_filter_options(
    df_projetos_initial)

# Layout simples para teste
app.layout = html.Div(style=custom_style['body'], children=[
    # Cabeçalho com logo, título e botão de atualização
    html.Div(style=custom_style['header'], children=[
        html.Div([  # Container para logo e título
            html.Img(src=logo_src,
                     style=custom_style['logo']) if logo_src else None,
            html.H1("Status Mensal Codeart", style=custom_style['title'])
        ], style={'display': 'flex', 'align-items': 'center'}),
        html.Div([  # Container para botão e hora
            dbc.Button(
                [html.I(className="fas fa-sync-alt me-2"), " Atualizar Dados"],
                id="refresh-data-button",
                color="primary",
                className="me-2",
                style={
                    'backgroundColor': codeart_colors['blue_sky'], 'borderColor': codeart_colors['blue_sky']}
            ),
            html.Span(id="last-update-time",
                      style=custom_style['last_update_style'])
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
                    dbc.Col([html.Div([html.Div(id="total-projetos", className="metric-value"), html.Div(
                        "Total de Projetos", className="metric-label")], style=custom_style['metric-card'])], width=3),
                    dbc.Col([html.Div([html.Div(id="total-clientes", className="metric-value"), html.Div(
                        "Total de Clientes", className="metric-label")], style=custom_style['metric-card'])], width=3),
                    dbc.Col([html.Div([html.Div(id="projetos-atrasados", className="metric-value"), html.Div(
                        "Projetos Atrasados", className="metric-label")], style=custom_style['metric-card'])], width=3),
                    dbc.Col([html.Div([html.Div(id="projetos-criticos", className="metric-value"), html.Div(
                        "Projetos Críticos", className="metric-label")], style=custom_style['metric-card'])], width=3),
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
                                    options=[{"label": mes, "value": mes}
                                             for mes in meses_anos_initial],
                                    multi=True,
                                    placeholder="Selecione o mês/ano"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Gestora"),
                                dcc.Dropdown(
                                    id="gestora-filter",
                                    options=[{"label": gp, "value": gp}
                                             for gp in gestoras_initial],
                                    multi=True,
                                    placeholder="Selecione a gestora"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Status"),
                                dcc.Dropdown(
                                    id="status-filter",
                                    options=[{"label": status, "value": status}
                                             for status in status_list_initial],
                                    multi=True,
                                    placeholder="Selecione o status"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Coordenação"),
                                dcc.Dropdown(
                                    id="coordenacao-filter",
                                    options=[{"label": coord, "value": coord}
                                             for coord in coordenacoes_initial],
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
                                    options=[{"label": seg, "value": seg}
                                             for seg in segmentos_initial],
                                    multi=True,
                                    placeholder="Selecione o segmento"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Tipo"),
                                dcc.Dropdown(
                                    id="tipo-filter",
                                    options=[{"label": tipo, "value": tipo}
                                             for tipo in tipos_initial],
                                    multi=True,
                                    placeholder="Selecione o tipo"
                                )
                            ], width=3),
                            dbc.Col([
                                html.Label("Financeiro"),
                                dcc.Dropdown(
                                    id="financeiro-filter",
                                    options=[{"label": fin, "value": fin}
                                             for fin in financeiro_list_initial],
                                    multi=True,
                                    placeholder="Selecione o status financeiro"
                                )
                            ], width=3),
                            dbc.Col(width=3),
                        ]),
                        dbc.Row([
                            dbc.Col([
                                html.Div([
                                    dbc.Button("Aplicar Filtros", id="apply-project-filters", color="primary",
                                               className="me-2", style={'backgroundColor': codeart_colors['dark_blue']}),
                                    dbc.Button(
                                        "Limpar Filtros", id="reset-project-filters", color="secondary")
                                ], className="mt-3 mb-4")
                            ], width=12),
                        ])
                    ])
                ]),

                # Gráficos
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="status-chart")],
                            style=custom_style['chart-container'])], width=6),
                    dbc.Col([html.Div([dcc.Graph(id="financeiro-chart")],
                            style=custom_style['chart-container'])], width=6),
                ]),

                # Gráfico de NPS e gráfico de Projetos por Gestora (em tela cheia)
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="nps-chart")],
                            style=custom_style['chart-container'])], width=6),
                    dbc.Col([html.Div([dcc.Graph(id="segmento-chart")],
                            style=custom_style['chart-container'])], width=6),
                ]),

                # Gráfico de Projetos por Gestora
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="projetos-gp-chart")],
                            style=custom_style['chart-container'])], width=12),
                ]),

                # Gráficos adicionais
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="horas-chart")],
                            style=custom_style['chart-container'])], width=12),
                ]),

                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="saldo-chart")],
                            style=custom_style['chart-container'])], width=12),
                ]),

                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="atraso-coordenacao-chart")],
                            style=custom_style['chart-container'])], width=12),
                ]),

                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="evolucao-quitados-chart")],
                            style=custom_style['chart-container'])], width=12),
                ]),

                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="evolucao-atrasados-chart")],
                            style=custom_style['chart-container'])], width=12),
                ]),

                # Tabela de dados
                dbc.Row([
                    dbc.Col([
                        html.Div([
                            html.H3("Lista de Projetos",
                                    className="mt-4 mb-3 d-inline-block"),
                            html.Div([
                                dbc.Button(
                                    [html.I(
                                        className="fas fa-file-export me-2"), " Exportar Dados"],
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

                # Adicionar campo de busca para a tabela de projetos
                dbc.Row([
                    dbc.Col([
                        html.Div([
                            html.Label("Buscar:", className="me-2"),
                            dbc.Input(
                                id="projetos-table-search",
                                type="text",
                                placeholder="Digite para buscar...",
                                className="mb-3",
                                style={"width": "100%"}
                            )
                        ])
                    ], width=12)
                ]),

                html.Div([dash_table.DataTable(
                    id="projetos-table",
                    columns=[
                        # Nova coluna para o ícone de ação
                        {"name": "", "id": "action_icon", "type": "text"},
                        {"name": "Mês", "id": "MesAnoFormatado"},
                        {"name": "Projeto", "id": "Projeto"},
                        {"name": "Cliente", "id": "Cliente"},
                        {"name": "Gestora", "id": "GP Responsável"},
                        {"name": "Coordenação", "id": "Coordenação"},
                        {"name": "Segmento", "id": "Segmento"},
                        {"name": "Tipo", "id": "Tipo"},
                        {"name": "Status", "id": "Status"},
                        {"name": "Horas Previstas", "id": "Previsão", "type": "numeric",
                            "format": Format(precision=1, scheme=Scheme.fixed)},
                        {"name": "Horas Realizadas", "id": "Real", "type": "numeric",
                            "format": Format(precision=1, scheme=Scheme.fixed)},
                        {"name": "Saldo", "id": "Saldo Acumulado", "type": "numeric",
                            "format": Format(precision=1, scheme=Scheme.fixed)},
                        {"name": "Atraso (dias)", "id": "Atraso em dias ", "type": "numeric", "format": Format(
                            precision=0, scheme=Scheme.fixed)},
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
                    css=[{"selector": ".dash-cell div.dash-cell-value",
                          "rule": "display: inline; white-space: inherit; overflow: inherit; text-overflow: inherit;"}],
                    cell_selectable=True,
                    row_selectable=False,
                    selected_cells=[]
                )], style={'overflowX': 'auto'}),
            ]),

            # Aba de Ações
            dbc.Tab(label="Ações", tab_id="tab-acoes", children=[
                # Métricas de ações
                dbc.Row([
                    dbc.Col([html.Div([html.Div(id="total-acoes", className="metric-value"), html.Div(
                        "Total de Ações", className="metric-label")], style=custom_style['metric-card'])], width=3),
                    dbc.Col([html.Div([html.Div(id="acoes-pendentes", className="metric-value"), html.Div(
                        "Ações Pendentes", className="metric-label")], style=custom_style['metric-card'])], width=3),
                    dbc.Col([html.Div([html.Div(id="acoes-concluidas", className="metric-value"), html.Div(
                        "Ações Concluídas", className="metric-label")], style=custom_style['metric-card'])], width=3),
                    dbc.Col([html.Div([html.Div(id="acoes-atrasadas", className="metric-value"), html.Div(
                        "Ações Atrasadas", className="metric-label")], style=custom_style['metric-card'])], width=3),
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
                                    options=[{"label": mes, "value": mes}
                                             for mes in meses_anos_initial],
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
                                        {"label": "Em Andamento",
                                            "value": "Em Andamento"},
                                        {"label": "Concluída",
                                            "value": "Concluída"}
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
                                    dbc.Button("Aplicar Filtros", id="apply-acoes-filters", color="primary",
                                               className="me-2", style={'backgroundColor': codeart_colors['dark_blue']}),
                                    dbc.Button(
                                        "Limpar Filtros", id="reset-acoes-filters", color="secondary", className="me-2"),
                                    dbc.Button([html.I(className="fas fa-plus me-2"), " Nova Ação"], id="nova-acao-btn",
                                               color="success", style={'backgroundColor': codeart_colors['success']})
                                ], className="mt-3 mb-4")
                            ], width=12)
                        ])
                    ])
                ]),

                # Gráficos para ações
                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="status-acoes-chart")],
                            style=custom_style['chart-container'])], width=6),
                    dbc.Col([html.Div([dcc.Graph(id="prioridade-acoes-chart")],
                            style=custom_style['chart-container'])], width=6),
                ]),

                dbc.Row([
                    dbc.Col([html.Div([dcc.Graph(id="responsaveis-acoes-chart")],
                            style=custom_style['chart-container'])], width=6),
                    dbc.Col([html.Div([dcc.Graph(id="evolucao-acoes-chart")],
                            style=custom_style['chart-container'])], width=6),
                ]),

                # Tabela de ações
                dbc.Row([
                    dbc.Col([
                        html.Div([
                            html.H3("Lista de Ações", className="mt-4 mb-3"),
                            html.Div([
                                dbc.Button(
                                    [html.I(
                                        className="fas fa-download me-2"), " Exportar"],
                                    id="export-acoes-button",
                                    color="success",
                                    className="mb-3",
                                    style={
                                        'backgroundColor': codeart_colors['success']}
                                ),
                                dash_table.DataTable(
                                    id="acoes-table",
                                    columns=[
                                        {"name": "ID", "id": "ID da Ação"},
                                        {"name": "Mês", "id": "Mês de Referência"},
                                        {"name": "Projeto", "id": "Projeto"},
                                        {"name": "Descrição",
                                            "id": "Descrição da Ação"},
                                        {"name": "Responsáveis",
                                            "id": "Responsáveis"},
                                        {"name": "Data Limite",
                                            "id": "Data Limite"},
                                        {"name": "Data Conclusão",
                                            "id": "Data de Conclusão"},
                                        {"name": "Status", "id": "Status"},
                                        {"name": "Prioridade", "id": "Prioridade"},
                                        {"name": "Observações",
                                            "id": "Observações de conclusão"}
                                    ],
                                    page_size=20,
                                    style_table={'overflowX': 'auto'},
                                    style_cell={
                                        'textAlign': 'left',
                                        'padding': '8px',
                                        'minWidth': '100px',
                                        'maxWidth': '300px',
                                        'whiteSpace': 'normal',
                                        'overflow': 'hidden',
                                        'textOverflow': 'ellipsis'
                                    },
                                    tooltip_data=[
                                        {
                                            column: {'value': str(value), 'type': 'markdown'}
                                            for column, value in row.items()
                                        } for row in [] # Será preenchido pelo callback
                                    ],
                                    tooltip_duration=None,
                                    style_header={
                                        'backgroundColor': codeart_colors['dark_gray'],
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    style_data_conditional=[
                                        {
                                            'if': {'row_index': 'odd'},
                                            'backgroundColor': '#f8f9fa'
                                        },
                                        {
                                            'if': {'column_id': 'Observações de conclusão'},
                                            'maxWidth': '200px'
                                        },
                                        {
                                            'if': {'column_id': 'Descrição da Ação'},
                                            'maxWidth': '200px'
                                        },
                                        {
                                            'if': {'filter_query': '{Status} = "Concluída"'},
                                            'backgroundColor': '#d4edda',
                                            'color': '#155724'
                                        },
                                        {
                                            'if': {'filter_query': '{Atrasada} = 1'},
                                            'backgroundColor': '#f8d7da',
                                            'color': '#721c24'
                                        }
                                    ],
                                    cell_selectable=True,
                                    row_selectable=False,
                                    style_data={
                                        'cursor': 'pointer'  # Cursor de mão para indicar que é clicável
                                    },
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
                            dcc.Dropdown(id="modal-responsaveis",
                                         options=[], multi=True)
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
                            dcc.DatePickerSingle(
                                id="modal-data-limite", date=None)
                        ], width=4),
                        dbc.Col([
                            html.Label("Status"),
                            dcc.Dropdown(
                                id="modal-status",
                                options=[
                                    {"label": "Pendente", "value": "Pendente"},
                                    {"label": "Em Andamento",
                                        "value": "Em Andamento"},
                                    {"label": "Concluída", "value": "Concluída"}
                                ],
                                value="Pendente"
                            )
                        ], width=4),
                        dbc.Col([
                            html.Label("Data de Conclusão"),
                            dcc.DatePickerSingle(
                                id="modal-data-conclusao", date=None)
                        ], width=4),
                    ]),
                    dbc.Alert("Preencha todos os campos obrigatórios",
                              id="modal-alert-text", color="danger", is_open=False)
                ]),
                dbc.ModalFooter([
                    dbc.Button("Cancelar", id="modal-cancel",
                               color="secondary", className="me-2"),
                    dbc.Button("Salvar", id="modal-save", color="primary",
                               style={'backgroundColor': codeart_colors['dark_blue']})
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
                            dcc.Dropdown(
                                id="modal-edit-mes-referencia", options=[])
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
                            dcc.Dropdown(id="modal-edit-responsaveis",
                                         options=[], multi=True)
                        ], width=12),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Data Limite (Não Editável)"),
                            html.Div(id="modal-edit-data-limite-display", className="form-control", style={"backgroundColor": "#f0f0f0", "padding": "8px", "minHeight": "36px"})
                        ], width=4),
                        dbc.Col([
                            html.Label("Status"),
                            dcc.Dropdown(
                                id="modal-edit-status",
                                options=[
                                    {"label": "Pendente", "value": "Pendente"},
                                    {"label": "Em Andamento",
                                        "value": "Em Andamento"},
                                    {"label": "Concluída", "value": "Concluída"}
                                ]
                            )
                        ], width=4),
                        dbc.Col([
                            html.Label("Data de Conclusão"),
                            dcc.DatePickerSingle(
                                id="modal-edit-data-conclusao", date=None)
                        ], width=4),
                    ]),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Observações de Conclusão"),
                            dbc.Textarea(id="modal-edit-observacoes", rows=2)
                        ], width=12),
                    ]),
                    dbc.Alert("Preencha todos os campos obrigatórios",
                              id="modal-edit-alert-text", color="danger", is_open=False)
                ]),
                dbc.ModalFooter([
                    dbc.Button("Cancelar", id="modal-edit-cancel",
                               color="secondary", className="me-2"),
                    dbc.Button("Salvar", id="modal-edit-save", color="primary",
                               style={'backgroundColor': codeart_colors['dark_blue']})
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
                            html.Label(
                                "Projeto", className="form-label fw-bold"),
                            dcc.Dropdown(id="modal-acao-projeto", options=[])
                        ], width=6),
                        dbc.Col([
                            html.Label("Mês de Referência",
                                       className="form-label fw-bold"),
                            dcc.Dropdown(
                                id="modal-acao-mes-referencia", options=[])
                        ], width=6),
                    ], className="mb-3"),
                    dbc.Row([
                        dbc.Col([
                            html.Label(
                                "Prioridade", className="form-label fw-bold"),
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
                            dcc.Dropdown(id="modal-acao-responsaveis",
                                         options=[], multi=True)
                        ], width=6),
                    ], className="mb-3"),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Descrição da Ação",
                                       className="form-label fw-bold"),
                            dbc.Textarea(id="modal-acao-descricao",
                                         rows=3, className="w-100")
                        ], width=12),
                    ], className="mb-3"),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Data Limite",
                                       className="form-label fw-bold"),
                            html.Div([
                                dcc.DatePickerSingle(
                                    id="modal-acao-data-limite",
                                    date=None,
                                    display_format='DD/MM/YYYY',
                                    className="w-100"
                                )
                            ], style={"width": "100%"})
                        ], width=4),
                        dbc.Col([
                            html.Label(
                                "Status", className="form-label fw-bold"),
                            dcc.Dropdown(
                                id="modal-acao-status",
                                options=[
                                    {"label": "Pendente", "value": "Pendente"},
                                    {"label": "Em Andamento",
                                        "value": "Em Andamento"},
                                    {"label": "Concluída", "value": "Concluída"}
                                ],
                                value="Pendente"
                            )
                        ], width=4),
                        dbc.Col([
                            html.Label("Data de Conclusão",
                                       className="form-label fw-bold"),
                            html.Div([
                                dcc.DatePickerSingle(
                                    id="modal-acao-data-conclusao",
                                    date=None,
                                    display_format='DD/MM/YYYY',
                                    className="w-100"
                                )
                            ], style={"width": "100%"})
                        ], width=4),
                    ], className="mb-3"),
                    dbc.Alert("Preencha todos os campos obrigatórios",
                              id="modal-acao-alert-text", color="danger", is_open=False)
                ]),
                dbc.ModalFooter([
                    dbc.Button("Cancelar", id="modal-acao-cancel",
                               color="secondary", className="me-2"),
                    dbc.Button("Salvar", id="modal-acao-save", color="primary",
                               style={'backgroundColor': codeart_colors['dark_blue']})
                ]),
            ],
            id="modal-nova-acao",
            size="lg",
            is_open=False,
            style={"maxWidth": "800px"}
        ),

        # Stores para dados
        dcc.Store(id="raw-data-store",
                  data=df_projetos_initial.to_dict('records')),
        dcc.Store(id="codenautas-store",
                  data=df_codenautas_initial.to_dict('records')),
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
        html.Div(id="meses_anos-invisible", style={"display": "none"})
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
        print("\n===== Atualizando todos os dados do Google Sheets =====")
        # Forçar atualização invalidando completamente o cache
        LAST_CACHE_UPDATE = 0
        CACHE_PROJETOS = None
        CACHE_CODENAUTAS = None
        CACHE_ACOES = None  # Garantir que o cache de ações seja limpo

        # Recarregar dados do Google Sheets
        df_projetos_refreshed = load_data_from_sheets()
        df_projetos_refreshed = process_data(df_projetos_refreshed)

        # Salvar cópia local
        if not df_projetos_refreshed.empty:
            save_data_to_local(df_projetos_refreshed, "projetos")
        else:
            # Tentar carregar do backup local apenas se não conseguir dados do Google Sheets
            print("Alerta: Não foi possível obter dados de projetos do Google Sheets")
            df_local = load_data_from_local("projetos")
            if df_local is not None:
                df_projetos_refreshed = process_data(df_local)
                print(f"Usando backup local com {len(df_projetos_refreshed)} projetos")

        # Recarregar dados dos codenautas
        df_codenautas_refreshed = load_codenautas_from_sheets()

        # Salvar cópia local
        if not df_codenautas_refreshed.empty:
            save_data_to_local(df_codenautas_refreshed, "codenautas")
        else:
            # Tentar carregar do backup local apenas se não conseguir dados do Google Sheets
            print("Alerta: Não foi possível obter dados de codenautas do Google Sheets")
            df_local = load_data_from_local("codenautas")
            if df_local is not None:
                df_codenautas_refreshed = df_local
                print(f"Usando backup local com {len(df_codenautas_refreshed)} codenautas")

        # Recarregar dados das ações diretamente do Google Sheets
        print("Carregando dados de ações diretamente do Google Sheets...")
        try:
            spreadsheet = connect_google_sheets()
            if spreadsheet:
                sheet = spreadsheet.worksheet('Ações')
                data = sheet.get_all_records()
                
                if data:
                    df_acoes_refreshed = pd.DataFrame(data)
                    print(f"✅ Carregados {len(df_acoes_refreshed)} registros da guia Ações do Google Sheets")
                    print(f"Colunas encontradas: {df_acoes_refreshed.columns.tolist()}")
                    
                    # Converter colunas de data para datetime
                    date_cols = ['Data de Cadastro', 'Data Limite', 'Data de Conclusão']
                    for col in date_cols:
                        if col in df_acoes_refreshed.columns:
                            df_acoes_refreshed[col] = pd.to_datetime(df_acoes_refreshed[col], errors='coerce')
                            print(f"Coluna {col} convertida para datetime")
                    
                    # Processar dados
                    df_acoes_refreshed = process_acoes(df_acoes_refreshed)
                    
                    # Atualizar cache e salvar backup
                    CACHE_ACOES = df_acoes_refreshed
                    LAST_CACHE_UPDATE = time.time()
                    save_data_to_local(df_acoes_refreshed, "acoes")
                    print(f"Cache e backup de ações atualizados com {len(df_acoes_refreshed)} registros")
                else:
                    print("A guia Ações está vazia no Google Sheets")
                    df_acoes_refreshed = pd.DataFrame()
            else:
                print("Não foi possível conectar ao Google Sheets para obter ações")
                df_acoes_refreshed = pd.DataFrame()
        except Exception as e:
            print(f"Erro ao carregar ações diretamente: {e}")
            import traceback
            traceback.print_exc()
            df_acoes_refreshed = pd.DataFrame()
            
        # Se não conseguiu dados do Google Sheets, tentar o backup local
        if df_acoes_refreshed.empty:
            df_local = load_data_from_local("acoes")
            if df_local is not None and not df_local.empty:
                df_acoes_refreshed = process_acoes(df_local)
                print(f"Usando backup local com {len(df_acoes_refreshed)} ações")
            else:
                # Último recurso: chamar a função original
                print("Tentando método alternativo para carregar ações...")
                df_acoes_refreshed = load_acoes_from_sheets()
                df_acoes_refreshed = process_acoes(df_acoes_refreshed)

        print("===== Atualização de dados concluída =====\n")
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

    # Definir figuras vazias para usar como fallback
    empty_fig = go.Figure().update_layout(title="Sem dados disponíveis")
    status_fig = empty_fig
    financeiro_fig = empty_fig
    nps_fig = empty_fig
    segmento_fig = empty_fig
    gp_fig = empty_fig
    horas_fig = empty_fig
    saldo_fig = empty_fig
    atraso_coord_fig = empty_fig
    evolucao_quitados_fig = empty_fig
    evolucao_atrasados_fig = empty_fig
    df_table = pd.DataFrame()

    # Se o DataFrame estiver vazio, retornar valores vazios
    if df.empty:
        return "0", "0", "0", "0", status_fig, df_table.to_dict('records'), financeiro_fig, nps_fig, segmento_fig, gp_fig, horas_fig, saldo_fig, atraso_coord_fig, evolucao_quitados_fig, evolucao_atrasados_fig

    # Criar cópia para exibição na tabela (mostra todos os registros)
    df_table = df.copy()
    
    # Identificar projetos únicos para os totalizadores e gráficos
    # Consideramos um projeto como único combinando o nome do projeto e cliente
    if 'Projeto' in df.columns and 'Cliente' in df.columns:
        df['projeto_cliente'] = df['Projeto'] + ' - ' + df['Cliente']
        # Pegamos a versão mais recente de cada projeto (assumindo que MesAnoFormatado está presente)
        if 'MesAnoFormatado' in df.columns:
            # Ordenar por data (o mais recente primeiro)
            df = df.sort_values('MesAnoFormatado', ascending=False)
            # Remove registros duplicados, mantendo apenas o primeiro (mais recente) de cada projeto
            df_unique = df.drop_duplicates(subset=['projeto_cliente'])
        else:
            # Se não houver data, apenas remover duplicatas
            df_unique = df.drop_duplicates(subset=['projeto_cliente'])
    else:
        # Se não tiver as colunas necessárias, usar o dataframe original
        df_unique = df.copy()

    # Calcular métricas com projetos únicos
    total_projetos = len(df_unique)
    total_clientes = len(df_unique['Cliente'].unique()) if 'Cliente' in df_unique.columns else 0
    projetos_atrasados = len(df_unique[df_unique['Status'] == 'Atrasado']) if 'Status' in df_unique.columns else 0
    projetos_criticos = len(df_unique[df_unique['Prioridade'] == 'Crítico']) if 'Prioridade' in df_unique.columns else 0

    # Criar gráfico de status
    status_counts = df_unique['Status'].value_counts().reset_index() if 'Status' in df_unique.columns else pd.DataFrame(columns=['Status', 'Quantidade'])
    status_counts.columns = ['Status', 'Quantidade']

    status_fig = px.pie(
        status_counts, names='Status', values='Quantidade',
        title='Distribuição por Status',
        color_discrete_sequence=codeart_chart_palette,
    )
    status_fig.update_traces(textposition='inside', textinfo='percent+label')

    # Criar gráfico de financeiro
    financeiro_counts = df_unique['Financeiro'].value_counts().reset_index() if 'Financeiro' in df_unique.columns else pd.DataFrame(columns=['Financeiro', 'Quantidade'])
    financeiro_counts.columns = ['Financeiro', 'Quantidade']

    financeiro_fig = px.pie(
        financeiro_counts, names='Financeiro', values='Quantidade',
        title='Distribuição por Status Financeiro',
        color_discrete_sequence=codeart_chart_palette,
    )
    financeiro_fig.update_traces(textposition='inside', textinfo='percent+label')

    # Criar gráfico de NPS
    nps_counts = df_unique['NPS '].value_counts().reset_index() if 'NPS ' in df_unique.columns else pd.DataFrame(columns=['NPS', 'Quantidade'])
    nps_counts.columns = ['NPS', 'Quantidade']

    nps_fig = px.pie(
        nps_counts, names='NPS', values='Quantidade',
        title='Distribuição por NPS',
        color_discrete_sequence=codeart_chart_palette,
    )
    nps_fig.update_traces(textposition='inside', textinfo='percent+label')

    # Criar gráfico de Segmento
    segmento_counts = df_unique['Segmento'].value_counts().reset_index() if 'Segmento' in df_unique.columns else pd.DataFrame(columns=['Segmento', 'Quantidade'])
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
    gp_counts = df_unique['GP Responsável'].value_counts().reset_index() if 'GP Responsável' in df_unique.columns else pd.DataFrame(columns=['GP Responsável', 'Quantidade'])
    gp_counts.columns = ['GP Responsável', 'Quantidade']

    gp_fig = px.bar(
        gp_counts, x='GP Responsável', y='Quantidade',
        title='Projetos por Gestora',
        color_discrete_sequence=[codeart_colors['blue_sky']],
        text_auto=True
    )
    gp_fig.update_traces(textposition='outside')

    # NOVOS GRÁFICOS (usando df_unique ao invés de df)

    # Gráfico de Horas Previstas vs Realizadas
    horas_fig = go.Figure()
    if 'Previsão' in df_unique.columns and 'Real' in df_unique.columns and 'Projeto' in df_unique.columns:
        # Selecionar top 10 projetos por horas previstas
        top_projetos = df_unique.sort_values('Previsão', ascending=False).head(10)

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
    if 'Saldo Acumulado' in df_unique.columns and 'Projeto' in df_unique.columns:
        # Filtrar projetos com saldo não zero
        df_saldo = df_unique[df_unique['Saldo Acumulado'] != 0].copy()

        if not df_saldo.empty:
            # Ordenar por saldo (do menor para o maior)
            df_saldo = df_saldo.sort_values('Saldo Acumulado')

            # Limitar a 15 projetos para melhor visualização
            if len(df_saldo) > 15:
                df_saldo = pd.concat([df_saldo.head(7), df_saldo.tail(8)])

            # Verificar se há valores extremamente altos (potencialmente formatados incorretamente)
            for idx, saldo in enumerate(df_saldo['Saldo Acumulado']):
                if abs(saldo) > 1000:  # Se o valor for maior que 1000 (possivelmente um erro de formatação)
                    projeto = df_saldo.iloc[idx]['Projeto']
                    print(
                        f"Verificando saldo possivelmente incorreto para {projeto}: {saldo}")
                    # Se for um valor como -21058 quando deveria ser -210.58, corrigimos
                    if abs(saldo) > 1000 and abs(saldo) < 100000:
                        df_saldo.iloc[idx, df_saldo.columns.get_loc(
                            'Saldo Acumulado')] = saldo / 100
                        print(
                            f"Corrigido saldo de {projeto} de {saldo} para {saldo/100}")

            # Definir cores baseadas no saldo
            colors = ['#dc3545' if x <
                      0 else '#28a745' for x in df_saldo['Saldo Acumulado']]

            # Formatação textual personalizada para evitar notação científica ou 'k'
            text_values = [f"{x:.1f}" for x in df_saldo['Saldo Acumulado']]

            saldo_fig = go.Figure(data=[go.Bar(
                x=df_saldo['Projeto'],
                y=df_saldo['Saldo Acumulado'],
                marker_color=colors,
                text=text_values,
                textposition='outside'
            )])

            saldo_fig.update_layout(
                title='Saldo de Horas por Projeto',
                xaxis_tickangle=-45,
                yaxis=dict(
                    tickformat='.1f'  # Formato fixo com 1 casa decimal
                )
            )
        else:
            saldo_fig.update_layout(
                title="Sem projetos com saldo diferente de zero")
    else:
        saldo_fig.update_layout(title="Sem dados de saldo")

    # Gráfico de Atraso por Coordenação
    atraso_coord_fig = go.Figure()
    if 'Coordenação' in df_unique.columns and 'Status' in df_unique.columns:
        # Agrupar por coordenação e contar projetos atrasados
        atraso_coord_data = df_unique.groupby('Coordenação').apply(
            lambda x: pd.Series({
                'Total Projetos': len(x),
                'Projetos Atrasados': len(x[x['Status'] == 'Atrasado'])
            })
        ).reset_index()

        if not atraso_coord_data.empty:
            # Calcular percentual de projetos atrasados
            atraso_coord_data['Percentual'] = (
                atraso_coord_data['Projetos Atrasados'] / atraso_coord_data['Total Projetos'] * 100).round(1)

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
                y=atraso_coord_data['Total Projetos'] -
                atraso_coord_data['Projetos Atrasados'],
                name='Projetos no Prazo',
                marker_color=codeart_colors['success'],
                text=atraso_coord_data['Total Projetos'] -
                atraso_coord_data['Projetos Atrasados'],
                textposition='outside'
            ))

            atraso_coord_fig.update_layout(
                title='Projetos Atrasados por Coordenação',
                barmode='stack',
                xaxis_tickangle=-45
            )
        else:
            atraso_coord_fig.update_layout(
                title="Sem dados de atraso por coordenação")
    else:
        atraso_coord_fig.update_layout(
            title="Sem dados de coordenação ou status")

    # Gráfico de Evolução de Projetos Quitados
    evolucao_quitados_fig = go.Figure()
    if 'Financeiro' in df_unique.columns and 'MesAnoFormatado' in df_unique.columns:
        # Contar projetos quitados por mês
        quitados_por_mes = df_unique.groupby('MesAnoFormatado').apply(
            lambda x: len(x[x['Financeiro'] == 'Quitado'])
        ).reset_index()
        quitados_por_mes.columns = ['MesAnoFormatado', 'Projetos Quitados']
        
        # Extrair mês e ano para ordenação cronológica
        def extrair_mes_ano(mes_ano_str):
            # Formato esperado: 'Mmm/AAAA' (ex: Jan/2023)
            meses = {'Jan': 1, 'Fev': 2, 'Mar': 3, 'Abr': 4, 'Mai': 5, 'Jun': 6,
                     'Jul': 7, 'Ago': 8, 'Set': 9, 'Out': 10, 'Nov': 11, 'Dez': 12}
            try:
                mes_abrev = mes_ano_str.split('/')[0]
                ano = int(mes_ano_str.split('/')[1])
                mes_num = meses.get(mes_abrev, 0)
                return ano * 100 + mes_num  # Exemplo: Jan/2023 = 202301
            except:
                return 0
        
        # Adicionar coluna de ordenação e ordenar os dados
        quitados_por_mes['ordem'] = quitados_por_mes['MesAnoFormatado'].apply(extrair_mes_ano)
        quitados_por_mes = quitados_por_mes.sort_values('ordem')
        
        if not quitados_por_mes.empty:
            # Criar gráfico com ordem fixa dos meses
            evolucao_quitados_fig = px.line(
                quitados_por_mes, x='MesAnoFormatado', y='Projetos Quitados',
                title='Evolução de Projetos Quitados',
                markers=True,
                color_discrete_sequence=[codeart_colors['success']]
            )

            # Garantir que a ordem dos meses no eixo X seja mantida conforme os dados ordenados
            evolucao_quitados_fig.update_layout(
                xaxis=dict(
                    categoryorder='array',
                    categoryarray=quitados_por_mes['MesAnoFormatado'].tolist(),
                    tickangle=-45
                )
            )
        else:
            evolucao_quitados_fig.update_layout(
                title="Sem dados de projetos quitados")
    else:
        evolucao_quitados_fig.update_layout(
            title="Sem dados de financeiro ou período")

    # Gráfico de Evolução de Projetos Atrasados
    evolucao_atrasados_fig = go.Figure()
    if 'Status' in df_unique.columns and 'MesAnoFormatado' in df_unique.columns:
        # Contar projetos atrasados por mês
        atrasados_por_mes = df_unique.groupby('MesAnoFormatado').apply(
            lambda x: len(x[x['Status'] == 'Atrasado'])
        ).reset_index()
        atrasados_por_mes.columns = ['MesAnoFormatado', 'Projetos Atrasados']
        
        # Extrair mês e ano para ordenação cronológica - reutilizando a função definida acima
        atrasados_por_mes['ordem'] = atrasados_por_mes['MesAnoFormatado'].apply(extrair_mes_ano)
        atrasados_por_mes = atrasados_por_mes.sort_values('ordem')

        if not atrasados_por_mes.empty:
            # Criar gráfico com ordem fixa dos meses
            evolucao_atrasados_fig = px.line(
                atrasados_por_mes, x='MesAnoFormatado', y='Projetos Atrasados',
                title='Evolução de Projetos Atrasados',
                markers=True,
                color_discrete_sequence=[codeart_colors['danger']]
            )

            # Garantir que a ordem dos meses no eixo X seja mantida conforme os dados ordenados
            evolucao_atrasados_fig.update_layout(
                xaxis=dict(
                    categoryorder='array',
                    categoryarray=atrasados_por_mes['MesAnoFormatado'].tolist(),
                    tickangle=-45
                )
            )
        else:
            evolucao_atrasados_fig.update_layout(
                title="Sem dados de projetos atrasados")
    else:
        evolucao_atrasados_fig.update_layout(
            title="Sem dados de status ou período")

    # Retornar a tabela com todos os registros, mas totalizadores e gráficos só com projetos únicos
    return str(total_projetos), str(total_clientes), str(projetos_atrasados), str(projetos_criticos), status_fig, df_table.to_dict('records'), financeiro_fig, nps_fig, segmento_fig, gp_fig, horas_fig, saldo_fig, atraso_coord_fig, evolucao_quitados_fig, evolucao_atrasados_fig

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
    options = [{"label": nome, "value": nome}
               for nome in sorted(df_codenautas['Nome'].unique())]
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
    options = [{"label": nome, "value": nome}
               for nome in sorted(df_codenautas['Nome'].unique())]
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
    options = [{"label": nome, "value": nome}
               for nome in sorted(df_codenautas['Nome'].unique())]
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
    options = [{"label": nome, "value": nome}
               for nome in sorted(df_codenautas['Nome'].unique())]
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

    coordenacoes = [{"label": coord, "value": coord}
                    for coord in filter_options_data.get("coordenacoes", [])]
    meses_anos = [{"label": mes, "value": mes}
                  for mes in filter_options_data.get("meses_anos", [])]
    gestoras = [{"label": gp, "value": gp}
                for gp in filter_options_data.get("gestoras", [])]
    status_list = [{"label": status, "value": status}
                   for status in filter_options_data.get("status_list", [])]
    financeiro_list = [{"label": fin, "value": fin}
                       for fin in filter_options_data.get("financeiro_list", [])]
    segmentos = [{"label": seg, "value": seg}
                 for seg in filter_options_data.get("segmentos", [])]
    tipos = [{"label": tipo, "value": tipo}
             for tipo in filter_options_data.get("tipos", [])]

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
    options = [{"label": projeto, "value": projeto}
               for projeto in sorted(df['Projeto'].unique())]
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
    options = [{"label": projeto, "value": projeto}
               for projeto in sorted(df['Projeto'].unique())]
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
    options = [{"label": projeto, "value": projeto}
               for projeto in sorted(df['Projeto'].unique())]
    return options

# Callback para preencher as opções do dropdown de mês de referência no modal de nova ação


@app.callback(
    Output("modal-acao-mes-referencia", "options"),
    Input("filter-options-store", "data")
)
def update_acao_mes_referencia_options(filter_options_data):
    if not filter_options_data:
        return []
    
    # Primeiro, adicionar as opções de meses do ano (Janeiro a Dezembro) com nome completo
    meses_opcoes = [
        {"label": "Janeiro", "value": "Janeiro"},
        {"label": "Fevereiro", "value": "Fevereiro"},
        {"label": "Março", "value": "Março"},
        {"label": "Abril", "value": "Abril"},
        {"label": "Maio", "value": "Maio"},
        {"label": "Junho", "value": "Junho"},
        {"label": "Julho", "value": "Julho"},
        {"label": "Agosto", "value": "Agosto"},
        {"label": "Setembro", "value": "Setembro"},
        {"label": "Outubro", "value": "Outubro"},
        {"label": "Novembro", "value": "Novembro"},
        {"label": "Dezembro", "value": "Dezembro"}
    ]
    
    # Dicionário para converter abreviações para nomes completos
    meses_map = {
        "Jan/2023": "Janeiro",
        "Fev/2023": "Fevereiro",
        "Mar/2023": "Março",
        "Abr/2023": "Abril",
        "Mai/2023": "Maio",
        "Jun/2023": "Junho",
        "Jul/2023": "Julho",
        "Ago/2023": "Agosto",
        "Set/2023": "Setembro",
        "Out/2023": "Outubro",
        "Nov/2023": "Novembro",
        "Dez/2023": "Dezembro"
    }
    
    # Adicionar também as opções existentes nos dados
    if 'meses_anos' in filter_options_data:
        for mes in filter_options_data.get("meses_anos", []):
            # Verificar se é uma abreviação que podemos converter
            nome_completo = meses_map.get(mes, mes)
            
            # Verificar se o mês já não está nas opções
            if not any(op["value"] == nome_completo for op in meses_opcoes):
                meses_opcoes.append({"label": nome_completo, "value": nome_completo})
    
    return meses_opcoes

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
                filtered_df = filtered_df[filtered_df['Mês de Referência'].isin(
                    mes_ano)]

            # Filtrar por responsável
            if responsavel:
                if not isinstance(responsavel, list):
                    responsavel = [responsavel]
                # Considerando que um responsável pode estar em uma lista separada por vírgulas ou como string única
                mask = filtered_df['Responsáveis'].apply(
                    lambda x: any(resp in str(x).split(',')
                                  for resp in responsavel)
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
                filtered_df = filtered_df[filtered_df['Prioridade'].isin(
                    prioridade)]
        else:
            filtered_df = df_acoes
    else:
        filtered_df = df_acoes

    # Calcular métricas
    total_acoes = len(filtered_df)
    pendentes = len(filtered_df[filtered_df['Status'] != 'Concluída'])
    concluidas = len(filtered_df[filtered_df['Status'] == 'Concluída'])
    atrasadas = filtered_df['Atrasada'].sum(
    ) if 'Atrasada' in filtered_df.columns else 0

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
    prioridade_fig.update_traces(
        textposition='inside', textinfo='percent+label')

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
        responsaveis_counts = pd.Series(
            responsaveis_expandidos).value_counts().reset_index()
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
        filtered_df['Data de Cadastro'] = pd.to_datetime(
            filtered_df['Data de Cadastro'], errors='coerce')

        # Criar coluna de mês/ano para agrupamento
        filtered_df['Mês Cadastro'] = filtered_df['Data de Cadastro'].dt.strftime(
            '%Y-%m')

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

    # Preparar dados para a tabela
    try:
        # Criar uma cópia do DataFrame filtrado para manipulação
        table_df = filtered_df.copy()
        
        # Formatar datas para exibição amigável
        # Converter datas para datetime e depois para o formato brasileiro dd/mm/yyyy
        for col in ['Data de Cadastro', 'Data Limite', 'Data de Conclusão']:
            if col in table_df.columns:
                table_df[col] = pd.to_datetime(table_df[col], errors='coerce')
                table_df[col] = table_df[col].apply(
                    lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) else '')

        # Garantir que observações vazias sejam strings vazias e não None
        if 'Observações de conclusão' in table_df.columns:
            table_df['Observações de conclusão'] = table_df['Observações de conclusão'].fillna('')
        
        # Adicionar indicadores de status
        if 'Status' in table_df.columns and 'Atrasada' in table_df.columns:
            # Permitir filtro por status e atrasos
            def get_status_detalhado(row):
                status = row['Status']
                if status == 'Concluída':
                    return 'Concluída'
                elif row['Atrasada'] == 1:
                    return 'Atrasada'
                else:
                    return status
            
            table_df['Status Detalhado'] = table_df.apply(get_status_detalhado, axis=1)
        
        # Preparar dados amigáveis para a tabela
        table_data = table_df.to_dict('records')
    except Exception as e:
        print(f"Erro ao preparar dados para tabela de ações: {e}")
        import traceback
        traceback.print_exc()
        # Em caso de erro, usar os dados originais
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
    # Definir figuras vazias para usar como fallback
    empty_fig = go.Figure().update_layout(title="Sem dados disponíveis")
    status_fig = empty_fig
    financeiro_fig = empty_fig
    nps_fig = empty_fig
    segmento_fig = empty_fig
    gp_fig = empty_fig
    horas_fig = empty_fig
    saldo_fig = empty_fig
    atraso_coord_fig = empty_fig
    evolucao_quitados_fig = empty_fig
    evolucao_atrasados_fig = empty_fig
    df_table = pd.DataFrame()

    # Se não houver dados, retornar valores vazios
    if not data:
        return "0", "0", "0", "0", status_fig, df_table.to_dict('records'), financeiro_fig, nps_fig, segmento_fig, gp_fig, horas_fig, saldo_fig, atraso_coord_fig, evolucao_quitados_fig, evolucao_atrasados_fig

    # Converter para DataFrame
    df = pd.DataFrame(data)

    # Se o DataFrame estiver vazio, retornar valores vazios
    if df.empty:
        return "0", "0", "0", "0", status_fig, df_table.to_dict('records'), financeiro_fig, nps_fig, segmento_fig, gp_fig, horas_fig, saldo_fig, atraso_coord_fig, evolucao_quitados_fig, evolucao_atrasados_fig

    # Verificar qual botão foi clicado
    ctx = dash.callback_context
    button_id = ctx.triggered[0]['prop_id'].split(
        '.')[0] if ctx.triggered else None

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
            filtered_df = filtered_df[filtered_df['MesAnoFormatado'].isin(
                mes_ano)]

        # Filtrar por gestora
        if gestora:
            if not isinstance(gestora, list):
                gestora = [gestora]
            filtered_df = filtered_df[filtered_df['GP Responsável'].isin(
                gestora)]

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
            filtered_df = filtered_df[filtered_df['Coordenação'].isin(
                coordenacao)]

        # Filtrar por financeiro
        if financeiro:
            if not isinstance(financeiro, list):
                financeiro = [financeiro]
            filtered_df = filtered_df[filtered_df['Financeiro'].isin(
                financeiro)]
    else:
        filtered_df = df

    # Criar cópia para exibição na tabela (mostra todos os registros filtrados)
    df_table = filtered_df.copy()
    
    # Identificar projetos únicos para os totalizadores e gráficos
    # Consideramos um projeto como único combinando o nome do projeto e cliente
    if 'Projeto' in filtered_df.columns and 'Cliente' in filtered_df.columns:
        filtered_df['projeto_cliente'] = filtered_df['Projeto'] + ' - ' + filtered_df['Cliente']
        # Pegamos a versão mais recente de cada projeto (assumindo que MesAnoFormatado está presente)
        if 'MesAnoFormatado' in filtered_df.columns:
            # Ordenar por data (o mais recente primeiro)
            filtered_df = filtered_df.sort_values('MesAnoFormatado', ascending=False)
            # Remove registros duplicados, mantendo apenas o primeiro (mais recente) de cada projeto
            df_unique = filtered_df.drop_duplicates(subset=['projeto_cliente'])
        else:
            # Se não houver data, apenas remover duplicatas
            df_unique = filtered_df.drop_duplicates(subset=['projeto_cliente'])
    else:
        # Se não tiver as colunas necessárias, usar o dataframe original
        df_unique = filtered_df.copy()

    # Calcular métricas com projetos únicos
    total_projetos = len(df_unique)
    total_clientes = len(df_unique['Cliente'].unique())
    projetos_atrasados = len(df_unique[df_unique['Status'] == 'Atrasado'])
    projetos_criticos = len(df_unique[df_unique['Prioridade'] == 'Crítico'])

    # Criar gráfico de status
    status_counts = df_unique['Status'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Quantidade']

    status_fig = px.pie(
        status_counts, names='Status', values='Quantidade',
        title='Distribuição por Status',
        color_discrete_sequence=codeart_chart_palette,
    )
    status_fig.update_traces(textposition='inside', textinfo='percent+label')

    # Criar gráfico de financeiro
    financeiro_counts = df_unique['Financeiro'].value_counts().reset_index()
    financeiro_counts.columns = ['Financeiro', 'Quantidade']

    financeiro_fig = px.pie(
        financeiro_counts, names='Financeiro', values='Quantidade',
        title='Distribuição por Status Financeiro',
        color_discrete_sequence=codeart_chart_palette,
    )
    financeiro_fig.update_traces(textposition='inside', textinfo='percent+label')

    # Criar gráfico de NPS
    nps_counts = df_unique['NPS '].value_counts().reset_index()
    nps_counts.columns = ['NPS', 'Quantidade']

    nps_fig = px.pie(
        nps_counts, names='NPS', values='Quantidade',
        title='Distribuição por NPS',
        color_discrete_sequence=codeart_chart_palette,
    )
    nps_fig.update_traces(textposition='inside', textinfo='percent+label')

    # Criar gráfico de Segmento
    segmento_counts = df_unique['Segmento'].value_counts().reset_index() if 'Segmento' in df_unique.columns else pd.DataFrame(columns=['Segmento', 'Quantidade'])
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
    gp_counts = df_unique['GP Responsável'].value_counts().reset_index()
    gp_counts.columns = ['GP Responsável', 'Quantidade']

    gp_fig = px.bar(
        gp_counts, x='GP Responsável', y='Quantidade',
        title='Projetos por Gestora',
        color_discrete_sequence=[codeart_colors['blue_sky']],
        text_auto=True
    )
    gp_fig.update_traces(textposition='outside')

    # NOVOS GRÁFICOS (usando df_unique ao invés de filtered_df)

    # Gráfico de Horas Previstas vs Realizadas
    horas_fig = go.Figure()
    if 'Previsão' in df_unique.columns and 'Real' in df_unique.columns and 'Projeto' in df_unique.columns:
        # Selecionar top 10 projetos por horas previstas
        top_projetos = df_unique.sort_values('Previsão', ascending=False).head(10)

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
    if 'Saldo Acumulado' in df_unique.columns and 'Projeto' in df_unique.columns:
        # Filtrar projetos com saldo não zero
        df_saldo = df_unique[df_unique['Saldo Acumulado'] != 0].copy()

        if not df_saldo.empty:
            # Ordenar por saldo (do menor para o maior)
            df_saldo = df_saldo.sort_values('Saldo Acumulado')

            # Limitar a 15 projetos para melhor visualização
            if len(df_saldo) > 15:
                df_saldo = pd.concat([df_saldo.head(7), df_saldo.tail(8)])

            # Verificar se há valores extremamente altos (potencialmente formatados incorretamente)
            for idx, saldo in enumerate(df_saldo['Saldo Acumulado']):
                if abs(saldo) > 1000:  # Se o valor for maior que 1000 (possivelmente um erro de formatação)
                    projeto = df_saldo.iloc[idx]['Projeto']
                    print(
                        f"Verificando saldo possivelmente incorreto para {projeto}: {saldo}")
                    # Se for um valor como -21058 quando deveria ser -210.58, corrigimos
                    if abs(saldo) > 1000 and abs(saldo) < 100000:
                        df_saldo.iloc[idx, df_saldo.columns.get_loc(
                            'Saldo Acumulado')] = saldo / 100
                        print(
                            f"Corrigido saldo de {projeto} de {saldo} para {saldo/100}")

            # Definir cores baseadas no saldo
            colors = ['#dc3545' if x <
                      0 else '#28a745' for x in df_saldo['Saldo Acumulado']]

            # Formatação textual personalizada para evitar notação científica ou 'k'
            text_values = [f"{x:.1f}" for x in df_saldo['Saldo Acumulado']]

            saldo_fig = go.Figure(data=[go.Bar(
                x=df_saldo['Projeto'],
                y=df_saldo['Saldo Acumulado'],
                marker_color=colors,
                text=text_values,
                textposition='outside'
            )])

            saldo_fig.update_layout(
                title='Saldo de Horas por Projeto',
                xaxis_tickangle=-45,
                yaxis=dict(
                    tickformat='.1f'  # Formato fixo com 1 casa decimal
                )
            )
        else:
            saldo_fig.update_layout(
                title="Sem projetos com saldo diferente de zero")
    else:
        saldo_fig.update_layout(title="Sem dados de saldo")

    # Gráfico de Atraso por Coordenação
    atraso_coord_fig = go.Figure()
    if 'Coordenação' in df_unique.columns and 'Status' in df_unique.columns:
        # Agrupar por coordenação e contar projetos atrasados
        atraso_coord_data = df_unique.groupby('Coordenação').apply(
            lambda x: pd.Series({
                'Total Projetos': len(x),
                'Projetos Atrasados': len(x[x['Status'] == 'Atrasado'])
            })
        ).reset_index()

        if not atraso_coord_data.empty:
            # Calcular percentual de projetos atrasados
            atraso_coord_data['Percentual'] = (
                atraso_coord_data['Projetos Atrasados'] / atraso_coord_data['Total Projetos'] * 100).round(1)

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
                y=atraso_coord_data['Total Projetos'] -
                atraso_coord_data['Projetos Atrasados'],
                name='Projetos no Prazo',
                marker_color=codeart_colors['success'],
                text=atraso_coord_data['Total Projetos'] -
                atraso_coord_data['Projetos Atrasados'],
                textposition='outside'
            ))

            atraso_coord_fig.update_layout(
                title='Projetos Atrasados por Coordenação',
                barmode='stack',
                xaxis_tickangle=-45
            )
        else:
            atraso_coord_fig.update_layout(
                title="Sem dados de atraso por coordenação")
    else:
        atraso_coord_fig.update_layout(
            title="Sem dados de coordenação ou status")

    # Gráfico de Evolução de Projetos Quitados
    evolucao_quitados_fig = go.Figure()
    if 'Financeiro' in df_unique.columns and 'MesAnoFormatado' in df_unique.columns:
        # Contar projetos quitados por mês
        quitados_por_mes = df_unique.groupby('MesAnoFormatado').apply(
            lambda x: len(x[x['Financeiro'] == 'Quitado'])
        ).reset_index()
        quitados_por_mes.columns = ['MesAnoFormatado', 'Projetos Quitados']
        
        # Extrair mês e ano para ordenação cronológica
        def extrair_mes_ano(mes_ano_str):
            # Formato esperado: 'Mmm/AAAA' (ex: Jan/2023)
            meses = {'Jan': 1, 'Fev': 2, 'Mar': 3, 'Abr': 4, 'Mai': 5, 'Jun': 6,
                     'Jul': 7, 'Ago': 8, 'Set': 9, 'Out': 10, 'Nov': 11, 'Dez': 12}
            try:
                mes_abrev = mes_ano_str.split('/')[0]
                ano = int(mes_ano_str.split('/')[1])
                mes_num = meses.get(mes_abrev, 0)
                return ano * 100 + mes_num  # Exemplo: Jan/2023 = 202301
            except:
                return 0
        
        # Adicionar coluna de ordenação e ordenar os dados
        quitados_por_mes['ordem'] = quitados_por_mes['MesAnoFormatado'].apply(extrair_mes_ano)
        quitados_por_mes = quitados_por_mes.sort_values('ordem')
        
        if not quitados_por_mes.empty:
            # Criar gráfico com ordem fixa dos meses
            evolucao_quitados_fig = px.line(
                quitados_por_mes, x='MesAnoFormatado', y='Projetos Quitados',
                title='Evolução de Projetos Quitados',
                markers=True,
                color_discrete_sequence=[codeart_colors['success']]
            )

            # Garantir que a ordem dos meses no eixo X seja mantida conforme os dados ordenados
            evolucao_quitados_fig.update_layout(
                xaxis=dict(
                    categoryorder='array',
                    categoryarray=quitados_por_mes['MesAnoFormatado'].tolist(),
                    tickangle=-45
                )
            )
        else:
            evolucao_quitados_fig.update_layout(
                title="Sem dados de projetos quitados")
    else:
        evolucao_quitados_fig.update_layout(
            title="Sem dados de financeiro ou período")

    # Gráfico de Evolução de Projetos Atrasados
    evolucao_atrasados_fig = go.Figure()
    if 'Status' in df_unique.columns and 'MesAnoFormatado' in df_unique.columns:
        # Contar projetos atrasados por mês
        atrasados_por_mes = df_unique.groupby('MesAnoFormatado').apply(
            lambda x: len(x[x['Status'] == 'Atrasado'])
        ).reset_index()
        atrasados_por_mes.columns = ['MesAnoFormatado', 'Projetos Atrasados']
        
        # Extrair mês e ano para ordenação cronológica - reutilizando a função definida acima
        atrasados_por_mes['ordem'] = atrasados_por_mes['MesAnoFormatado'].apply(extrair_mes_ano)
        atrasados_por_mes = atrasados_por_mes.sort_values('ordem')

        if not atrasados_por_mes.empty:
            # Criar gráfico com ordem fixa dos meses
            evolucao_atrasados_fig = px.line(
                atrasados_por_mes, x='MesAnoFormatado', y='Projetos Atrasados',
                title='Evolução de Projetos Atrasados',
                markers=True,
                color_discrete_sequence=[codeart_colors['danger']]
            )

            # Garantir que a ordem dos meses no eixo X seja mantida conforme os dados ordenados
            evolucao_atrasados_fig.update_layout(
                xaxis=dict(
                    categoryorder='array',
                    categoryarray=atrasados_por_mes['MesAnoFormatado'].tolist(),
                    tickangle=-45
                )
            )
        else:
            evolucao_atrasados_fig.update_layout(
                title="Sem dados de projetos atrasados")
    else:
        evolucao_atrasados_fig.update_layout(
            title="Sem dados de status ou período")

    # Retornar a tabela com todos os registros filtrados, mas totalizadores e gráficos só com projetos únicos
    return str(total_projetos), str(total_clientes), str(projetos_atrasados), str(projetos_criticos), df_table.to_dict('records'), status_fig, financeiro_fig, nps_fig, segmento_fig, gp_fig, horas_fig, saldo_fig, atraso_coord_fig, evolucao_quitados_fig, evolucao_atrasados_fig

# Callback para adicionar ícone de ação na tabela de projetos


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
    
    # Garantir que a coluna MesAnoFormatado existe
    if 'Mês' in df.columns and 'MesAnoFormatado' not in df.columns:
        # Processar dados para adicionar a coluna MesAnoFormatado
        df = process_data(df)
    
    # Adicionar coluna de ícone de ação (usando texto simples em vez de markdown)
    df['action_icon'] = '+'
    
    # Se já temos dados na tabela, manter apenas a coluna action_icon e mesclar com os dados existentes
    if table_data and len(table_data) > 0:
        df_table = pd.DataFrame(table_data)
        if 'action_icon' not in df_table.columns:
            df_table['action_icon'] = '+'
        if 'MesAnoFormatado' not in df_table.columns and 'MesAnoFormatado' in df.columns:
            # Adicionar a coluna MesAnoFormatado dos dados originais
            mes_dict = {row['Projeto']: row['MesAnoFormatado'] for _, row in df.iterrows() if 'Projeto' in row and 'MesAnoFormatado' in row}
            df_table['MesAnoFormatado'] = df_table['Projeto'].map(mes_dict).fillna('')
        return df_table.to_dict('records')
    
    return df.to_dict('records')

# Callback para capturar clique na célula do ícone de ação
@app.callback(
    [
        Output("selected-project-store", "data"),
        Output("modal-cadastro-acao", "is_open"),
        Output("modal-projeto", "value"),
        Output("modal-mes-referencia", "value")
    ],
    Input("projetos-table", "selected_cells"),
    [
        State("projetos-table", "data"),
        State("mes-ano-filter", "value")  # Obter o mês/ano selecionado no filtro atual
    ],
    prevent_initial_call=True
)
def handle_action_icon_click(selected_cells, table_data, mes_ano_atual):
    if not selected_cells or not table_data:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update

    # Verificar se a célula selecionada é da coluna de ação
    if selected_cells[0]['column_id'] == 'action_icon':
        row_idx = selected_cells[0]['row']
        projeto = table_data[row_idx]['Projeto']
        
        # Obter o mês/ano do projeto para preencher automaticamente
        mes_referencia = None
        if 'MesAnoFormatado' in table_data[row_idx]:
            mes_referencia = table_data[row_idx]['MesAnoFormatado']
        elif mes_ano_atual and not isinstance(mes_ano_atual, list):
            # Usar o mês/ano selecionado no filtro se não houver no projeto
            mes_referencia = mes_ano_atual
        
        print(f"Adicionando ação para o projeto: {projeto}, mês: {mes_referencia}")
        return projeto, True, projeto, mes_referencia

    return dash.no_update, dash.no_update, dash.no_update, dash.no_update

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

    print("\n===== Salvando nova ação (modal principal) =====")

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
        print(f"❌ Campos vazios no cadastro da ação: {', '.join(campos_vazios)}")
        return dash.no_update, True, mensagem_erro, dash.no_update

    try:
        # Converter responsáveis para string
        if isinstance(responsaveis, list):
            responsaveis_str = ', '.join(responsaveis)
        else:
            responsaveis_str = str(responsaveis)

        print(f"Projeto: {projeto}")
        print(f"Mês Referência: {mes_referencia}")
        print(f"Responsáveis: {responsaveis_str}")
        print(f"Status: {status}")

        # Carregar ações novamente direto da planilha para garantir que estamos trabalhando com os dados mais atuais
        try:
            print("Tentando recarregar ações diretamente do Google Sheets antes de adicionar nova ação...")
            spreadsheet = connect_google_sheets()
            if spreadsheet:
                sheet = spreadsheet.worksheet('Ações')
                sheet_data = sheet.get_all_records()
                
                if sheet_data:
                    df_atual = pd.DataFrame(sheet_data)
                    print(f"✅ Carregadas {len(df_atual)} ações do Google Sheets")
                    
                    # Verificar se já existe o ID no dataframe atual
                    max_id = 0
                    if 'ID da Ação' in df_atual.columns:
                        ids_acao = pd.to_numeric(df_atual['ID da Ação'], errors='coerce')
                        max_id = ids_acao.max() if not pd.isna(ids_acao.max()) else 0
                        
                    next_id = int(max_id) + 1
                    print(f"Próximo ID baseado nos dados atuais da planilha: {next_id}")
                    
                    # Atualizar acoes_data com os dados da planilha
                    acoes_data = df_atual.to_dict('records')
                else:
                    print("Nenhum dado encontrado na guia Ações do Google Sheets")
                    # Gerar ID baseado nos dados do cache
                    next_id = 1
                    if acoes_data and len(acoes_data) > 0:
                        df_acoes = pd.DataFrame(acoes_data)
                        if 'ID da Ação' in df_acoes.columns:
                            ids_numericos = pd.to_numeric(df_acoes['ID da Ação'], errors='coerce')
                            next_id = int(ids_numericos.max()) + 1 if not pd.isna(ids_numericos.max()) else 1
            else:
                print("Não foi possível conectar ao Google Sheets para recarregar ações")
                # Inicializar acoes_data como lista vazia se for None
                if acoes_data is None:
                    acoes_data = []
                    print("Nenhuma ação encontrada no cache, iniciando lista vazia")
                elif not isinstance(acoes_data, list):
                    # Tentar converter para lista se não for
                    try:
                        acoes_data = list(acoes_data)
                        print(f"Convertido acoes_data para lista com {len(acoes_data)} itens")
                    except Exception as conv_e:
                        print(f"Erro ao converter acoes_data: {conv_e}")
                        acoes_data = []
                
                # Determinar próximo ID usando os dados em cache
                next_id = 1
                if acoes_data and len(acoes_data) > 0:
                    df_acoes = pd.DataFrame(acoes_data)
                    if 'ID da Ação' in df_acoes.columns:
                        ids_numericos = pd.to_numeric(df_acoes['ID da Ação'], errors='coerce')
                        next_id = int(ids_numericos.max()) + 1 if not pd.isna(ids_numericos.max()) else 1
        except Exception as reload_e:
            print(f"Erro ao recarregar ações: {reload_e}")
            # Inicializar acoes_data como lista vazia se for None
            if acoes_data is None:
                acoes_data = []
                print("Nenhuma ação encontrada no cache, iniciando lista vazia")
            elif not isinstance(acoes_data, list):
                # Tentar converter para lista se não for
                try:
                    acoes_data = list(acoes_data)
                    print(f"Convertido acoes_data para lista com {len(acoes_data)} itens")
                except Exception as conv_e:
                    print(f"Erro ao converter acoes_data: {conv_e}")
                    acoes_data = []
            
            # Determinar próximo ID usando os dados em cache
            next_id = 1
            if acoes_data and len(acoes_data) > 0:
                df_acoes = pd.DataFrame(acoes_data)
                if 'ID da Ação' in df_acoes.columns:
                    ids_numericos = pd.to_numeric(df_acoes['ID da Ação'], errors='coerce')
                    next_id = int(ids_numericos.max()) + 1 if not pd.isna(ids_numericos.max()) else 1

        # Preparar nova linha
        # Formatar corretamente as datas para salvar no formato correto (YYYY-MM-DD)
        data_cadastro_formatada = datetime.now().strftime('%Y-%m-%d')
        
        # Formatar data limite corretamente
        data_limite_formatada = None
        if data_limite:
            # Verificar se já está no formato string
            if isinstance(data_limite, str):
                # Se estiver no formato ISO (YYYY-MM-DD), manter como está
                if '-' in data_limite and len(data_limite.split('-')) == 3:
                    data_limite_formatada = data_limite
                # Se estiver em outro formato, tentar converter
                else:
                    try:
                        data_obj = pd.to_datetime(data_limite)
                        data_limite_formatada = data_obj.strftime('%Y-%m-%d')
                    except:
                        data_limite_formatada = data_limite
            # Se for objeto datetime, converter para string
            elif hasattr(data_limite, 'strftime'):
                data_limite_formatada = data_limite.strftime('%Y-%m-%d')
            # Se for outra coisa, usar como está
            else:
                data_limite_formatada = data_limite
        
        # Formatar data de conclusão
        data_conclusao_formatada = None
        if data_conclusao:
            # Verificar se já está no formato string
            if isinstance(data_conclusao, str):
                # Se estiver no formato ISO (YYYY-MM-DD), manter como está
                if '-' in data_conclusao and len(data_conclusao.split('-')) == 3:
                    data_conclusao_formatada = data_conclusao
                # Se estiver em outro formato, tentar converter
                else:
                    try:
                        data_obj = pd.to_datetime(data_conclusao)
                        data_conclusao_formatada = data_obj.strftime('%Y-%m-%d')
                    except:
                        data_conclusao_formatada = data_conclusao
            # Se for objeto datetime, converter para string
            elif hasattr(data_conclusao, 'strftime'):
                data_conclusao_formatada = data_conclusao.strftime('%Y-%m-%d')
            # Se for outra coisa, usar como está
            else:
                data_conclusao_formatada = data_conclusao
        
        print(f"Data limite original: {data_limite}")
        print(f"Data limite formatada: {data_limite_formatada}")
        
        nova_acao = {
            'ID da Ação': next_id,
            'Data de Cadastro': data_cadastro_formatada,
            'Mês de Referência': mes_referencia,
            'Projeto': projeto,
            'Descrição da Ação': descricao,
            'Responsáveis': responsaveis_str,
            'Data Limite': data_limite_formatada,
            'Status': status,
            'Prioridade': prioridade,
            'Data de Conclusão': data_conclusao_formatada,
            'Observações de conclusão': ""
        }

        # Adicionar nova ação aos dados existentes
        acoes_data.append(nova_acao)
        print(f"Nova ação adicionada ao cache (ID: {next_id})")

        # Atualizar dados na planilha do Google Sheets
        df_acoes = pd.DataFrame(acoes_data)
        
        # Salvar localmente antes de enviar ao Google Sheets (backup)
        save_data_to_local(df_acoes, "acoes_pre_upload")
        
        # Tentar atualizar no Google Sheets
        success = update_acoes_in_sheets(df_acoes)
        if success:
            print(f"✅ Ação cadastrada com sucesso: ID {next_id}")
            
            # Garantir que o cache está atualizado
            global CACHE_ACOES, LAST_CACHE_UPDATE
            CACHE_ACOES = df_acoes
            LAST_CACHE_UPDATE = time.time()
        else:
            print(f"⚠️ Falha ao salvar ação na planilha do Google Sheets, mas foi salva localmente")
            save_data_to_local(df_acoes, "acoes")

        print("===== Nova ação salva (modal principal) =====\n")
        return False, False, "", acoes_data

    except Exception as e:
        print(f"❌ Erro ao salvar ação: {e}")
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
        State("modal-edit-status", "value"),
        State("modal-edit-data-conclusao", "date"),
        State("modal-edit-observacoes", "value"),
        State("acoes-store", "data")
    ],
    prevent_initial_call=True
)
def save_action_edit(n_clicks, acao_id, projeto, mes_referencia, prioridade, descricao, responsaveis, status, data_conclusao, observacoes, acoes_data):
    if not n_clicks:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update

    print(f"\n===== Salvando edição da ação ID: {acao_id} =====")
    print(f"Mês de referência: {mes_referencia} (tipo: {type(mes_referencia)})")

    # Validação para campos obrigatórios
    campos_vazios = []

    if not projeto:
        campos_vazios.append("Projeto")

    if not mes_referencia:
        campos_vazios.append("Mês de Referência")

    if not prioridade:
        campos_vazios.append("Prioridade")

    if not descricao:
        campos_vazios.append("Descrição da Ação")

    if not responsaveis or len(responsaveis) == 0:
        campos_vazios.append("Responsáveis")

    if campos_vazios:
        mensagem = f"Preencha os seguintes campos obrigatórios: {', '.join(campos_vazios)}"
        return False, True, mensagem, dash.no_update

    # Encontrar o índice da ação existente no DataFrame
    if acoes_data:
        df_acoes = pd.DataFrame(acoes_data)
        acao_id_str = str(acao_id)  # Converter o ID para string para garantir correspondência
        
        # Verificar se a ação existe
        acoes_filtradas = df_acoes.loc[df_acoes['ID da Ação'].astype(str) == acao_id_str]
        
        if len(acoes_filtradas) > 0:
            idx = acoes_filtradas.index[0]
            
            # Guardar uma cópia dos valores originais para diagnóstico
            data_limite_original = df_acoes.at[idx, 'Data Limite'] if 'Data Limite' in df_acoes.columns else None
            print(f"Data limite original: {data_limite_original} (tipo: {type(data_limite_original)})")
            
            # Atualizar os dados da ação
            df_acoes.at[idx, 'Projeto'] = projeto
            df_acoes.at[idx, 'Mês de Referência'] = mes_referencia
            df_acoes.at[idx, 'Prioridade'] = prioridade
            df_acoes.at[idx, 'Descrição da Ação'] = descricao
            df_acoes.at[idx, 'Responsáveis'] = ', '.join(responsaveis) if isinstance(responsaveis, list) else responsaveis
            
            # Data Limite original é mantida - não foi removida do layout
            print(f"Mantendo a data limite original: {data_limite_original}")
            
            if status:
                df_acoes.at[idx, 'Status'] = status
                
                # Se status for "Concluída" e não houver data de conclusão, definir como data atual
                if status == "Concluída" and not data_conclusao:
                    data_conclusao = datetime.now().strftime('%Y-%m-%d')
                    print("Status 'Concluída' selecionado, definindo data de conclusão para hoje")
            
            # Voltando ao tratamento original da Data de Conclusão
            # Se houver data de conclusão, usar o valor fornecido
            if data_conclusao:
                df_acoes.at[idx, 'Data de Conclusão'] = data_conclusao
                print(f"Data de Conclusão atualizada para: {data_conclusao}")
            elif status != "Concluída":
                # Se a ação não está concluída, limpar a data de conclusão
                df_acoes.at[idx, 'Data de Conclusão'] = None
                print("Limpando Data de Conclusão pois o status não é Concluída")
            
            # Atualizar observações apenas se fornecidas
            if observacoes:
                df_acoes.at[idx, 'Observações de conclusão'] = observacoes
            
            # DIAGNÓSTICO: Ver todos os valores da linha atualizada
            print("\n=== VALORES FINAIS DA AÇÃO ATUALIZADA ===")
            for coluna in df_acoes.columns:
                valor = df_acoes.at[idx, coluna]
                print(f"{coluna}: {valor} (tipo: {type(valor)})")
            
            # VERIFICAÇÃO ADICIONAL: Verificar se a Data Limite está correta
            if 'Data Limite' in df_acoes.columns:
                print(f"Valor final da Data Limite: {df_acoes.at[idx, 'Data Limite']} (tipo: {type(df_acoes.at[idx, 'Data Limite'])})")
                
                        # Verificação final dos campos críticos após processamento
            campos_criticos = ['Projeto', 'Descrição da Ação', 'Responsáveis', 'Data Limite', 'Status']
            for campo in campos_criticos:
                if campo in df_acoes.columns:
                    valor = df_acoes.at[idx, campo]
                    print(f"VERIFICAÇÃO FINAL - {campo}: {valor} (tipo: {type(valor)})")
            
            # Recalcular campos derivados
            if 'Status' in df_acoes.columns and 'Data Limite' in df_acoes.columns:
                # Calcular dias restantes
                df_acoes = process_acoes(df_acoes)
        
        # Atualizar o Google Sheets
        success = update_acoes_in_sheets(df_acoes)
        
        if success:
            return False, False, "", df_acoes.to_dict('records')
        else:
            return True, True, "Erro ao atualizar a planilha. Tente novamente.", dash.no_update
    else:
        # Criar DataFrame do zero com apenas essa ação
        nova_acao = {
            'ID da Ação': 1,
            'Data de Cadastro': datetime.now().strftime('%Y-%m-%d'),
            'Mês de Referência': mes_referencia,
            'Projeto': projeto,
            'Descrição da Ação': descricao,
            'Responsáveis': ', '.join(responsaveis) if isinstance(responsaveis, list) else responsaveis,
            'Data Limite': data_limite,  # Pode ser None ou vazio
            'Status': status,
            'Prioridade': prioridade,
            'Data de Conclusão': data_conclusao,
            'Observações de conclusão': ""
        }
        
        df_nova_acao = pd.DataFrame([nova_acao])
        df_nova_acao = process_acoes(df_nova_acao)
        
        # Atualizar o Google Sheets
        success = update_acoes_in_sheets(df_nova_acao)
        
        if success:
            return False, False, "", df_nova_acao.to_dict('records')
        else:
            return True, True, "Erro ao atualizar a planilha. Tente novamente.", dash.no_update

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
        print(
            f"INFO: A tabela tem {len(df)} linhas e a coluna 'Observacoes' tem {(df['Observacoes'] != '').sum()} valores não vazios.")

    return df.to_dict('records')

# Nova callback para o campo de busca
@app.callback(
    Output("projetos-table", "data", allow_duplicate=True),
    [Input("projetos-table-search", "value")],
    [State("raw-data-store", "data")],
    prevent_initial_call=True
)
def filter_table_by_search(search_term, data):
    if not search_term:
        # Se o campo de busca estiver vazio, usar os dados originais
        df = pd.DataFrame(data) if data else pd.DataFrame()
        if df.empty:
            return []
        df_table = process_data(df)
        return df_table.to_dict('records')
    
    # Filtra os dados com base no termo de busca
    df = pd.DataFrame(data) if data else pd.DataFrame()
    if df.empty:
        return []
    
    # Processa os dados para ter o mesmo formato da tabela
    df_table = process_data(df)
    
    # Converter termo de busca para minúsculas
    search_term = search_term.lower()
    
    # Filtrar linhas que contenham o termo de busca em qualquer coluna de texto
    filtered_data = []
    for row in df_table.to_dict('records'):
        row_values = str(row).lower()
        if search_term in row_values:
            filtered_data.append(row)
    
    return filtered_data

# Callback para abrir o modal de nova ação
@app.callback(
    Output("modal-nova-acao", "is_open"),
    Input("nova-acao-btn", "n_clicks"),
    prevent_initial_call=True
)
def abrir_modal_nova_acao(n_clicks):
    print("Callback abrir_modal_nova_acao foi chamado!")
    if n_clicks:
        return True
    return False

# Callback para abrir o modal de edição ao clicar na tabela de ações
@app.callback(
    [
        Output("modal-edicao-acao", "is_open"),
        Output("modal-edit-id", "value"),
        Output("modal-edit-projeto", "value"),
        Output("modal-edit-mes-referencia", "value"),
        Output("modal-edit-prioridade", "value"),
        Output("modal-edit-descricao", "value"),
        Output("modal-edit-responsaveis", "value"),
        Output("modal-edit-data-limite-display", "children"),
        Output("modal-edit-status", "value"),
        Output("modal-edit-data-conclusao", "date"),
        Output("modal-edit-observacoes", "value"),
    ],
    [Input("acoes-table", "active_cell"), Input("acoes-table", "derived_virtual_data")],
    [
        State("acoes-table", "data"),
        State("filter-options-store", "data")
    ],
    prevent_initial_call=True
)
def open_edit_acao_modal(active_cell, derived_data, table_data, filter_options):
    # Verificar se uma célula foi clicada e se temos dados na tabela
    if active_cell is None or not derived_data or not table_data:
        return False, "", "", "", "", "", [], None, "", None, ""
    
    # Identificar a linha clicada
    row_idx = active_cell["row"]
    if row_idx < 0 or row_idx >= len(derived_data):
        return False, "", "", "", "", "", [], None, "", None, ""
    
    # Obter os dados da ação
    row = derived_data[row_idx]
    print(f"\n===== Abrindo modal de edição para ação {row.get('ID da Ação', 'N/A')} =====")
    
    # Preparar data limite para o datepicker (formato ISO)
    data_limite = None
    if 'Data Limite' in row and row['Data Limite']:
        try:
            # Mais detalhes para diagnóstico
            data_limite_original = row['Data Limite']
            print(f"Data Limite original: {data_limite_original} (tipo: {type(data_limite_original)})")
            
            # Garantir que temos uma string para trabalhar
            data_limite_str = str(data_limite_original)
            
            # Análise detalhada do formato
            if isinstance(data_limite_original, str):
                # Verificar formato DD/MM/YYYY
                if '/' in data_limite_str:
                    partes = data_limite_str.split('/')
                    if len(partes) == 3:
                        # Converter de DD/MM/YYYY para YYYY-MM-DD
                        data_limite = f"{partes[2]}-{partes[1]}-{partes[0]}"
                        print(f"Data limite convertida de DD/MM/YYYY: {data_limite}")
                # Verificar formato YYYY-MM-DD
                elif '-' in data_limite_str:
                    partes = data_limite_str.split('-')
                    if len(partes) == 3 and len(partes[0]) == 4:
                        # Já está no formato esperado
                        data_limite = data_limite_str
                        print(f"Data limite já em formato ISO YYYY-MM-DD: {data_limite}")
                    else:
                        # Tentar usar pandas para converter
                        try:
                            data_obj = pd.to_datetime(data_limite_str)
                            data_limite = data_obj.strftime('%Y-%m-%d')
                            print(f"Data limite convertida via pandas (formato com '-'): {data_limite}")
                        except Exception as e:
                            # Usar como está
                            data_limite = data_limite_str
                            print(f"Usando data limite original: {data_limite}")
                else:
                    # Tentar usar pandas para converter
                    try:
                        data_obj = pd.to_datetime(data_limite_str)
                        data_limite = data_obj.strftime('%Y-%m-%d')
                        print(f"Data limite convertida via pandas: {data_limite}")
                    except Exception as e:
                        # Usar como está
                        data_limite = data_limite_str
                        print(f"Usando data limite original: {data_limite}")
            elif hasattr(data_limite_original, 'strftime'):
                # É um objeto datetime ou similar
                print("Data Limite é um objeto datetime")
                data_limite = data_limite_original.strftime('%Y-%m-%d')
                print(f"Data limite formatada de datetime: {data_limite}")
            else:
                print(f"Data Limite tem tipo desconhecido: {type(data_limite_original)}")
                # Tentar converter com pandas
                try:
                    data_obj = pd.to_datetime(data_limite_str)
                    data_limite = data_obj.strftime('%Y-%m-%d')
                    print(f"Data limite convertida via pandas (tipo desconhecido): {data_limite}")
                except Exception as e:
                    print(f"Erro ao converter data limite com pandas (tipo desconhecido): {e}")
                    # Usar a string bruta
                    data_limite = data_limite_str
                    print(f"Usando string bruta: {data_limite}")
            
            print(f"Data limite final para o modal: {data_limite}")
        except Exception as e:
            print(f"Erro ao processar data limite: {e}")
            print(f"Usando data original sem processamento: {row.get('Data Limite')}")
            data_limite = row.get('Data Limite', '')
    else:
        print("Não foi encontrada Data Limite no registro")
    
    # Preparar data de conclusão para o datepicker (formato ISO)
    data_conclusao = None
    if 'Data de Conclusão' in row and row['Data de Conclusão']:
        try:
            # Mais detalhes para diagnóstico
            data_conclusao_original = row['Data de Conclusão']
            print(f"Data de Conclusão original: {data_conclusao_original} (tipo: {type(data_conclusao_original)})")
            
            # Garantir que temos uma string para trabalhar
            data_conclusao_str = str(data_conclusao_original)
            
            # Análise detalhada do formato
            if isinstance(data_conclusao_original, str):
                # Verificar formato DD/MM/YYYY
                if '/' in data_conclusao_str:
                    partes = data_conclusao_str.split('/')
                    if len(partes) == 3:
                        # Converter de DD/MM/YYYY para YYYY-MM-DD
                        data_conclusao = f"{partes[2]}-{partes[1]}-{partes[0]}"
                        print(f"Data de conclusão convertida de DD/MM/YYYY: {data_conclusao}")
                # Verificar formato YYYY-MM-DD
                elif '-' in data_conclusao_str:
                    partes = data_conclusao_str.split('-')
                    if len(partes) == 3 and len(partes[0]) == 4:
                        # Já está no formato esperado
                        data_conclusao = data_conclusao_str
                        print(f"Data de conclusão já em formato ISO YYYY-MM-DD: {data_conclusao}")
                    else:
                        # Tentar usar pandas para converter
                        try:
                            data_obj = pd.to_datetime(data_conclusao_str)
                            data_conclusao = data_obj.strftime('%Y-%m-%d')
                            print(f"Data de conclusão convertida via pandas (formato com '-'): {data_conclusao}")
                        except Exception as e:
                            # Usar como está
                            data_conclusao = data_conclusao_str
                            print(f"Usando data de conclusão original: {data_conclusao}")
                else:
                    # Tentar usar pandas para converter
                    try:
                        data_obj = pd.to_datetime(data_conclusao_str)
                        data_conclusao = data_obj.strftime('%Y-%m-%d')
                        print(f"Data de conclusão convertida via pandas: {data_conclusao}")
                    except Exception as e:
                        # Usar como está
                        data_conclusao = data_conclusao_str
                        print(f"Usando data de conclusão original: {data_conclusao}")
            elif hasattr(data_conclusao_original, 'strftime'):
                # É um objeto datetime ou similar
                data_conclusao = data_conclusao_original.strftime('%Y-%m-%d')
                print(f"Data de conclusão formatada de datetime: {data_conclusao}")
            else:
                # Tentar converter com pandas
                try:
                    data_obj = pd.to_datetime(data_conclusao_str)
                    data_conclusao = data_obj.strftime('%Y-%m-%d')
                    print(f"Data de conclusão convertida via pandas (tipo desconhecido): {data_conclusao}")
                except Exception as e:
                    # Usar a string bruta
                    data_conclusao = data_conclusao_str
                    print(f"Usando string bruta para data de conclusão: {data_conclusao}")
            
            print(f"Data de conclusão final para o modal: {data_conclusao}")
        except Exception as e:
            print(f"Erro ao processar data de conclusão: {e}")
            data_conclusao = row.get('Data de Conclusão', '')
    else:
        print("Não foi encontrada Data de Conclusão no registro")
    
    # Preparar responsáveis (pode ser string ou lista)
    responsaveis = row.get('Responsáveis', '')
    if isinstance(responsaveis, str) and ',' in responsaveis:
        responsaveis = [resp.strip() for resp in responsaveis.split(',') if resp.strip()]
    elif isinstance(responsaveis, str):
        responsaveis = [responsaveis] if responsaveis.strip() else []
    
    mes_referencia = row.get('Mês de Referência', '')
    print(f"Enviando mês de referência para o modal: '{mes_referencia}'")
    print(f"Enviando data limite para o modal: '{data_limite}'")
    
    # Preencher o modal com os dados da ação
    return (
        True,  # Abrir o modal
        row.get('ID da Ação', ''),
        row.get('Projeto', ''),
        mes_referencia,
        row.get('Prioridade', 'Média'),
        row.get('Descrição da Ação', ''),
        responsaveis,
        str(data_limite) if data_limite else "Não definida",  # Agora usamos children em vez de date
        row.get('Status', 'Pendente'),
        data_conclusao,  # Pode ser None ou string em formato YYYY-MM-DD
        row.get('Observações de conclusão', '')
    )

# Callback para preencher as opções do dropdown de mês de referência no modal de edição de ação
@app.callback(
    [
        Output("modal-edit-mes-referencia", "options"),
        Output("modal-edit-mes-referencia", "value", allow_duplicate=True)
    ],
    [
        Input("modal-edicao-acao", "is_open"),
        Input("modal-edit-mes-referencia", "value")
    ],
    [
        State("filter-options-store", "data"),
        State("acoes-store", "data"),
        State("modal-edit-id", "value")
    ],
    prevent_initial_call=True
)
def update_edit_mes_referencia_options(is_open, atual_value, filter_options_data, acoes_data, acao_id):
    ctx = dash.callback_context
    trigger = ctx.triggered[0]['prop_id'].split('.')[0] if ctx.triggered else None
    
    if not is_open or not filter_options_data:
        return [], dash.no_update
    
    # Primeiro, adicionar as opções de meses do ano (Janeiro a Dezembro) com nome completo
    meses_opcoes = [
        {"label": "Janeiro", "value": "Janeiro"},
        {"label": "Fevereiro", "value": "Fevereiro"},
        {"label": "Março", "value": "Março"},
        {"label": "Abril", "value": "Abril"},
        {"label": "Maio", "value": "Maio"},
        {"label": "Junho", "value": "Junho"},
        {"label": "Julho", "value": "Julho"},
        {"label": "Agosto", "value": "Agosto"},
        {"label": "Setembro", "value": "Setembro"},
        {"label": "Outubro", "value": "Outubro"},
        {"label": "Novembro", "value": "Novembro"},
        {"label": "Dezembro", "value": "Dezembro"}
    ]
    
    # Dicionário para converter abreviações para nomes completos
    meses_map = {
        "Jan/2023": "Janeiro",
        "Fev/2023": "Fevereiro",
        "Mar/2023": "Março",
        "Abr/2023": "Abril",
        "Mai/2023": "Maio",
        "Jun/2023": "Junho",
        "Jul/2023": "Julho",
        "Ago/2023": "Agosto",
        "Set/2023": "Setembro",
        "Out/2023": "Outubro",
        "Nov/2023": "Novembro",
        "Dez/2023": "Dezembro"
    }
    
    # Adicionar também as opções existentes nos dados
    if 'meses_anos' in filter_options_data:
        for mes in filter_options_data['meses_anos']:
            # Verificar se é uma abreviação que podemos converter
            nome_completo = meses_map.get(mes, mes)
            
            # Verificar se o mês já não está nas opções
            if not any(op["value"] == nome_completo for op in meses_opcoes):
                meses_opcoes.append({"label": nome_completo, "value": nome_completo})
    
    # Se o callback foi disparado pela abertura do modal
    if trigger == "modal-edicao-acao" and acao_id and acoes_data:
        # Encontrar o mês de referência da ação atual
        df_acoes = pd.DataFrame(acoes_data)
        acao = df_acoes[df_acoes['ID da Ação'].astype(str) == str(acao_id)]
        
        if not acao.empty and 'Mês de Referência' in acao.columns:
            mes_referencia = acao['Mês de Referência'].iloc[0]
            print(f"Mês de referência encontrado para ação ID {acao_id}: '{mes_referencia}'")
            
            # Converter para nome completo se necessário
            nome_completo = meses_map.get(mes_referencia, mes_referencia)
            
            # Verificar se o mês existe nas opções
            if any(op["value"] == nome_completo for op in meses_opcoes):
                return meses_opcoes, nome_completo
            elif nome_completo:
                # Se o mês não está nas opções mas é válido, adicionar
                meses_opcoes.append({"label": nome_completo, "value": nome_completo})
                return meses_opcoes, nome_completo
    
    # Também converter o valor atual para nome completo se necessário
    if atual_value:
        atual_value_nome_completo = meses_map.get(atual_value, atual_value)
        # Retornar apenas as opções e manter o valor atual
        return meses_opcoes, atual_value_nome_completo
    
    # Retornar apenas as opções e manter o valor atual se já existir
    return meses_opcoes, dash.no_update

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
            print(
                "   Verifique o arquivo de credenciais e as permissões da conta de serviço.")

        # Em produção (como no Render), o app será executado pelo Gunicorn
        # Em desenvolvimento local, usamos o servidor integrado do Dash
        import os
        debug = os.environ.get('ENV', 'development') == 'development'
        port = int(os.environ.get('PORT', 8050))  # Porta padrão 8050 para Dash

        print(
            f"\nIniciando servidor Dash {'em modo debug' if debug else 'em produção'} na porta {port}")
        app.run(debug=True)
    except Exception as e:
        print(f"❌ ERRO CRÍTICO ao iniciar o aplicativo: {e}")
        import traceback
        traceback.print_exc()
