# Status Mensal Codeart

Dashboard para visualização e gerenciamento de projetos e ações da Codeart Solutions.

## Características

- Painel de visualização de projetos com métricas e gráficos
- Gerenciamento de ações com atribuição de responsáveis
- Integração com o Google Sheets como banco de dados
- Filtros avançados para análise de dados
- Interface moderna e responsiva

## Requisitos

- Python 3.8+
- Pip (gerenciador de pacotes Python)
- Acesso a uma planilha no Google Sheets

## Dependências

- dash
- plotly
- pandas
- dash-bootstrap-components
- gspread
- oauth2client
- python-dotenv

## Configuração da integração com o Google Sheets

Para integrar o aplicativo com o Google Sheets, siga os passos abaixo:

1. Acesse o [Google Cloud Console](https://console.cloud.google.com)
2. Crie um novo projeto (ou selecione um existente)
3. Ative a API do Google Sheets e Google Drive para o projeto
4. Crie uma conta de serviço:
   - Menu lateral > IAM e administrador > Contas de serviço
   - Clique em "Criar conta de serviço"
   - Adicione um nome e descrição
   - Conceda o papel "Editor" para a conta
   - Clique em "Criar chave" e selecione o formato JSON
   - Faça o download do arquivo de credenciais

5. Prepare o arquivo de credenciais:
   - Crie uma pasta chamada `credentials` na raiz do projeto
   - Renomeie o arquivo de credenciais baixado para `google-credentials.json`
   - Copie o arquivo para a pasta `credentials`

6. Compartilhe sua planilha do Google Sheets com o e-mail da conta de serviço (disponível no arquivo de credenciais)

## Estrutura da planilha no Google Sheets

O aplicativo espera uma planilha chamada "Revisao Projetos - Geral" com as seguintes abas:

1. **Projetos** - Contendo os dados dos projetos com as colunas:
   - Mês
   - Projeto
   - GP Responsável
   - Status
   - Segmento
   - Tipo
   - Coordenação
   - Financeiro
   - Previsão
   - Real
   - Saldo Acumulado
   - Atraso em dias
   - NPS
   - Observacoes
   - Decisões

2. **Codenautas** - Contendo a lista de codenautas com as colunas:
   - Nome
   - Email
   - Cargo
   - Equipe

3. **Ações** - Contendo as ações com as colunas:
   - ID da Ação
   - Data de Cadastro
   - Mês de Referência
   - Projeto
   - Descrição da Ação
   - Responsáveis
   - Data Limite
   - Status
   - Prioridade
   - Data de Conclusão
   - Observações de conclusão

## Como executar

1. Clone o repositório
2. Instale as dependências:
```
pip install -r requirements.txt
```
3. Configure o acesso ao Google Sheets conforme instruções acima
4. Execute o aplicativo:
```
python app.py
```
5. Acesse o dashboard no navegador: http://127.0.0.1:8050/

## Observações importantes

- Certifique-se de que a conta de serviço tem acesso à planilha compartilhada
- Verifique se as colunas na planilha correspondem exatamente às esperadas pelo aplicativo
- Se encontrar problemas com a conexão, verifique os logs de erro no console 