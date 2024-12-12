import os
from dotenv import load_dotenv
import pandas as pd
import gspread
from werkzeug.utils import secure_filename
from flask import Flask, request, render_template
from google.oauth2.service_account import Credentials

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()

app = Flask(__name__)

# Diretório onde o upload será feito
UPLOAD_FOLDER = os.path.abspath('uploads')

# Criar diretório de uploads, caso não exista
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Extensões permitidas
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Função para verificar as extensões permitidas
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Função para carregar e filtrar a base de dados
def carregar_base(file_path, colunas_desejadas):
    base = pd.read_excel(file_path)
    
    # Filtra as colunas desejadas
    base_filtrada = base.reindex(columns=colunas_desejadas)
    
    # Substitui NaN, inf, -inf por None
    def clean_value(x):
        # Substituir valores inválidos por None
        if pd.isna(x) or x == float('inf') or x == float('-inf'):
            return None
        return x

    # Aplicar a função de limpeza a todas as células
    base_filtrada = base_filtrada.applymap(clean_value)
    
    # Verifique os dados antes de enviá-los ao Google Sheets
    print(base_filtrada.head())  # Exibe as 5 primeiras linhas para inspeção
    
    return base_filtrada

# Função para autenticar e acessar o Google Sheets
def acessar_google_sheets():
    # Carregar o caminho do arquivo de credenciais e o ID da planilha do arquivo .env
    credenciais_json = os.getenv("GOOGLE_CREDENTIALS_PATH")
    creds = Credentials.from_service_account_file(
        credenciais_json, 
        scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    )
    client = gspread.authorize(creds)
    return client

# Função para escrever os dados no Google Sheets
def escrever_no_sheets(sheet, dados):
    for i, row in enumerate(dados, start=2):  # Começar da linha 2 para não sobrescrever o cabeçalho
        sheet.append_row(row)

# Rota para exibir o formulário de upload
@app.route('/')
def index():
    return render_template('index.html')  # Página HTML de upload

# Função para salvar o arquivo enviado
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part', 400  # Se o arquivo não for enviado

    file = request.files['file']
    
    if file.filename == '':
        return 'No selected file', 400  # Se o arquivo não for selecionado
    
    if file and allowed_file(file.filename):  # Verificar a extensão permitida
        filename = secure_filename(file.filename)  # Sanitize filename
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)  # Caminho para salvar o arquivo
        
        try:
            # Salvar o arquivo
            file.save(file_path)
            
            # Defina as colunas que você deseja extrair
            colunas_desejadas = ['CLIENTE', 'ANO', 'p4', 'p6', 'p13', 'p14', 'p54', 'p164']
            
            # Carregar e filtrar a base de dados
            base_filtrada = carregar_base(file_path, colunas_desejadas)
            
            # Acessar o Google Sheets
            client = acessar_google_sheets()
            spreadsheet_id = os.getenv("SPREADSHEET_ID")  # Usando a variável de ambiente
            sheet = client.open_by_key(spreadsheet_id).sheet1  # Seleciona a primeira aba
            
            # Escrever os dados no Google Sheets
            dados = base_filtrada.values.tolist()  # Converte os dados para lista de listas
            escrever_no_sheets(sheet, dados)
            
            return f'Arquivo {filename} carregado com sucesso e dados enviados ao Google Sheets!', 200
        except Exception as e:
            return f'Ocorreu um erro: {e}', 500  # Retorna erro caso haja falha
    else:
        return 'Arquivo não permitido', 400

if __name__ == '__main__':
    app.run(debug=True)
