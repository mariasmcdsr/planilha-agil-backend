import os
import requests
from flask import Flask, jsonify, request
import openpyxl

app = Flask(__name__)

# ID da planilha no Google Drive fornecido
DRIVE_FILE_ID = "1927744103"
EXCEL_PATH = "planilha_atual.xlsx"

def baixar_planilha_do_drive():
    url = f"https://drive.google.com/uc?export=download&id={DRIVE_FILE_ID}"
    response = requests.get(url, stream=True)
    if response.status_code == 200:
        with open(EXCEL_PATH, 'wb') as f:
            for chunk in response.iter_content(chunk_size=1024):
                if chunk:
                    f.write(chunk)

@app.route('/')
def home():
    return jsonify({"status": "PlanilhaÁgil Backend Online e Conectado!"})

@app.route('/api/lancar', methods=['POST'])
def lancar_pagamento():
    dados = request.json
    # Baixa a versão mais recente da planilha do Drive
    try:
        baixar_planilha_do_drive()
        wb = openpyxl.load_workbook(EXCEL_PATH)
        sheet = wb.active
        
        # Aqui entra a lógica de escrita na coluna OUT.VLR e aplicação das cores
        # (Será detalhada no próximo passo com base nas colunas exatas)
        
        wb.save(EXCEL_PATH)
        return jsonify({"mensagem": "Pagamento registrado e planilha atualizada com sucesso!"})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))