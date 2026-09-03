import os
import requests
from flask import Flask, jsonify, request
import openpyxl
from openpyxl.styles import PatternFill

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

@app.route('/api/lancar-noite', methods=['POST'])
def lancar_noite():
    dados = request.json
    # Espera receber: {"linha": 6, "valor": 150.0, "tipo": "PIX"} ou {"linha": 6, "tipo": "ATRASO", "status": "AA"}
    try:
        baixar_planilha_do_drive()
        wb = openpyxl.load_workbook(EXCEL_PATH)
        sheet = wb.active
        
        linha = dados.get('linha')
        tipo = dados.get('tipo') # "PIX", "ESPECIE", "ATRASO"
        valor = dados.get('valor', 0)
        
        # Cores do Mapa
        CORES = {
            "PIX": PatternFill(start_color="6AA84F", end_color="6AA84F", fill_type="solid"),      # Verde
            "ESPECIE": PatternFill(start_color="8E7CC3", end_color="8E7CC3", fill_type="solid"),  # Roxo
            "ATRASO": PatternFill(start_color="CC0000", end_color="CC0000", fill_type="solid")    # Vermelho
        }
        
        # Identifica dinamicamente as colunas do dia (penúltima para Aberto, última para OUT.VLR)
        col_out_vlr = sheet.max_column
        col_aberto = col_out_vlr - 1
        
        if tipo in ["PIX", "ESPECIE"]:
            # Insere o valor recebido na coluna OUT.VLR
            sheet.cell(row=linha, column=col_out_vlr, value=valor)
            
            # Aplica a cor correspondente no Valor em Aberto e no OUT.VLR
            fill = CORES[tipo]
            sheet.cell(row=linha, column=col_aberto).fill = fill
            sheet.cell(row=linha, column=col_out_vlr).fill = fill
            
            # Verifica se quitou o contrato (Coluna K / Pendente zerado)
            pendente = sheet.cell(row=linha, column=11).value
            if pendente is not None and float(pendente) <= 0:
                from datetime import datetime
                # Carimba a data de quitação na Coluna D (Quitação)
                sheet.cell(row=linha, column=4, value=datetime.now().strftime("%d/%m/%Y"))
                
        elif tipo == "ATRASO":
            status = dados.get('status', 'A') # A, AA, AAA, etc.
            sheet.cell(row=linha, column=col_out_vlr, value=status)
            
            # Pinta de vermelho o atraso
            fill = CORES["ATRASO"]
            sheet.cell(row=linha, column=col_aberto).fill = fill
            sheet.cell(row=linha, column=col_out_vlr).fill = fill

        wb.save(EXCEL_PATH)
        return jsonify({"mensagem": "Lançamento noturno realizado com sucesso!"})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500