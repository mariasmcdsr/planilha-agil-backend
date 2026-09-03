import os
import requests
from flask import Flask, jsonify, request, render_template
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
    return render_template('index.html')

@app.route('/api/lancar', methods=['POST'])
def lancar_pagamento():
    dados = request.json
    try:
        baixar_planilha_do_drive()
        wb = openpyxl.load_workbook(EXCEL_PATH)
        sheet = wb.active
        wb.save(EXCEL_PATH)
        return jsonify({"mensagem": "Pagamento registrado e planilha atualizada com sucesso!"})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500

@app.route('/api/lancar-noite', methods=['POST'])
def lancar_noite():
    dados = request.json
    try:
        baixar_planilha_do_drive()
        wb = openpyxl.load_workbook(EXCEL_PATH)
        sheet = wb.active
        
        linha = dados.get('linha')
        tipo = dados.get('tipo') 
        valor = dados.get('valor', 0)
        
        CORES = {
            "PIX": PatternFill(start_color="6AA84F", end_color="6AA84F", fill_type="solid"),
            "ESPECIE": PatternFill(start_color="8E7CC3", end_color="8E7CC3", fill_type="solid"),
            "ATRASO": PatternFill(start_color="CC0000", end_color="CC0000", fill_type="solid")
        }
        
        col_out_vlr = sheet.max_column
        col_aberto = col_out_vlr - 1
        
        if tipo in ["PIX", "ESPECIE"]:
            sheet.cell(row=linha, column=col_out_vlr, value=valor)
            fill = CORES[tipo]
            sheet.cell(row=linha, column=col_aberto).fill = fill
            sheet.cell(row=linha, column=col_out_vlr).fill = fill
            
            pendente = sheet.cell(row=linha, column=11).value
            if pendente is not None and float(pendente) <= 0:
                from datetime import datetime
                sheet.cell(row=linha, column=4, value=datetime.now().strftime("%d/%m/%Y"))
                
        elif tipo == "ATRASO":
            status = dados.get('status', 'A')
            sheet.cell(row=linha, column=col_out_vlr, value=status)
            fill = CORES["ATRASO"]
            sheet.cell(row=linha, column=col_aberto).fill = fill
            sheet.cell(row=linha, column=col_out_vlr).fill = fill

        wb.save(EXCEL_PATH)
        return jsonify({"mensagem": "Lançamento noturno realizado com sucesso!"})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500

@app.route('/api/lancar-manha', methods=['POST'])
def lancar_manha():
    dados = request.json
    try:
        baixar_planilha_do_drive()
        wb = openpyxl.load_workbook(EXCEL_PATH)
        
        aba_ativa = wb.active
        nova_data = dados.get('nova_data', '04/09/2026')
        nova_aba = wb.copy_worksheet(aba_ativa)
        nova_aba.title = f"CAIXA 07 BSB {nova_data}"
        
        COR_QUITADO = PatternFill(start_color="FFE599", end_color="FFE599", fill_type="solid")
        COR_NOVO = PatternFill(start_color="1155CC", end_color="1155CC", fill_type="solid")
        
        for row in range(5, nova_aba.max_row + 1):
            pendente = nova_aba.cell(row=row, column=11).value
            if pendente is not None and float(pendente) == 0:
                for col in range(2, 15):
                    nova_aba.cell(row=row, column=col).fill = COR_QUITADO
                    
        novo_cliente = dados.get('novo_cliente')
        if novo_cliente:
            linha_nova = novo_cliente.get('linha')
            nova_aba.cell(row=linha_nova, column=2, value=novo_cliente.get('nome'))
            nova_aba.cell(row=linha_nova, column=2).fill = COR_NOVO

        wb.save(EXCEL_PATH)
        return jsonify({"mensagem": "Rotina da manhã executada com sucesso!"})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))