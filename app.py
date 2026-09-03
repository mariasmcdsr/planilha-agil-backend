import os
import requests
from flask import Flask, jsonify, request, render_template
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime

app = Flask(__name__)

DRIVE_FILE_ID = "14_9A0gjPBokDpdbw4wYGTepTdQKjbe8u"
EXCEL_PATH = "planilha_atual.xlsx"

def garantir_planilha():
    # Se o arquivo não existir localmente, baixa uma única vez para ser instantâneo
    if not os.path.exists(EXCEL_PATH):
        baixar_planilha_do_drive()

def baixar_planilha_do_drive():
    url = f"https://drive.google.com/uc?export=download&id={DRIVE_FILE_ID}"
    session = requests.Session()
    response = session.get(url, stream=True)
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            params = {'export': 'download', 'id': DRIVE_FILE_ID, 'confirm': value}
            response = session.get(url, params=params, stream=True)
            break
    if response.status_code == 200:
        with open(EXCEL_PATH, 'wb') as f:
            for chunk in response.iter_content(chunk_size=1024):
                if chunk:
                    f.write(chunk)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/buscar-cliente', methods=['GET'])
def buscar_cliente():
    nome_busca = request.args.get('nome', '').strip()
    if not nome_busca:
        return jsonify({"erro": "Informe o nome do cliente"}), 400
    
    garantir_planilha()
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    sheet = wb.active
    
    for row in range(5, sheet.max_row + 1):
        nome = sheet.cell(row=row, column=2).value # Coluna B (Nome)
        if nome and nome_busca.lower() in str(nome).lower():
            dados_cliente = {
                "linha": row,
                "nome": nome,
                "contrato": str(sheet.cell(row=row, column=3).value or ""), # Coluna C
                "quitacao": str(sheet.cell(row=row, column=4).value or ""), # Coluna D
                "valor_pego": sheet.cell(row=row, column=7).value or 0,    # Exemplo Coluna G
                "valor_pago": sheet.cell(row=row, column=9).value or 0,    # Exemplo Coluna I
                "pendente": sheet.cell(row=row, column=11).value or 0,     # Coluna K (Pendente)
            }
            return jsonify(dados_cliente)
            
    return jsonify({"erro": "Cliente não encontrado!"}), 404

@app.route('/api/lancar-noite', methods=['POST'])
def lancar_noite():
    dados = request.json
    try:
        garantir_planilha()
        wb = openpyxl.load_workbook(EXCEL_PATH)
        sheet = wb.active
        
        linha = dados.get('linha')
        tipo = dados.get('tipo') # "PIX", "ESPECIE", "ATRASO"
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
                sheet.cell(row=linha, column=4, value=datetime.now().strftime("%d/%m/%Y"))
                
        elif tipo == "ATRASO":
            status = dados.get('status', 'A') # A, AA, AAA, etc.
            sheet.cell(row=linha, column=col_out_vlr, value=status)
            fill = CORES["ATRASO"]
            sheet.cell(row=linha, column=col_aberto).fill = fill
            sheet.cell(row=linha, column=col_out_vlr).fill = fill

        wb.save(EXCEL_PATH)
        return jsonify({"mensagem": "Lançamento realizado instantaneamente!"})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500

@app.route('/api/lancar-manha', methods=['POST'])
def lancar_manha():
    dados = request.json
    try:
        baixar_planilha_do_drive() # Na manhã atualizamos baixando a matriz fresca do Drive
        wb = openpyxl.load_workbook(EXCEL_PATH)
        
        aba_ativa = wb.active
        nova_data = dados.get('nova_data', datetime.now().strftime("%d/%m"))
        nova_aba = wb.copy_worksheet(aba_ativa)
        nova_aba.title = f"CAIXA 07 BSB {nova_data}"
        
        COR_QUITADO = PatternFill(start_color="FFE599", end_color="FFE599", fill_type="solid")
        
        # Move quitados para baixo e pinta de amarelo
        for row in range(5, nova_aba.max_row + 1):
            pendente = nova_aba.cell(row=row, column=11).value
            if pendente is not None and float(pendente) == 0:
                for col in range(2, 15):
                    nova_aba.cell(row=row, column=col).fill = COR_QUITADO
                    
        wb.save(EXCEL_PATH)
        return jsonify({"mensagem": "Rotina da manhã executada com sucesso!"})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))