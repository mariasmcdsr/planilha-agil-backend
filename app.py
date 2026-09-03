import os
from flask import Flask, jsonify, request
import openpyxl

app = Flask(__name__)

@app.route('/')
def home():
    return jsonify({"status": "PlanilhaÁgil Backend Online!"})

@app.route('/api/lancar', methods=['POST'])
def lancar_pagamento():
    dados = request.json
    # Lógica de manipulação da planilha openpyxl
    return jsonify({"mensagem": "Pagamento registrado com sucesso!"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))