from flask import Flask, render_template, request, redirect, url_for
import pandas as pd

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('form_chamado.html')

@app.route('/criar_chamado', methods=['POST'])
def criar_chamado():
    try:
        data = request.form['data']
        horario = request.form['horario']
        numero_chamado = request.form['numero_chamado']
        endereco = request.form['endereco']
        numero = request.form['numero']
        cep = request.form['cep']
        cidade = request.form['cidade']
        nome_empresa = request.form['nome_empresa']
        contato_empresa = request.form['contato_empresa']
        responsavel_3am = request.form['responsavel_3am']
        defeito = request.form['defeito']
        servico = request.form['servico']
        modelo = request.form['modelo']
        numero_serie = request.form['numero_serie']
        status = request.form['status']
        observacoes = request.form.get('observacoes', '')  # Campo opcional
    except KeyError as e:
        return f"Missing form field: {e}", 400

    # Carregar dados do arquivo Excel existente
    excel_file = 'agenda_chamados.xlsx'
    try:
        df = pd.read_excel(excel_file, engine='openpyxl')
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            'data', 'horario', 'numero_chamado', 'endereco', 'numero', 'cep', 'cidade', 
            'nome_empresa', 'contato_empresa', 'responsavel_3am', 'defeito', 'servico', 
            'modelo', 'numero_serie', 'status', 'observacoes'
        ])

    # Adicionar novo chamado
    novo_chamado = pd.DataFrame([{
        'data': data,
        'horario': horario,
        'numero_chamado': numero_chamado,
        'endereco': endereco,
        'numero': numero,
        'cep': cep,
        'cidade': cidade,
        'nome_empresa': nome_empresa,
        'contato_empresa': contato_empresa,
        'responsavel_3am': responsavel_3am,
        'defeito': defeito,
        'servico': servico,
        'modelo': modelo,
        'numero_serie': numero_serie,
        'status': status,
        'observacoes': observacoes
    }])
    
    df = pd.concat([df, novo_chamado], ignore_index=True)

    # Salvar os dados atualizados de volta no arquivo Excel
    df.to_excel(excel_file, index=False, engine='openpyxl')

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
