from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Inches
import os
from datetime import datetime

app = Flask(__name__)
UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def preencher_docx(dados, fotos):
    doc = Document("MODELO_LAUDO.docx")
    for p in doc.paragraphs:
    for chave, valor in dados.items():
        alvo = f"{{{{{chave}}}}}"
        if alvo in p.text:
            for run in p.runs:
                if alvo in run.text:
                    run.text = run.text.replace(alvo, valor)
    tabelas = doc.tables
    if tabelas:
        tabela = tabelas[-1]
        for i in range(4):
            if fotos.get(f"foto{i+1}_antes"):
                tabela.cell(i+1, 0).paragraphs[0].add_run().add_picture(fotos[f"foto{i+1}_antes"], width=Inches(2))
            if fotos.get(f"foto{i+1}_depois"):
                tabela.cell(i+1, 1).paragraphs[0].add_run().add_picture(fotos[f"foto{i+1}_depois"], width=Inches(2))

    nome_arquivo = f"laudo_{dados['nroOS']}_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
    caminho_saida = os.path.join(app.config['UPLOAD_FOLDER'], nome_arquivo)
    doc.save(caminho_saida)
    return caminho_saida

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        dados = {
    'nroOS': request.form.get('nroOS', ''),
    'DATA': request.form.get('DATA', ''),
    'Nome Técnico': request.form.get('Nome_Tecnico', ''),
    'Nome Cliente': request.form.get('Nome_Cliente', ''),
    'Endereço Cliente': request.form.get('Endereco_Cliente', ''),
    'Telefone_cliente': request.form.get('Telefone_cliente', ''),
    'Modelo_equipamento': request.form.get('Modelo_equipamento', ''),
    'Numero Serie': request.form.get('Numero_Serie', ''),
    'Chamado Aberto': request.form.get('Chamado_Aberto', ''),
    'Defeitos Encontrados': request.form.get('Defeitos_Encontrados', ''),
    'Tarefas Executadas': request.form.get('Tarefas_Executadas', ''),
}

        fotos = {}
        for i in range(1, 5):
            for tipo in ['antes', 'depois']:
                file = request.files.get(f"foto{i}_{tipo}")
                if file and file.filename:
                    caminho_foto = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
                    file.save(caminho_foto)
                    fotos[f"foto{i}_{tipo}"] = caminho_foto

        docx_path = preencher_docx(dados, fotos)
        return send_file(docx_path, as_attachment=True)

    return render_template('formulario.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

