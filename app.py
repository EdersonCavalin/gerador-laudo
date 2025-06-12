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
            if f"{{{{{chave}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{chave}}}}}", valor)

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
        dados = {campo: request.form.get(campo, '') for campo in [
            'nroOS', 'DATA', 'Nome_Tecnico', 'Nome_Cliente', 'Endereco_Cliente',
            'Telefone_cliente', 'Modelo_equipamento', 'Numero_Serie',
            'Chamado_Aberto', 'Defeitos_Encontrados', 'Tarefas_Executadas']}

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
    import os
port = int(os.environ.get('PORT', 5000))
app.run(host='0.0.0.0', port=port)
