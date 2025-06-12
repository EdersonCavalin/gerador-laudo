from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
import os
from datetime import datetime

app = Flask(__name__)
UPLOAD_FOLDER = 'static/uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def preencher_docx(dados, fotos):
    tpl = DocxTemplate("MODELO_LAUDO.docx")

    # Converte cada foto para InlineImage (mantém formatação)
    for i in range(1, 5):
        if f"foto{i}_antes" in fotos:
            dados[f"foto{i}_antes"] = InlineImage(tpl, fotos[f"foto{i}_antes"], width=Inches(2))
        if f"foto{i}_depois" in fotos:
            dados[f"foto{i}_depois"] = InlineImage(tpl, fotos[f"foto{i}_depois"], width=Inches(2))

    tpl.render(dados)
    nome = f"laudo_{dados['nroOS']}_{datetime.now():%Y%m%d%H%M%S}.docx"
    caminho = os.path.join(app.config['UPLOAD_FOLDER'], nome)
    tpl.save(caminho)
    return caminho

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        dados = {
            "nroOS": request.form.get("nroOS", ""),
            "DATA": request.form.get("DATA", ""),
            "Nome_Tecnico": request.form.get("Nome_Tecnico", ""),
            "Nome_Cliente": request.form.get("Nome_Cliente", ""),
            "Endereco_Cliente": request.form.get("Endereco_Cliente", ""),
            "Telefone_cliente": request.form.get("Telefone_cliente", ""),
            "Modelo_equipamento": request.form.get("Modelo_equipamento", ""),
            "Numero_Serie": request.form.get("Numero_Serie", ""),
            "Chamado_Aberto": request.form.get("Chamado_Aberto", ""),
            "Defeitos_Encontrados": request.form.get("Defeitos_Encontrados", ""),
            "Tarefas_Executadas": request.form.get("Tarefas_Executadas", ""),
        }

        fotos = {}
        for i in range(1, 5):
            for tipo in ("antes", "depois"):
                file = request.files.get(f"foto{i}_{tipo}")
                if file and file.filename:
                    caminho = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
                    file.save(caminho)
                    fotos[f"foto{i}_{tipo}"] = caminho

        caminho_docx = preencher_docx(dados, fotos)
        return send_file(caminho_docx, as_attachment=True)

    return render_template("formulario.html")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
