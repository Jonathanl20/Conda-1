from flask import Flask, render_template, request, send_file
import os
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.enums import TA_JUSTIFY
from werkzeug.utils import secure_filename
import atexit
import shutil

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max-limit

# Criar pasta de uploads se não existir
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# Lista para armazenar arquivos temporários
temp_files = []

def cleanup_temp_files():
    """Função para limpar arquivos temporários ao fechar o servidor"""
    for file_path in temp_files:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except:
            pass

# Registrar função de limpeza
atexit.register(cleanup_temp_files)

def convert_docx_to_pdf(docx_path, pdf_path):
    # Ler o documento Word
    doc = Document(docx_path)
    
    # Criar o PDF
    pdf = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=72
    )
    
    # Estilos
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name='Justify',
        alignment=TA_JUSTIFY,
        fontName='Helvetica',
        fontSize=12,
        leading=14
    ))
    
    # Conteúdo do PDF
    story = []
    
    # Processar cada parágrafo do documento
    for para in doc.paragraphs:
        if para.text.strip():  # Ignorar parágrafos vazios
            p = Paragraph(para.text, styles['Justify'])
            story.append(p)
            story.append(Spacer(1, 12))
    
    # Gerar o PDF
    pdf.build(story)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return 'Nenhum arquivo enviado', 400
    
    file = request.files['file']
    if file.filename == '':
        return 'Nenhum arquivo selecionado', 400
    
    if not file.filename.endswith('.docx'):
        return 'Por favor, envie um arquivo .docx', 400
    
    try:
        # Salvar o arquivo Word
        filename = secure_filename(file.filename)
        word_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(word_path)
        temp_files.append(word_path)
        
        # Converter para PDF
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename.replace('.docx', '.pdf'))
        convert_docx_to_pdf(word_path, pdf_path)
        temp_files.append(pdf_path)
        
        # Enviar o arquivo PDF
        return send_file(
            pdf_path,
            as_attachment=True,
            download_name=filename.replace('.docx', '.pdf'),
            mimetype='application/pdf'
        )
    
    except Exception as e:
        # Limpar arquivos em caso de erro
        for file_path in [word_path, pdf_path]:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except:
                pass
        return f'Erro na conversão: {str(e)}', 500

if __name__ == '__main__':
    app.run(debug=True) 