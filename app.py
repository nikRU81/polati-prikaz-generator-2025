from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
import os

app = Flask(__name__)

# Путь к логотипу
LOGO_PATH = os.path.join(os.path.dirname(__file__), 'static', 'img', 'logo.png')

@app.route('/')
def index():
    """Главная страница с формой"""
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_prikaz():
    """Генерация приказа на основе данных формы"""
    try:
        data = request.get_json()
        
        # Валидация данных
        if not all(k in data for k in ['day', 'month', 'year', 'orderNumber', 'orderTitle', 'preamble', 'punkts']):
            return jsonify({'error': 'Не все обязательные поля заполнены'}), 400
        
        if not data['punkts'] or len(data['punkts']) == 0:
            return jsonify({'error': 'Необходимо добавить хотя бы один пункт приказа'}), 400
        
        # Генерируем документ
        doc_buffer = create_prikaz_document(data)
        
        # Формируем имя файла
        filename = f"Приказ_ПОЛАТИ_{data['orderNumber'].replace('/', '-')}.docx"
        
        return send_file(
            doc_buffer,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    
    except Exception as e:
        print(f"Ошибка при генерации: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Ошибка при генерации документа: {str(e)}'}), 500

def create_table_without_borders(doc, rows, cols):
    """Создание таблицы без границ"""
    table = doc.add_table(rows=rows, cols=cols)
    
    # Убираем границы
    tbl = table._element
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        border.set(qn('w:sz'), '0')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
        tblBorders.append(border)
    tblPr.append(tblBorders)
    
    return table

def create_prikaz_document(data):
    """Создание документа приказа согласно стандарту ПОЛАТИ 2025"""
    
    FONT_NAME = 'Times New Roman'
    
    # Создаем документ
    doc = Document()
    
    # Настройка полей страницы
    section = doc.sections[0]
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(1.5)
    
    # === HEADER С ЛОГОТИПОМ И РЕКВИЗИТАМИ ===
    section.different_first_page_header_footer = True
    first_page_header = section.first_page_header
    
    # Header как картинка
    HEADER_PATH = os.path.join(os.path.dirname(__file__), 'static', 'img', 'header.png')
    if os.path.exists(HEADER_PATH):
        para_header = first_page_header.paragraphs[0] if first_page_header.paragraphs else first_page_header.add_paragraph()
        para_header.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para_header.paragraph_format.space_after = Pt(0)
        run_header = para_header.add_run()
        try:
            # Используем всю ширину страницы (примерно 16 см)
            run_header.add_picture(HEADER_PATH, width=Cm(16))
        except Exception as e:
            print(f"Ошибка загрузки header: {e}")
            # Если ошибка, просто пропускаем
            pass
    
    # === ПРОПУСК СТРОКИ ПЕРЕД "ПРИКАЗ" ===
    doc.add_paragraph()
    
    # === ЗАГОЛОВОК "ПРИКАЗ" ===
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.space_before = Pt(12)
    para.paragraph_format.space_after = Pt(12)  # Пропуск строки после ПРИКАЗ
    run = para.add_run('ПРИКАЗ')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    run.font.bold = True
    
    # === ТАБЛИЦА С ДАТОЙ И НОМЕРОМ (3 ячейки) ===
    table_date = create_table_without_borders(doc, 1, 3)
    
    # Устанавливаем ширину колонок для даты
    table_date.columns[0].width = Cm(6.0)
    table_date.columns[1].width = Cm(5.0)
    table_date.columns[2].width = Cm(6.0)
    
    cells = table_date.rows[0].cells
    
    # Ячейка 1: День и месяц (слева)
    p = cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(0)
    
    run = p.add_run('«')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    run = p.add_run(data['day'])
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    run.font.underline = True
    
    run = p.add_run('» ')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    run = p.add_run(data['month'])
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    run.font.underline = True
    
    # Ячейка 2: Год (центр)
    p = cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    
    run = p.add_run(data['year'])
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    run.font.underline = True
    
    run = p.add_run(' г.')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    # Ячейка 3: Номер (справа)
    p = cells[2].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(0)
    
    run = p.add_run('№ ')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    run = p.add_run(data['orderNumber'])
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    run.font.underline = True
    
    # Пропуск строки после даты и номера
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(0)
    
    # === г. Мытищи ===
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(12)
    run = para.add_run('г. Мытищи')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    # === НАЗВАНИЕ ПРИКАЗА ===
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.space_after = Pt(12)
    run = para.add_run(data['orderTitle'])
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    run.font.bold = True
    
    # === ПРЕАМБУЛА ===
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para.paragraph_format.first_line_indent = Cm(1.25)
    para.paragraph_format.space_after = Pt(12)
    run = para.add_run(data['preamble'])
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    # === ПРИКАЗЫВАЮ ===
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.space_after = Pt(12)
    run = para.add_run('ПРИКАЗЫВАЮ:')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    run.font.bold = True
    
    # === ПУНКТЫ ПРИКАЗА ===
    for punkt in data['punkts']:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.space_after = Pt(0)
        run = para.add_run(f"{punkt['number']}. {punkt['text']}")
        run.font.name = FONT_NAME
        run.font.size = Pt(12)
    
    # === ФИНАЛЬНЫЕ ПУНКТЫ ===
    last_punkt_num = len(data['punkts']) + 1
    
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para.paragraph_format.space_after = Pt(12)
    run = para.add_run(f'{last_punkt_num}. Контроль исполнения настоящего приказа оставляю за собой.')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    # === 3 ПУСТЫЕ СТРОКИ (вместо 5) ===
    for _ in range(3):
        doc.add_paragraph()
    
    # === ПОДПИСЬ ГД (через таблицу) ===
    table_sign = create_table_without_borders(doc, 1, 3)
    
    table_sign.columns[0].width = Cm(6.0)
    table_sign.columns[1].width = Cm(5.0)
    table_sign.columns[2].width = Cm(6.0)
    
    cells = table_sign.rows[0].cells
    
    p = cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Генеральный директор')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    p = cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('__________________')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    p = cells[2].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('А.\u00A0А.\u00A0Зазыгин')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    # === БЛОК ОЗНАКОМЛЕНИЯ ===
    # Блок всегда добавляется с текстом "ФИО"
    doc.add_paragraph()
    doc.add_paragraph()
    
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = para.add_run('С приказом ознакомлен(-ы):')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    doc.add_paragraph()  # Только одна пустая строка
    
    # Строка с ФИО
    table_fio = create_table_without_borders(doc, 1, 2)
    
    table_fio.columns[0].width = Cm(2.5)
    table_fio.columns[1].width = Cm(14.5)
    
    cells = table_fio.rows[0].cells
    
    # Ячейка 1: ФИО (всегда)
    p = cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('ФИО')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    # Ячейка 2: Линия и дата (выравнивание по левому краю)
    p = cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    run = p.add_run('_________________________________')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    run = p.add_run('\u00A0')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    run = p.add_run('«__»_______20__г.')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    doc.add_paragraph()  # Только одна пустая строка
    
    # Строка "Подпись"
    table_podpis = create_table_without_borders(doc, 1, 2)
    
    table_podpis.columns[0].width = Cm(2.5)
    table_podpis.columns[1].width = Cm(14.5)
    
    cells = table_podpis.rows[0].cells
    
    p = cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Подпись')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    p = cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    run = p.add_run('_________________________________________')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    
    # Сохраняем в буфер
    doc_buffer = BytesIO()
    doc.save(doc_buffer)
    doc_buffer.seek(0)
    
    return doc_buffer

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
