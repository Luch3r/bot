import json
import os
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

class PresentationGenerator:
    def __init__(self, json_file):
        self.json_file = json_file
        self.prs = Presentation()
        self.slide_titles = [] 
        self.load_data()
    
    def load_data(self):
        """Загрузка данных из JSON файла"""
        with open(self.json_file, 'r', encoding='utf-8') as f:
            self.data = json.load(f)
    
    def set_presentation_properties(self):
        """Установка свойств презентации"""
        if 'title' in self.data['presentation']:
            self.prs.core_properties.title = self.data['presentation']['title']
        if 'author' in self.data['presentation']:
            self.prs.core_properties.author = self.data['presentation']['author']
    
    def add_slide_numbers(self):
        """Добавление номеров слайдов на все слайды"""
        for i, slide in enumerate(self.prs.slides):
            left = Inches(8.5)
            top = Inches(7)
            width = Inches(1)
            height = Inches(0.5)
            
            textbox = slide.shapes.add_textbox(left, top, width, height)
            text_frame = textbox.text_frame
            text_frame.text = str(i + 1)
            if text_frame.paragraphs:
                text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
                text_frame.paragraphs[0].font.size = Pt(12)
    
    def create_table_of_contents(self):
        """Создание слайда с оглавлением"""
        if len(self.slide_titles) <= 1:
            return
            
        slide_layout = self.prs.slide_layouts[1]
        slide = self.prs.slides.add_slide(slide_layout)
        
        if slide.shapes.title:
            slide.shapes.title.text = "Оглавление"
        
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9)
        height = Inches(5)
        
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        
        
        text_frame.clear()
        for i, title in enumerate(self.slide_titles[1:], 2): 
            p = text_frame.add_paragraph()
            p.text = f"{i-1}. {title}"
            p.level = 0
            p.font.size = Pt(18)
    
    def create_slide(self, slide_data):
        """Создание слайда на основе данных"""
        layout_index = slide_data.get('layout', 1)
        slide_layout = self.prs.slide_layouts[layout_index]
        slide = self.prs.slides.add_slide(slide_layout)
        
        title_text = slide_data.get('title', '')
        if title_text:
            self.slide_titles.append(title_text)
        
        if slide.shapes.title:
            slide.shapes.title.text = title_text
        
        self.handle_slide_content(slide, slide_data)
        
        return slide
    
    def handle_slide_content(self, slide, slide_data):
        """Универсальная обработка контента слайда"""
        layout = slide_data.get('layout', 1)
        
        if layout == 0 and 'subtitle' in slide_data:
            self.add_subtitle(slide, slide_data['subtitle'])
        
        if 'content' in slide_data:
            self.add_content_to_slide(slide, slide_data['content'], layout)
        
        if layout == 3: 
            if 'left_content' in slide_data:
                self.add_content_to_slide(slide, slide_data['left_content'], layout, 'left')
            if 'right_content' in slide_data:
                self.add_content_to_slide(slide, slide_data['right_content'], layout, 'right')
        
        if 'images' in slide_data:
            for img_data in slide_data['images']:
                self.add_image(slide, img_data)
    
    def add_subtitle(self, slide, subtitle_text):
        """Добавление подзаголовка на титульный слайд"""
        
        for placeholder in slide.placeholders:
            if placeholder.placeholder_format.type == 15:  
                placeholder.text = subtitle_text
                return
        
        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = subtitle_text
    
    def add_content_to_slide(self, slide, content_data, layout, position='center'):
        """Добавление контента на слайд"""
        
        if layout == 3:  
            if position == 'left':
                left, top, width, height = Inches(0.5), Inches(1.5), Inches(4.2), Inches(5)
            else:  
                left, top, width, height = Inches(4.8), Inches(1.5), Inches(4.2), Inches(5)
        else:
            left, top, width, height = Inches(0.5), Inches(1.5), Inches(9), Inches(5)
        
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        
        for item in content_data:
            if item['type'] == 'text':
                self.add_text_item(text_frame, item)
            elif item['type'] == 'table':
                self.add_table_to_slide(slide, item)
    
    def add_text_item(self, text_frame, item):
        """Добавление текстового элемента с форматированием"""
        p = text_frame.add_paragraph()
        p.text = item['text']
        p.level = item.get('level', 0)
        
        if 'style' in item:
            self.apply_text_styles(p, item['style'])
    
    def apply_text_styles(self, paragraph, styles):
        """Применение стилей к тексту"""
        font = paragraph.font
        
        if 'bold' in styles and styles['bold']:
            font.bold = True
        if 'italic' in styles and styles['italic']:
            font.italic = True
        if 'size' in styles:
            font.size = Pt(styles['size'])
        if 'color' in styles:
            color = styles['color']
            font.color.rgb = RGBColor(color[0], color[1], color[2])
    
    def add_table_to_slide(self, slide, table_data):
        """Добавление таблицы на слайд"""
        rows = len(table_data['data'])
        cols = len(table_data['data'][0]) if rows > 0 else 0
        
        if rows == 0 or cols == 0:
            return
        
        left = Inches(1)
        top = Inches(2)
        width = Inches(8)
        height = Inches(0.6 * rows)  
        
        table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        table = table_shape.table
        
        for i, row in enumerate(table_data['data']):
            for j, cell_text in enumerate(row):
                cell = table.cell(i, j)
                cell.text = str(cell_text)
                
                if cell.text_frame.paragraphs:
                    cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                
                if i == 0 and table_data.get('header', False):
                    if cell.text_frame.paragraphs:
                        paragraph = cell.text_frame.paragraphs[0]
                        paragraph.font.bold = True
                        paragraph.font.size = Pt(14)
    
    def add_image(self, slide, img_data):
        """Добавление изображения"""
        try:
            if os.path.exists(img_data['path']):
                left = Inches(img_data.get('left', 1))
                top = Inches(img_data.get('top', 1))
                width = Inches(img_data.get('width', 4))
                height = Inches(img_data.get('height', 3))
                
                slide.shapes.add_picture(
                    img_data['path'], 
                    left, top, width, height
                )
                print(f"Изображение добавлено: {img_data['path']}")
            else:
                print(f"Файл изображения не найден: {img_data['path']}")
        except Exception as e:
            print(f"Ошибка при добавлении изображения {img_data['path']}: {e}")
    
    def generate(self):
        """Генерация презентации"""
        try:
            self.set_presentation_properties()
            
            for slide_data in self.data['presentation']['slides']:
                self.create_slide(slide_data)
            
            if self.data['presentation'].get('table_of_contents', False):
                self.create_table_of_contents()
            
            self.add_slide_numbers()
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"presentation_{timestamp}.pptx"
            self.prs.save(filename)
            
            print(f"Презентация успешно создана: {filename}")
            print(f"Количество слайдов: {len(self.prs.slides)}")
            return filename
            
        except Exception as e:
            print(f"Ошибка при создании презентации: {e}")
            import traceback
            traceback.print_exc()
            return None

def create_test_presentation():
    """Создание тестовой презентации с простой структурой"""
    test_data = {
        "presentation": {
            "title": "Тестовая презентация",
            "author": "Presentation Generator",
            "table_of_contents": True,
            "slides": [
                {
                    "layout": 0,
                    "title": "Моя тестовая презентация",
                    "subtitle": "Создано автоматически"
                },
                {
                    "layout": 1,
                    "title": "Первый слайд с текстом",
                    "content": [
                        {"type": "text", "text": "Основной заголовок", "level": 0, "style": {"bold": True, "size": 16}},
                        {"type": "text", "text": "Первый пункт", "level": 1},
                        {"type": "text", "text": "Второй пункт", "level": 1},
                        {"type": "text", "text": "Вложенный пункт", "level": 2}
                    ]
                },
                {
                    "layout": 1,
                    "title": "Слайд с таблицей",
                    "content": [
                        {
                            "type": "table",
                            "header": True,
                            "data": [
                                ["Параметр", "Значение", "Статус"],
                                ["Скорость", "100 м/с", "Отлично"],
                                ["Качество", "95%", "Хорошо"],
                                ["Надежность", "99.9%", "Отлично"]
                            ]
                        }
                    ]
                },
                {
                    "layout": 3,
                    "title": "Слайд с двумя колонками",
                    "left_content": [
                        {"type": "text", "text": "Левая колонка", "level": 0, "style": {"bold": True}},
                        {"type": "text", "text": "Пункт 1", "level": 1},
                        {"type": "text", "text": "Пункт 2", "level": 1}
                    ],
                    "right_content": [
                        {"type": "text", "text": "Правая колонка", "level": 0, "style": {"bold": True}},
                        {"type": "text", "text": "Пункт A", "level": 1},
                        {"type": "text", "text": "Пункт B", "level": 1}
                    ]
                }
            ]
        }
    }
    
    with open('test_presentation.json', 'w', encoding='utf-8') as f:
        json.dump(test_data, f, ensure_ascii=False, indent=2)
    
    return 'test_presentation.json'

if __name__ == "__main__":
    print("Создание тестовой презентации...")
    
    test_file = create_test_presentation()
    print(f"Создан тестовый файл: {test_file}")
    
    generator = PresentationGenerator(test_file)
    result = generator.generate()
    
    if result:
        print("✅ Презентация успешно создана!")
        print(f"📁 Файл: {result}")
    else:
        print("❌ Не удалось создать презентацию")