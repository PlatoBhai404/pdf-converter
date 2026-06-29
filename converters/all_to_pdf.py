import os
from PIL import Image
from docx2pdf import convert as docx_to_pdf_convert

def images_to_pdf(image_paths, output_pdf_path):
    if not image_paths: return False
    opened_images = []
    try:
        for img_path in image_paths:
            if os.path.exists(img_path):
                img = Image.open(img_path)
                if img.mode != 'RGB': img = img.convert('RGB')
                opened_images.append(img)
        if not opened_images: return False
        opened_images[0].save(output_pdf_path, "PDF", save_all=True, append_images=opened_images[1:])
        return True
    except Exception: return False
    finally:
        for img in opened_images: img.close()

def docx_to_pdf(docx_path, output_pdf_path):
    try:
        docx_to_pdf_convert(docx_path, output_pdf_path)
        return True
    except Exception: return False

def txt_to_pdf(txt_path, output_pdf_path):
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        doc = SimpleDocTemplate(output_pdf_path, pagesize=letter)
        story = []
        style = ParagraphStyle('Norm', parent=getSampleStyleSheet()['Normal'], fontSize=11, leading=16)
        with open(txt_path, 'r', encoding='utf-8') as f:
            for line in f.readlines():
                if line.strip() == "": story.append(Spacer(1, 10))
                else: story.append(Paragraph(line.replace('&','&amp;').replace('<','&lt;'), style))
        doc.build(story)
        return True
    except Exception: return False

def html_to_pdf(html_path, output_pdf_path):
    """Renders HTML layout structures cleanly to PDF via WeasyPrint."""
    try:
        from weasyprint import HTML
        HTML(html_path).write_pdf(output_pdf_path)
        return True
    except Exception as e:
        print(f"HTML conversion error: {e}")
        return False

def pptx_to_pdf(pptx_path, output_pdf_path):
    """Leverages local Windows COM Interop automation to render PPTX slides to PDF."""
    try:
        import win32com.client
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        # Open presentation quietly in background background
        deck = powerpoint.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        # 32 is the code value representing PDF format execution in PowerPoint API
        deck.SaveAs(os.path.abspath(output_pdf_path), 32)
        deck.Close()
        powerpoint.Quit()
        return True
    except Exception as e:
        print(f"PPTX conversion error: {e}")
        return False