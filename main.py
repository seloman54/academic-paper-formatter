import os
from docx import Document

def format_paper(file_path):
    doc = Document(file_path)

    # Genel yazı tipi ve punto
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = 12

    # Başlık biçimleri
    for i in range(1, 4):
        try:
            h = doc.styles[f'Heading {i}']
            h.font.name = 'Times New Roman'
            h.font.size = 14 if i == 1 else 12
            h.font.bold = True
        except:
            pass

    # Sayfa kenar boşlukları (örnek)
    section = doc.sections[0]
    section.top_margin = 72 * 1  # 1 inç
    section.bottom_margin = 72 * 1
    section.left_margin = 72 * 1
    section.right_margin = 72 * 1

    # Kaydedilen dosyayı yeni isimle oluştur
    new_name = "formatted_" + os.path.basename(file_path)
    doc.save(os.path.join("outputs", new_name))
    return new_name


def main():
    uploads_folder = "uploads"
    outputs_folder = "outputs"
    os.makedirs(outputs_folder, exist_ok=True)

    for filename in os.listdir(uploads_folder):
        if filename.endswith(".docx"):
            file_path = os.path.join(uploads_folder, filename)
            new_file = format_paper(file_path)
            print(f"{filename} dosyası biçimlendirildi → {new_file}")

if __name__ == "__main__":
    main()
