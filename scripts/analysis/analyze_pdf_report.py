"""
Анализ структуры PDF отчета
"""
import PyPDF2
import os

pdf_folder = '123'
pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]

print("=" * 80)
print("АНАЛИЗ PDF ОТЧЕТОВ")
print("=" * 80)

for pdf_file in pdf_files:
    pdf_path = os.path.join(pdf_folder, pdf_file)
    print(f"\n📄 Файл: {pdf_file}")
    print("-" * 80)
    
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            num_pages = len(pdf_reader.pages)
            print(f"Страниц: {num_pages}\n")
            
            # Читаем первые 3 страницы для анализа структуры
            for page_num in range(min(3, num_pages)):
                print(f"\n--- Страница {page_num + 1} ---")
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                
                # Показываем первые 2000 символов
                print(text[:2000])
                print("\n...")
                
    except Exception as e:
        print(f"Ошибка при чтении: {e}")

print("\n" + "=" * 80)
