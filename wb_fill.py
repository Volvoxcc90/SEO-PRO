import openpyxl
from utils import load_brands_ru, normalize_brand_key, guess_ru, logging
import time

def get_ai_generator():
    try:
        from transformers import pipeline
        import torch
        return pipeline('text-generation', model='sberbank-ai/rugpt3small_based_on_gpt2', device=0 if torch.cuda.is_available() else -1)
    except:
        return None

def fill_wb_template(input_xlsx: str, brand: str, collection: str, quality: str = 'medium', length: str = 'normal', style: str = 'normal', use_ai: bool = False, progress_callback=None):
    start_time = time.time()
    logging.info(f"Starting processing: {input_xlsx}")
    ru_brand = load_brands_ru().get(normalize_brand_key(brand), guess_ru(brand))
    generator = get_ai_generator() if use_ai else None

    wb = openpyxl.load_workbook(input_xlsx)
    ws = wb.active

    headers = [cell.value for cell in ws[1] if cell.value]
    num_rows = ws.max_row - 1
    if num_rows <= 0:
        raise ValueError("Нет данных")

    brand_col = headers.index('Бренд') + 1 if 'Бренд' in headers else None
    name_col = headers.index('Название') + 1 if 'Название' in headers else None
    desc_col = headers.index('Описание') + 1 if 'Описание' in headers else None
    if not any([brand_col, name_col, desc_col]):
        raise ValueError("Отсутствуют колонки")

    count = 0
    report = []
    for row_idx in range(2, ws.max_row + 1):
        if brand_col:
            ws.cell(row=row_idx, column=brand_col).value = ru_brand

        model = ""
        if name_col:
            current_name = ws.cell(row=row_idx, column=name_col).value or ''
            model = current_name.strip()
            if generator:
                prompt = f"SEO-название для очков {ru_brand} {model} {collection} (стиль: {style}, длина: {length})"
                new_name = generator(prompt, max_length=50)[0]['generated_text'].strip()
            else:
                new_name = f"Солнцезащитные очки {ru_brand} {model} {collection} стильные с UV защитой"
            ws.cell(row=row_idx, column=name_col).value = new_name
            report.append(f"Строка {row_idx}: {new_name}")

        if desc_col:
            current_desc = ws.cell(row=row_idx, column=desc_col).value or ''
            if generator:
                prompt = f"SEO-описание для очков {ru_brand} {model} {collection} (качество: {quality}, длина: {length})"
                seo_desc = generator(prompt, max_length=200)[0]['generated_text'].strip()
            else:
                seo_desc = f"{current_desc}\nКупить {ru_brand} {collection} по выгодной цене..."
            ws.cell(row=row_idx, column=desc_col).value = seo_desc

        count += 1
        if progress_callback:
            progress_callback((count / num_rows) * 100)

    output_path = input_xlsx.replace('.xlsx', '_seo.xlsx')
    wb.save(output_path)

    elapsed = time.time() - start_time
    report_str = '\n'.join(report) + f"\nВремя: {elapsed:.2f} сек"
    return output_path, count, report_str
