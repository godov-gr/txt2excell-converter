import csv

def parse_text_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f if line.strip()]

    # Без строк — пустая ячейка.
    if not lines:
        return []

    # Определение разделителя.
    sample = "\n".join(lines[:5])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters='\t,;|')
        delimiter = dialect.delimiter
    except Exception:
        delimiter = ','  

    # Парсинг.
    data = [line.split(delimiter) for line in lines]

    # Выравнивание строк по макс длине.
    max_len = max(len(row) for row in data)
    for row in data:
        if len(row) < max_len:
            row.extend([''] * (max_len - len(row)))

    return data