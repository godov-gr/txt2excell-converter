def parse_text_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    data = []
    for line in lines:
        stripped = line.strip()
        if stripped:
            # Разделитель — таб или запятая.
            parts = stripped.split('\t') if '\t' in stripped else stripped.split(',')
            data.append(parts)
    return data
