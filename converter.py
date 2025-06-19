import pandas as pd

def export_to_excel(data, output_path):
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)
