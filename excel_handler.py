import pandas as pd
from config import EXCEL_FILE

def save_to_excel(data_list):
    df = pd.DataFrame(data_list)
    df.to_excel(EXCEL_FILE, index=False)
