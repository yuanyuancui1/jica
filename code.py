
import pandas as pd
from pathlib import Path
import os
import unicodedata

def convert_to_halfwidth(text):
    return unicodedata.normalize('NFKC', text)

path = os.path.join(str(Path.home()),'Documents\Magic Folder')
df = pd.concat([pd.read_excel(f) for f in Path(path).rglob('*.xlsx')])

df['message_subject'] = df.message_subject.astype(str).apply(lambda x: convert_to_halfwidth(x).lower())

with open(os.path.join(path, 'search.txt'), 'r', encoding='utf-8') as file:
    search = file.read()
    search = convert_to_halfwidth(search).lower()

df_new = df[df.message_subject.str.contains(search)]
df_new.to_excel('final.xlsx', index = False)
