import pandas as pd
from docx import Document

file_path = "C:\\tools\\fs_20240114.docx"

doc = Document(file_path)
print(file_path)

data = []
for paragraph in doc.paragraphs:
    if paragraph.text.strip():  # 空行を除外
        data.append(paragraph.text)

df = pd.DataFrame(data, columns=["Content"])


df["Content"] = df["Content"].str.lstrip()

df['Index'] = range(1, len(df) + 1)

expanded_data = []
for index, row in df.iterrows():
    split_contents = row['Content'].split('\n\t')  # \n\tで分割
    for content in split_contents:
        expanded_data.append({'Index': row['Index'], 'Content': content})

new_df = pd.DataFrame(expanded_data)
new_df.reset_index(drop=True, inplace=True)
new_df['Index'] = range(1, len(new_df) + 1)


new_df['artist'] = new_df['Content'].where(new_df['Content'].str.endswith(":"), None)
new_df['artist'] = new_df['artist'].fillna(method='ffill')
new_df = new_df[new_df['Content'].str.endswith(":") == False].reset_index(drop=True)
new_df['Index'] = range(1, len(new_df) + 1)

new_df['artist'] = new_df['artist'].str.rstrip(':')
new_df['period'] = new_df['Content'].str.extract(r'\[(.*?)\]')
new_df['Content'] = new_df['Content'].str.replace(r'\s*\[.*?\]$', '', regex=True)
new_df.rename(columns={'Content': 'song'}, inplace=True)


# period列の値を「・」または全角スペース「　」で分割
split_periods = new_df['period'].str.split(r'[・　]', expand=True)

# period列を1つ目の要素に書き換える
new_df['period'] = split_periods[0]

# 新しい列 (period2, period3, ...) を動的に作成
for i in range(1, split_periods.shape[1]):
    new_df[f'period{i+1}'] = split_periods[i]

print(new_df[['song', 'artist']][760:780])
#new_df.to_excel('C:\\tools\\favorite_songs.xlsx', index=False)