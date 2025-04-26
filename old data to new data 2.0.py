import pandas as pd
import re
# データファイルを読み込む。
new_data_path = r"C:\Users\XXX.xlsx"

# エクセルファイルを読む。
new_data = pd.read_excel(new_data_path, header=None)
print(new_data.iloc[1,3])

# 照合されたダイアログを格納するDataFrameを初期化する。
#filtered_dialogues_organized = pd.DataFrame(columns=['Speaker', 'Time', 'Content'])
filtered_dialogues_organized = pd.DataFrame()

# ダイアログにマッチする正規表現を定義する。「がん相談員」と「相談者」の対話をより明確に分割する。
dialogue_pattern = re.compile(r'\[(がん相談員|相談者)?\s?(\d{2}:\d{2}:\d{2})\](.*?)(?=\[\D|\Z)', re.DOTALL)

# 各セルでの対話内容を処理する。
for i, cell_content in enumerate(new_data[8]):
    dialogues = dialogue_pattern.findall(str(cell_content))

    for dialogue in dialogues:
        speaker_tag, time, content = dialogue
        speaker = speaker_tag if speaker_tag else "相談者"

        # 新しい DataFrame エントリを作成します。
        entry_df = pd.DataFrame({'ChatID': [new_data.iloc[i,0]], 'Speaker': [speaker], 'date': [str(new_data.iloc[i,3]).split(' ')[0] ], 'Time': [time], 'Content': [content.strip()], 'evalution': [new_data.iloc[i,9]], 'visitortag': [new_data.iloc[i,13]], 'visitortag2': [new_data.iloc[i,14]], 'visitortag3': [new_data.iloc[i,15]], 'visitortag4': [new_data.iloc[i,16]], 'visitortag5': [new_data.iloc[i,17]] })

        # DataFrameをまとめる
        filtered_dialogues_organized = pd.concat([filtered_dialogues_organized, entry_df], ignore_index=True)
# 保存先のファイルパスを指定する。
filtered_dialogues_organized_file_path = r"C:\Users\XXXXXX.xlsx"
# 新しいexcelファイルを保存する。
filtered_dialogues_organized.to_excel(filtered_dialogues_organized_file_path, index=False)
