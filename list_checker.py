import pandas as pd

# Excelファイルの読み込み
file_path = "list_of_members.xlsx"  # ファイル名を正しく指定
df = pd.read_excel(file_path, sheet_name="Check")

# G列（番号）と I列（名前）を抽出（位置で指定）
numbers = df.iloc[:, 6]
names = df.iloc[:, 8]

# 番号と名前のペアを含むデータフレーム作成
pair_df = pd.DataFrame({'番号': numbers, '名前': names}).dropna()

# 番号ごとに名前をグループ化し、重複している番号のみ抽出
grouped = pair_df.groupby('番号')['名前'].apply(list)
duplicates = grouped[grouped.apply(lambda x: len(x) > 1)]

# 重複番号を横に展開（1列目に番号、右に名前）
final_rows = [[number] + name_list for number, name_list in duplicates.items()]
max_names = max(len(row) for row in final_rows)
columns = ["番号"] + [f"名前{i+1}" for i in range(max_names - 1)]
final_df = pd.DataFrame(final_rows, columns=columns)

# 出力ファイルに保存
final_df.to_excel("final_result.xlsx", index=False)
