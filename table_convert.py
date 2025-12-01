import pandas as pd
import numpy as np

file_path = "C:\\tools\\favorite_songs_before_changed.xlsx"
df = pd.read_excel(file_path, engine="openpyxl")

df['song'] = df['song'].str.lstrip('*')

date_mapping = {
    "2025": pd.Timestamp("2025-04-01"),
    "2024": pd.Timestamp("2024-04-01"),
    "2023": pd.Timestamp("2023-04-01"),
    "2022": pd.Timestamp("2022-04-01"),
    "2021": pd.Timestamp("2021-04-01"),
    "大4": pd.Timestamp("2020-04-01"),
    "大3": pd.Timestamp("2019-04-01"),
    "大2": pd.Timestamp("2018-04-01"),
    "大1": pd.Timestamp("2017-04-01"),
    "高3": pd.Timestamp("2016-04-01"),
    "高2": pd.Timestamp("2015-04-01"),
    "高1": pd.Timestamp("2014-04-01"),
    "中3": pd.Timestamp("2013-04-01"),
    "中2": pd.Timestamp("2012-04-01"),
    "中1": pd.Timestamp("2011-04-01"),
    "小6": pd.Timestamp("2010-04-01"),
    "小5": pd.Timestamp("2009-04-01"),
    "小4": pd.Timestamp("2008-04-01"),
    "小3": pd.Timestamp("2007-04-01"),
    "小2": pd.Timestamp("2006-04-01"),
    "小1": pd.Timestamp("2005-04-01"),
    "年長": pd.Timestamp("2004-04-01"),
    "年中": pd.Timestamp("2003-04-01")
}

# 日付変換関数を定義
def parse_period(row):
    try:
        period = row['period']
        idx = row.name  # index番号を取得

        if pd.isna(period):  # 欠損値はそのまま
            return None

        # スペースが含まれる場合（年＋月または年＋月.日）
        if " " in period:
            base_period, rest = period.split(" ", 1)
            if "." in rest:  # 月.日形式の場合
                month, day = map(int, rest.split("."))
            else:  # 月のみの場合
                month = int(rest)
                day = 1  # デフォルトで1日
            if base_period in date_mapping:
                # ベースの日付から年を取得し、新しい日付を作成
                base_date = date_mapping[base_period]
                return pd.Timestamp(year=base_date.year, month=month, day=day)

        # 通常のマッピングを適用
        return date_mapping.get(period, None)

    except Exception as e:
        # エラーが発生した場合にindexとエラー内容を出力
        print(f"Error at index {idx}: {e}")
        return None

def parse_period2(row):
    try:
        period2 = row['period2']
        idx = row.name  # index番号を取得

        if pd.isna(period2):  # 欠損値はそのまま
            return None

        # スペースが含まれる場合（年＋月または年＋月.日）
        if " " in period2:
            base_period, rest = period2.split(" ", 1)
            if "." in rest:  # 月.日形式の場合
                month, day = map(int, rest.split("."))
            else:  # 月のみの場合
                month = int(rest)
                day = 1  # デフォルトで1日
            if base_period in date_mapping:
                # ベースの日付から年を取得し、新しい日付を作成
                base_date = date_mapping[base_period]
                return pd.Timestamp(year=base_date.year, month=month, day=day)

        # 通常のマッピングを適用
        return date_mapping.get(period2, None)

    except Exception as e:
        # エラーが発生した場合にindexとエラー内容を出力
        print(f"Error at index {idx}: {e}")
        return None

def parse_period3(row):
    try:
        period3 = row['period3']
        idx = row.name  # index番号を取得

        if pd.isna(period3):  # 欠損値はそのまま
            return None

        # スペースが含まれる場合（年＋月または年＋月.日）
        if " " in period3:
            base_period, rest = period3.split(" ", 1)
            if "." in rest:  # 月.日形式の場合
                month, day = map(int, rest.split("."))
            else:  # 月のみの場合
                month = int(rest)
                day = 1  # デフォルトで1日
            if base_period in date_mapping:
                # ベースの日付から年を取得し、新しい日付を作成
                base_date = date_mapping[base_period]
                return pd.Timestamp(year=base_date.year, month=month, day=day)

        # 通常のマッピングを適用
        return date_mapping.get(period3, None)

    except Exception as e:
        # エラーが発生した場合にindexとエラー内容を出力
        print(f"Error at index {idx}: {e}")
        return None

def parse_period4(row):
    try:
        period4 = row['period4']
        idx = row.name  # index番号を取得

        if pd.isna(period4):  # 欠損値はそのまま
            return None

        # スペースが含まれる場合（年＋月または年＋月.日）
        if " " in period4:
            base_period, rest = period4.split(" ", 1)
            if "." in rest:  # 月.日形式の場合
                month, day = map(int, rest.split("."))
            else:  # 月のみの場合
                month = int(rest)
                day = 1  # デフォルトで1日
            if base_period in date_mapping:
                # ベースの日付から年を取得し、新しい日付を作成
                base_date = date_mapping[base_period]
                return pd.Timestamp(year=base_date.year, month=month, day=day)

        # 通常のマッピングを適用
        return date_mapping.get(period4, None)
    except Exception as e:
        # エラーが発生した場合にindexとエラー内容を出力
        print(f"Error at index {idx}: {e}")
        return None


# period列を解析してperiod-date列を作成
df['period-date'] = df.apply(parse_period, axis=1)
df['period-date2'] = df.apply(parse_period2, axis=1)
df['period-date3'] = df.apply(parse_period3, axis=1)
df['period-date4'] = df.apply(parse_period4, axis=1)

print(df)

df.to_excel('C:\\tools\\favorite_songs.xlsx', index=False)