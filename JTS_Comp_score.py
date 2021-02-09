#%%
import pandas as pd
import glob
import os

#%%
# ファイルの場所を指定

input_data = 'input_files/'
output_data = 'output_files/'


# データからの検索をするためのリスト

teams = [
    '金沢学院大学クラブ',
    'アベノジュニアトランポリンクラブ',
    '厚木FUSiONスポーツクラブ',
    '日本体育大学トランポリンクラブ',
    '星稜クラブ',
    'キタイスポーツクラブ'
]

names = [
    'ひかる',
    '裕美',
    '愛菜',
    '沙季',
    'セリナ',
    'ここね',
    '守杏',
    '晴茄',
    '希湖',
    '美咲希'
]

stages = [
    'Qualification',
    'Final'
]

#%%
# 上記のリストを利用し、クラブ名、名前、ルーティン名で検索して一人のデータ取り出して df1 に格納
# その後、各ジャッジの点数が列として格納されているため、一旦ジャッジ名と点数のみを取り出して転置し df2 に格納
# 各スキルの原点を取り出して df3 に格納
# さらに、df1, df2, df3 のインデックスを一旦クリアし結合処理
# 上記すべてを関数化

def main(team, name, stage):
    df1 = df[df['Representing'] == team]
    df1 = df1[df1['Given Name'] == name]
    df1 = df1[df1['Stage'] == stage]
    df1 = df1.iloc[:1, [0, 1, 2, 3, 4, 5]]

    df2= df[df['Representing'] == team]
    df2 = df2[df2['Given Name'] == name]
    df2 = df2[df2['Stage'] == stage]
    df2 = df2.loc[:, ['Judge','∑']]
    df2 = df2.set_index('Judge')
    df2 = df2[0:10].T

    df3 = df[df['Representing'] == team]
    df3 = df3[df3['Given Name'] == name]
    df3 = df3[df3['Stage'] == stage]
    df3 = df3.iloc[6:7, [8, 9, 10, 11, 12, 13, 14, 15, 16, 17]]

    df1.reset_index(drop=True, inplace=True)
    df2.reset_index(drop=True, inplace=True)
    df3.reset_index(drop=True, inplace=True)

    df4 = pd.concat([df1, df2, df3], axis=1)

    return df4

#%%
# 情報を取り出したいフォルダにある xlsx ファイルを一つずつ取り出して
# ファイル名のみを file_name に格納して DataFrame として読み込む

for fname in glob.glob(input_data+'*.xlsx'):
    file_name = os.path.split(fname)[1]
    
    # 各大会のスコアデータの読み込み時に必要なデータのみ抽出
    df = pd.read_excel(input_data+file_name)
    df = df[[
    'Competition',
    'Stage',
    'Surname',
    'Given Name',
    'Representing',
    'Mark',
    'Judge',
    '∑',
    'S1',
    'S2',
    'S3',
    'S4',
    'S5',
    'S6',
    'S7',
    'S8',
    'S9',
    'S10',
    'L'
    ]]
#%%
    # result_df という空のリストを用意
    # for文にて、各選手の情報をリストから取り出して検索し、main() に渡し、
    # 取り出された各選手のデータを result_df にリストとして格納
    # result_df に格納された各選手のデータを結合処理

    result_df = []

    for team in teams:
        for name in names:
            for stage in stages:
                result_df.append(main(team, name, stage))

    result_df = pd.concat(result_df)
    
    # すべてデータの無い行をドロップ
    result_df = result_df.dropna(how='all')
        
#%%    
    # lambda を使って、各列のスコアの桁数を調整
    result_df['Mark'] = result_df.apply(lambda x: x['Mark']/1000, axis=1)
    result_df['E1'] = result_df.apply(lambda x: x['E1']/10, axis=1)
    result_df['E2'] = result_df.apply(lambda x: x['E2']/10, axis=1)
    result_df['E3'] = result_df.apply(lambda x: x['E3']/10, axis=1)
    result_df['E4'] = result_df.apply(lambda x: x['E4']/10, axis=1)
    result_df['E5'] = result_df.apply(lambda x: x['E5']/10, axis=1)
    result_df['E8'] = result_df.apply(lambda x: x['E8']/10, axis=1)
    result_df['E∑'] = result_df.apply(lambda x: x['E∑']/10, axis=1)
    result_df['T'] = result_df.apply(lambda x: x['T']/1000, axis=1)
    result_df['D'] = result_df.apply(lambda x: x['D']/10, axis=1)
    result_df['H'] = result_df.apply(lambda x: x['H']/100, axis=1)
 
    # 読み込んだデータファイルに result_ を prefix として追加して output_files に保存
    result_df.to_excel(output_data+'results_'+file_name, index=False)