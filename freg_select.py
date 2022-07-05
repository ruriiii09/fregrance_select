import gspread
from google.oauth2.service_account import Credentials
import itertools
import numpy as np
import json
from operator import itemgetter
import pprint
import collections

#これ指定しないとキーを発行し続けないといけないらしい
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
#連携のための定型分以下2行　jsonファイルを指定フルパスの方がいい　scopes=scopeにしたらエラー無くなった
creds = Credentials.from_service_account_file('/Users/ruri/Documents/programing/python_games/service_account.json',scopes=scope)
client = gspread.authorize(creds)

#ワークシートを取得
spreadsheet_key = "1Fh4rs8UPnPRd4kltYPlo0o5qQbXVSxdGQoPXkoo_DWY"

#ワークシートを開く
wb = client.open_by_key(spreadsheet_key)

# シートの一覧を取得する。（リスト形式）
worksheets = wb.worksheets()
#print(worksheets)

# シートを開く
worksheet = wb.worksheet('入力')
freglist = wb.worksheet("香水リスト")

# セルA1に”test”という文字列を代入する。
#worksheet.update_cell(1, 2, 'test')

#セルの内容を取得
images = worksheet.get("A1:C1")
fregname = freglist.get("A2:I30")

#多次元リストをitertoolによって平坦化
images = list(itertools.chain.from_iterable(images))

#numpyで扱えるリスト型に変換
freg = np.array(fregname)
#worksheet.update_acell("A1","テスト")


#------------以下条件一致を探索、返す処理------------

#イメージに合致する香水名をリストansに入れる
ans = []
for i in range(len(freg)):
    #それぞれのイメージをimgとおく
    img = freg[i,5:8]
    name = freg[i,1]
    #受け取ったイメージが香水のimgの中にあればその香水名を返す
    for j in images:
        if j in img:
            ans.append(name)

#受け取った香水名の出現回数をカウント
res = collections.Counter(ans)
#カウント数が一番多いものを出力
print("今日のあなたにぴったりな香水はこれ！")
for i in range(len(res)):
    if res.most_common()[i][1] == res.most_common()[0][1]:
        print(res.most_common()[i][0])
