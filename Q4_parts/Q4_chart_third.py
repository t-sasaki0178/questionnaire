# 職業（ラジオボタンに含まれるテキスト回答その１）
import csv
import codecs
#グラフ用
import matplotlib
import matplotlib.pyplot as plt
# データの扱いに必要
import pandas as pd
# 画像の保存先の指定に必要
import os
# PP（パワーポイント）の作成、挿入
from pptx import Presentation
from pptx.util import Inches
# 画像の読み込み
from glob import glob
import re

from jinja2 import Environment,FileSystemLoader 
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
# driverまでのパスを求められることがあるが、以下１文のみで解決できる（理由は不明。おそらくモジュール内のdriverを使ってる？）
import chromedriver_binary
# nanの判定のため
import math

def Q4_chart_third():
    # フォントの指定（日本語が文字化けしないもの）
    matplotlib.rcParams['font.family'] = 'IPAexGothic'

    # csv読み込み版
    layout_load = pd.read_csv('./samples/layout.csv',encoding='cp932',header=0)
    # ldata:id（インデックス）、残り全部itemに入ってる
    for ldata,item in layout_load.iterrows():
        if ('Q4_1' in item[1]) == True:
            # 数学モジュール（math）で判定する。nan（値の入っていない状態を判別するため）
            # item[4]がnanであればTrue
            try:
                if math.isnan(item[4]):
                    l_title = item[5]
            except:
                l_title = item[4]


    # labelデータから問題の回答を取得する
    raw_load = pd.read_csv('./samples/rawdata.csv',header=None,encoding='shift_jisx0213')

    # labelデータの一列目から関連するインデックスを取得する
    # Q4_1に関連する部分を割り出す
    rels = []
    ans = []
    rcount = 0
    for rdata,vals in raw_load.iterrows():
        if rcount == 0:
            for rrans in vals:
                if ('Q4_1' in rrans) == True:
                    rels.append(1)
                else:
                    rels.append(0)
        else:
            ans.append(vals)
        rcount+=1

    # リストrelsにより関連する行のあぶり出しができたのでこのインデックス数を取得する
    index_num = [n for n,v in enumerate(rels) if v == 1]

    # laをforループさせたときに分離させたans（純回答データ）から
    # index_numを当てはめた引数変数を取得し、Q7の回答のみのリストを作成する
    this_ans = []
    this_ids = []
    ids_count = 1
    for this in ans:
        this_ans.append(this[index_num[0]])
        this_ids.append(ids_count)
        ids_count+=1


    # jinja2テンプレート出力
    env = Environment(loader=FileSystemLoader('./', encoding='utf8'))
    tmpl = env.get_template('jinja2_templetes/templetes_Q4_chart_third.tmpl')
    # 商品情報を入れるリストを宣言、初期化
    items = []
    # ループで辞書型配列に置き換え、itemsにappendする
    # ti:ID,ta:回答データ
    for ti,ta in zip(this_ids,this_ans):
        try:
            if math.isnan(ta):
                pass
        except:
            items.append({'ID':ti,'ANSER':ta.replace("\u3000","")})

    # ここで shop に入る文字を指定している
    html = tmpl.render({'title':l_title,'item':items})
    with open('jinja2_templetes/templetes_Q4_chart_third.html',mode='w') as f:
        f.write(str(html))

    # seleniumでブラウザ表示（動作問題なし）
    options = Options()
    options.add_argument('--headless')
    driver = webdriver.Chrome(options=options)
    driver.get('file:///Users/t_sasaki/Documents/%E6%A5%AD%E5%8B%99/%E3%82%A2%E3%83%B3%E3%82%B1%E3%83%BC%E3%83%88/%E3%82%AB%E3%83%9F%E3%82%AA%E3%83%B3/macro_mk/jinja2_templetes/templetes_Q4_chart_third.html')
    page_width = driver.execute_script('return document.body.scrollWidth')
    page_height = driver.execute_script('return document.body.scrollHeight')
    driver.set_window_size(page_width,page_height)
    driver.save_screenshot('save_images/Q4_chart_third.png')
    driver.quit()


if __name__ in '__main__':
    Q4_chart_third()