#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Oct  9 02:52:50 2020

@author: developer
"""

import os
from retrying import retry
from bs4 import BeautifulSoup
import re
import xlrd
import requests
from requests_html import HTMLSession
import datetime
from PIL import Image

#=============================================================================
@retry(stop_max_attempt_number=10)

# 文字列結合関数
def join_string(str_input):
    return '<div style="text-align:center;">' + str(str_input) + '</div>'

# 画像サイズ変換（300:1024）
def resize(str_input):
    #　widthを抽出する
    pattern = r'width="[0-9]+'
    org_width = re.search(pattern, str(str_input))
    #print(org_width.group())   
    # width変換
    tmp = org_width.group().split('="')
    new_width = int(int(tmp[1]) * (1024 / 300))     #★ココでサイズ変更規則を変更できる
    #print(tmp[1])
    #print(new_width)   
    # 変換後widthに置き換え
    return_text = re.sub(pattern, 'width="' + str(new_width) , str(str_input))
    return return_text   

# テキストから挿入画像情報を取得
def get_images_list(file):
    with open(file, "r", encoding="utf_8_sig") as f:
        # テキスト から BeautifulSoup オブジェクトを作る
        soup = BeautifulSoup(f.read(), 'html.parser')
        # pタグの中身をリストに格納
        images_articles = soup.find_all('p')
        
        images_list = []
        
        # 記事ごとの画像リストを二次元配列に格納(image_list作成)
        for i in images_articles:
            # imgタグの中身をリストに格納           
            images_list.append(i.find_all('img'))           
    return images_list

# =============================================================================
#     # リストのすべての要素に<div style="text-align:center;"> と </div> を追加
#     #　map関数ですべてに上の文字列追加
#     tmp = list(map(join_string, images_list[0]))
#     print("aaa")
#     print(tmp)
#     #　map関数ですべてに300：1024で変換
#     #resize(images_list[0])
#     tmp = list(map(resize, tmp))
#     print("bbb")
#     print(tmp)
# 
# =============================================================================  

# 文字列からパターンに一致する文字列を削除する関数
def delete_text(text):
    pattern = r' class="[a-zA-Z0-9_-]*"'
    return_text = re.sub(pattern, '' , text)
    return return_text

# fuelle用テキスト作成関数
def get_content_for_fuelle(url, output_path, images_list):   
    #Requestsを利用してWebページを取得する
    html = requests.get(url)
    # #htmlパーサー
    soup = BeautifulSoup(html.content, 'html.parser')
    # HTMLの内容を書き込むようの配列を用意
    result = []   
    
    # 抽出対象は、すべて<div class = "p-postDetail">の中なので、探索対象をココに絞る
    content = soup.find('div', class_='p-postDetail')
    
    # 挿入画像の枚数カウンタ
    image_num = 0
    
    # h2, h3, p　タグの要素をすべて抽出し、ループで必要なものを取り出す
    for i in content.find_all(['h2', 'h3','h4', 'p', 'div' , 'img', 'table', 'li', 'a']):
        #print(i)
        if i.name == 'h2':
            # 「class="文字列"」を削除
            i = delete_text(str(i))
            result.append(i)
        elif i.name == 'h3':
            # 「class="文字列"」を削除
            i = delete_text(str(i))
            result.append(i)
        elif i.name == 'h4':
            # 「class="文字列"」を削除
            i = delete_text(str(i))
            result.append(i)
        elif i.name =='table':
            # 「class="文字列"」を削除
            i = delete_text(str(i))
            result.append(i)
        elif i.name =='li':
            # 「class="文字列"」を削除
            i = delete_text(str(i))
            result.append(i)
        elif i.name =='a':
            if(i.get('class') != None):
                # 画像右下の文字を追加
                if i.get('class')[0] == 'p-postDetail__imageQuote' :
                    # resultリストの末尾の要素を取得
                    result_tail = result[-1]
                    # str型⇒bs4.BeautifulSoup型に変換（htmlとして扱ってfindで<img>タブを検索するため）
                    soup_result_tail = BeautifulSoup(result_tail)
                    #　resultの末尾の要素を削除
                    del result[-1]
                    # 取り出した末尾の要素に情報を付加して、resultに追加                
                    result.append('<div style="text-align:center;"><figure class="image">'+ str(soup_result_tail.find('img')) + '<figcaption class="wi-caption-text">' + i.text + '</figcaption></figure></div>')
                    
        elif i.name == 'p':
            # 本文に「記事」というワードが入っている場合は除外
            if('記事' in str(i)):
                pass    #何もせずに抜ける
            elif(i.get('class') != None):
                # 広告（RELATED ARTICLE）は除外
                if(i.get('class')[0] != 'p-postDetail__relatedHeading' and i.get('class')[0] != 'p-postDetail__relatedTitle'):
                    # 「class="文字列"」を削除
                    i = delete_text(str(i))
                    result.append(i)
        # 写真を挿入する部分
        elif i.name == 'img':
            # 広告（RELATED ARTICLE）は除外
            if(i.get('alt') != 'RELATED ARTICLE'):
                result.append(images_list[image_num])
                image_num = image_num + 1
                #result.append(i)      
        elif i.name == 'div':
            if(i.get('class') != None):
                if(i.get('class')[0] == 'p-postDetail__productName'):
                    i = str(i.text).replace(' ', '')    # スペースを削除
                    i = i.replace(chr(165), '&yen;')    # yenマークを"&yen;"に置換  （文字コードは、chr(95)もあり）
                    i = i.replace('\n','')              # Unix系
                    i = i.replace('\r','')              # MacOS(9以前)
                    i = i.replace('\r\n','')            # Windows系
                    result.append('<p>' + i + '</p>')
    # end of [for i in content.find_all(['h2', 'h3','h4', 'p', 'div' , 'img', 'table']):]
                    
    # ======
    # HTMLの内容をテキストに出力
    with open('../input/html.txt', 'w',encoding="utf_8_sig") as f:
        for i in result:
            f.write("%s\n\n" % i)
    
   # print(match.group())
    with open(output_path, "w", encoding="utf_8_sig") as outfile:
    #with open(output_path, "w", encoding="utf_8_sig") as outfile:
        with open('../input/html.txt', "r", encoding="utf_8_sig") as infile:
                  outfile.write(infile.read())
        with open('../input/footer.txt', "r", encoding="utf_8_sig") as infile:
                  outfile.write(infile.read())        
   # =======
##### def get_content_for_fuelle(url): #####

# # FLOWER用テキスト作成関数
def get_content_for_FLOWER(url, output_path):   
    #Requestsを利用してWebページを取得する
    html = requests.get(url)
    # #htmlパーサー
    soup = BeautifulSoup(html.content, 'html.parser')
    # HTMLの内容を書き込むようの配列を用意
    result = []   
      
    # 抽出対象は、すべて<div class = "p-postDetail">の中なので、探索対象をココに絞る
    content = soup.find('div', class_='p-postDetail')
    
    # h2, h3, p　タグの要素をすべて抽出し、ループで必要なものを取り出す
    for i in content.find_all(['h2', 'h3','h4', 'p', 'div' , 'img', 'table', 'li']):
        #print(i.text)
        if i.name == 'h2':
            # 「class="文字列"」を削除
            i = delete_text(str(i.text))
            result.append('## ' + i)
        elif i.name == 'h3':
            # 「class="文字列"」を削除
            i = delete_text(str(i.text))
            result.append('### ' + i)
        elif i.name == 'h4':
            # 「class="文字列"」を削除
            i = delete_text(str(i.text))
            result.append(i)
        elif i.name =='table':
            
            # インデント揃えるように最初に文字数をカウントする
            #============================================
            # テーブル内の列番号をカウント。列ごとに文字数をカウント。
            count = 0
            str_max_len = [0] * 10 
            for tr_tab_content in i.find_all('tr'):
                # <tr>内を一列に結合（全角スペースで区切る）
                tr = ''
                # テーブル内の列番号をカウント。列ごとに文字数をカウント。
                count = 0
                for td_tab_content in tr_tab_content.find_all('td'):
                    # 文字数カウントし、列ごとに最も長い文字数を保存
                    if(str_max_len[count] < len(td_tab_content.text)):
                        str_max_len[count] = len(td_tab_content.text)
                    count += 1
           #============================================ 
                        
            # tableの中の<tr>タブの中身を抽出
            for tr_tab_content in i.find_all('tr'):
                # <tr>内を一列に結合（全角スペースで区切る）
                tr = ''
                # テーブル内の列番号をカウント。列ごとに文字数をカウントしてインデント揃える
                count = 0
                for td_tab_content in tr_tab_content.find_all('td'):
                    # 文字数をカウントして最大文字列に合わせて全角スペースを挿入して、インデントを揃える
                    #if(0 < str_max_len[count] - len(td_tab_content)):
                    space = ''
                    for j in range(str_max_len[count] - len(td_tab_content.text)):
                        space = space + '　'
                    # 結合（全角スペースで区切る）
                    tr = tr + '　' + td_tab_content.text + space
                    count += 1
                # 末尾の全角スペースを削除
                tr = tr.rstrip('　')
                result.append(tr)
        elif i.name =='li':
            # 「class="文字列"」を削除
            i = delete_text(str(i.text))
            result.append(i)
        elif i.name == 'p':
            # 本文に「記事」というワードが入っている場合は除外
            if('記事' in str(i)):
                pass    #何もせずに抜ける
            elif(i.get('class') != None):
                # 広告（RELATED ARTICLE）は除外
                if(i.get('class')[0] != 'p-postDetail__relatedHeading' and i.get('class')[0] != 'p-postDetail__relatedTitle'):
                    # 「class="文字列"」を削除
                    i = delete_text(str(i.text))
                    result.append(i)
        elif i.name == 'div':
            if(i.get('class') != None):
                if(i.get('class')[0] == 'p-postDetail__productName'):
                    i = str(i.text).replace(' ', '')    # スペースを削除
                    result.append(i)
    # end of [for i in content.find_all(['h2', 'h3','h4', 'p', 'div' , 'img', 'table']):]
                    
    # =================================================  
    # HTMLの内容をテキストに出力
    with open('../input/html.txt', 'w',encoding="utf_8_sig") as f:
        for i in result:
            f.write("%s\n\n" % i)
    with open(output_path, "w", encoding="utf_8_sig") as outfile:
        with open('../input/html.txt', "r", encoding="utf_8_sig") as infile:
                  outfile.write(infile.read())
        with open('../input/footer_for_FLOWER.txt', "r", encoding="utf_8_sig") as infile:
                  outfile.write(infile.read())        
   # =================================================
#====end of [def get_content_for_FLOWER(url, output_path):]=======



# Excelファイル読み込み
def read_excel(file_path):
    xl_bk = xlrd.open_workbook(file_path)
    xl_sh = xl_bk.sheet_by_index(0)
     
    ary=[[None for p in range(xl_sh.ncols)] for q in range(xl_sh.nrows)]
     
    for y in range(xl_sh.nrows):
        for x in range(xl_sh.ncols):
            ary[y][x]=xl_sh.cell(y,x).value
    return ary

def get_all_content(soup):
    # soupから必要な情報のあるタグの中身全体をcontentにコピー    
    content = soup.find('div', class_='p-postDetail')
    return content


def dir_m(path, num):
    if(os.path.isdir(path + str(num).zfill(2))):
        num = num + 1
        num = dir_m(path, num)
    else:
        path = path + str(num).zfill(2)
        os.mkdir(path)
    return num


def excel_date(date):
    from datetime import datetime, timedelta
    return(datetime(1899, 12, 30) + timedelta(days=date)).strftime("%Y%m%d")

def resize80(file_path,x):
    s = os.path.getsize(file_path)
    if s >= 80000:
        
        re_img = Image.open(file_path)
 
        width, height = re_img.size
        
        if width > height:
            resize_img = re_img.resize((x,int(height*(x/width))))
        else:
            resize_img = re_img.resize((int(width*(x/height)),x))
        
        
        try:
            resize_img.save(file_path)
        except:
            resize_img = resize_img.convert('RGB') 
            resize_img.save(file_path, "JPEG", quality=95)

        resize80(file_path,(x-50))

def make_fuelle_thmb(fuelle_pil_path):
    pil_img = Image.open(fuelle_pil_path)
    width, height = pil_img.size
    if width == height:
        repil_img = pil_img.resize((800,800))
        result = Image.new('RGB', (1200, 800), (0,0,0))
        result.paste(repil_img, (200, 0))
        return result
    else:
        print ("サムネイルエラー！！！")
        
def make_FLOWER_thmb(FLOWER_pil_path):
    pil_img = Image.open(FLOWER_pil_path)
    width, height = pil_img.size
    if width == height:
        repil_img = pil_img.resize((400,400))
        result = Image.new('RGB', (600, 400), (0,0,0))
        result.paste(repil_img, (100, 0))
        return result
    else:
        print ("サムネイルエラー！！！")

#=============================================================================


if __name__ == '__main__':

    # 現在時刻取得
    now = datetime.datetime.now()
    i = 0
    
    # IOファイルのパス
    #input_path = '/Users/developer/development/blog_RPA/blog_data.xlsx'   
    #output_path = '/Users/developer/Documents/転載/img'
    input_path = '../input/blog_data.xlsx'
    output_path = '../output'
    
    # 入力ファイル読み込み
    input_data = read_excel(input_path)

    # 入力ファイルを1行ずつ取り出す
    for row in input_data:
        if row[0] != 'code':
            #初期化
            img_url_list = []
            session = HTMLSession()
            num = 1

            # セルの中身を取り出す
            input_date = excel_date(row[2])
            input_time = row[4]
            input_title = row[11]
            input_url = row[16]
            dir_path = output_path + '/' + input_date
                  
            # 保存先フォルダ作成
            new_num = dir_m(dir_path, num)
            

            
            #Requestsを利用してWebページを取得する
            html = requests.get(input_url)
            # #htmlパーサー
            soup = BeautifulSoup(html.content, 'html.parser')
            
            j = 0
            
            # ★テキスト作成★
            # ================================================================================
            
            images_list = get_images_list('../input/images_list.txt')
            #　map関数ですべてに上の文字列追加
            #print("============================")
            #print(new_num)
            #print(images_list[j])
            tmp = list(map(join_string, images_list[new_num-1]))
            #　map関数ですべてに300：1024で変換
            new_images_list = list(map(resize, tmp))

            # fuelle用テキスト作成
            get_content_for_fuelle(input_url, dir_path + str(new_num).zfill(2) + '/content_for_fuelle.txt', new_images_list)   
            # 画像差し替え（サイズ変換・センタリング）
            
            
            # FLOWER用テキスト作成
            get_content_for_FLOWER(input_url, dir_path + str(new_num).zfill(2) + '/content_for_FLOWER.txt')
            # ================================================================================
            
            j = j + 1   
            # サムネイル     
"""

            content_icon = soup.find('div', class_='p-postBox__eyecatchImgWrapper')
            icon_tag = content_icon.find('img')
            if icon_tag:
                img_url_list.append(icon_tag.get('src'))
                
            else:
                j = j + 1
            
 
            # 各種画像
            content = soup.find('div', class_='p-postDetail')#p-postBox__eyecatchImgWrapper
            img_tag_list = content.find_all('img')

            for img_tag in img_tag_list:
                if not 'RELATED ARTICLE' in str(img_tag.get('alt')):
                    img_url_list.append(img_tag.get('src'))

            for img_url in img_url_list:
                if img_url.startswith('//'): #//で始まってるURLの場合
                    img_url = 'https:' + img_url
                
                file_path = dir_path + str(new_num).zfill(2) + '/' + str(j) + '.jpg'
                print (num)
                print (new_num)
                
                # 画像取得
                img = requests.get(img_url)
                with open(file_path, "wb") as f:
                    f.write(img.content) 
                    
                ################################################                
                if j == 0:
                    print ()
                    fuelle_thmb = make_fuelle_thmb(file_path)
                    FLOWER_thmb = make_FLOWER_thmb(file_path)
                                                
                    fuelle_path = dir_path + str(new_num).zfill(2) + '/fuelle.jpg'
                    #FLOWER_path = dir_path + str(new_num).zfill(2) + '/FLOWER' + str(new_num) + '.jpg'
                    FLOWER_path = dir_path + str(new_num).zfill(2) + '/FLOWER' + input_date + '.jpg'
                    
                    fuelle_thmb.save(fuelle_path)
                    FLOWER_thmb.save(FLOWER_path)
                else:
                    if os.path.isfile(file_path):
                        resize80(file_path, 500)
               #################################################             
            os.rename(dir_path + str(new_num).zfill(2) + '/0.jpg',dir_path + str(new_num).zfill(2) + '/bk' + str(j) + '.jpg')  
            #　====end of [for img_url in img_url_list:]=====   
"""  
 