# -*- coding: utf-8 -*-
"""
Created on Tue Dec 20 12:07:31 2022
程式功能：泰博水果餐盒自動登記系統

@author: dar

參考網頁：https://www.plus2net.com/python/pygsheets.php
需安裝 pip install pygsheets
"""
import os
import pygsheets
import requests
#==============================================================================
week_url = [
  '周一','周二','周三','周四','周五'  
]
#==============================================================================
#==============================================================================
def lineNotifyMessage(token, msg):
    headers = {
        "Authorization": "Bearer " + token, 
        "Content-Type" : "application/x-www-form-urlencoded"
    }

    payload = {'message': msg }
    r = requests.post("https://notify-api.line.me/api/notify", headers = headers, params = payload)
    return r.status_code
  
#==============================================================================
if __name__ == "__main__":
    token = "xxxxxx"  # 如果有要用 LINE 通知的話
    message = "本週訂到的午餐編號為 "    
    #===================================================================
    #強制變更工作目錄
    # Get the current working directory
    cwd = os.getcwd()
    # Print the current working directory
    print("Current working directory: {0}".format(cwd))    

    # Change the current working directory
    os.chdir('C:\\CodeBank\\Gsheets')     #若執行路徑有問題，才需要使用

    # Print the current working directory
    print("Current working directory: {0}".format(os.getcwd()))
    #===================================================================    
        
    gc = pygsheets.authorize(service_file='./dars-project-1220-42e1f8b20536.json')
    # 利用 Python 開啟 GoogleSheet
    sht = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gu2r3CSj3awGjhR9SKB-a2oM_AtKkw5QhK5U_d-dxVU/edit#gid=1089820924") #泰博午餐水果訂購 SHEET
    
    # 查看此 GoogleSheet 內 Sheet 清單
    wks_list = sht.worksheets()
    print(wks_list)
    
    # 選取要 Sheet 清單
    for j in range(5):
        wks = sht.worksheet_by_title(week_url[j]) #找到 周一 ~ 週五
        print(wks.title)
        message += str(wks.title)+" : "
    #讀取 df 也可以這樣寫
        df1 = wks.get_as_df() #將 GOOGLE 表單存到記憶體中(df1)
    #===================================================================    
        #讀取表格內容並判斷有沒有人填入工號
        for i in range(51):
            if i>=25:
                i+=1
            Work_NO = df1.iloc[(1+i,1)] # B3 工號那一格 改用 DF 效率更快
            name = df1.iloc[(1+i,2)] # C3 姓名那一格
            print("\r工號的值==>",Work_NO,end='')
            # print("工號的值==>",Work_NO)
            # print("姓名：",name)

            if Work_NO == "":  #如果沒有工號，就自動填我的工號
                wks.update_value((3+i,2),'XXX') # 配合原始表格內容[位置]更新其工號資料 
                wks.update_value((3+i,3),'NNN') # 姓名資料
                wkno = wks.cell((3+i,1)) # A3 編號那一格
                print("\n取得的號碼是:",wkno.value)
                message += str(wkno.value)+" ; "
                break

            # 因為做表格的人弄了二排，所以只好再分析另一排
            Work_NO = df1.iloc[(1+i,5)] # F3 工號那一格 改用 DF 效率更快
            name = df1.iloc[(1+i,6)] # G3 姓名那一格
            print("\r工號的值==>",Work_NO,end='')
            # print("工號的值==>",Work_NO)
            # print("姓名：",name)

            if Work_NO == "":
                wks.update_value((3+i,6),'XXX') # 配合原始表格內容[位置]更新其工號資料 
                wks.update_value((3+i,7),'NNN') # 姓名資料
                wkno = wks.cell((3+i,5)) # E3 編號那一格
                print("\n取得的號碼是:",wkno.value)
                message += str(wkno.value)+" ; "                
                break
    #===================================================================
    print("更新完成")

    print(message) # 回傳訂單的編號

    # lineNotifyMessage(token, message) # 如果要傳 LINE 的話，再加上
