#建立500x300的視窗
from tkinter import *
import threading
import 爬公開資訊股東會資料
aaa=open('公司代號對照表.csv','r',encoding="utf-8")
bbb=aaa.read().split("\n")
公司清單與代號={}
for i in range(965):
    公司清單與代號[f'{bbb[i][:4]}']=bbb[i][4:]
#print(公司清單與代號)

def pr():
    轉集合=set(公司代號.get().split(","))
    #print(轉集合)
    轉回列表=list(轉集合)
    #print(轉回列表)
    代號=",".join(轉回列表)
    #print(代號)
    
    爬公開資訊股東會資料.啟動(代號)
    
        
    
def newtask():
    
    l=len(公司代號.get())
    if l<4:
        num1.set("代號有四碼")
        return ""
    elif l==4:
        try:
            公司清單與代號[公司代號.get()]
        except:
            num1.set("請輸入正確代號")
            return ""
    if 公司代號.get().replace(",","").isdigit() or 公司代號.get().replace(".","").isdigit() :
        t = threading.Thread(target=pr)## 建立一個子執行緒
        t.start()# 執行該子執行緒
    
    else:
        num1.set("請輸入正確代號")
        return ""

    

window = Tk() #呼叫TK( )函式建立視窗，T大寫，k小寫
window.title('股東會資訊爬蟲')#視窗標題
window.geometry('320x50')#寬*高

label=Label(window,text='請輸入公司代號:',font=14)#(顯示視窗,文字,顏色,字型與大小,文字語塞框之間的距離x,y)

label.pack(side='left')
num1=IntVar()
num1.set("請輸入代號")#將文字盒1設定預設文字
公司代號=Entry(window, width=20, textvariable=num1)#文字盒1
公司代號.pack(side='left')
btn=Button(window, width=5, text='go',command=newtask)#按鈕直接呼叫複雜的函數會死機，所以先呼叫平行執行的模組

btn.pack(side='left')
window.mainloop() #呼叫mainloop( )函式讓程式運作直到關閉視窗為止
#pyinstaller -w -F --icon=avwpt-ykoq7-002.ico 查詢股票介面.py