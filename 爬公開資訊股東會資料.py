
"""
#邏輯概念:
公開資訊觀測站的股東會訊息是隱藏起來的表格(觀察html程式碼，可以發現表格被設定為'hidden')
<input type="hidden" name="colorchg" value="">

當按下詳細資料的按鈕後，會將按鈕包含的'onclick'屬性中的所有值的資料傳給隱藏的表格，
<input type="button" value="詳細資料" 
        onclick="document.fm.CAL.value="";
                document.fm.SE1.value="";
                document.fm.DA1.value="";
                document.fm.DATE1.value="20220504";   <<<<<傳入的資料，此資料不同公司傳入的日期都不同，所以需要先爬取此處資料才知道要傳入甚麼值
                document.fm.SEQ_NO.value="1";         <<<<<傳入的資料，固定傳入的值
                document.fm.COMP.value="2002";        <<<<<傳入的資料，傳入公司代號作為值
        openWindow(this.form ,"");">                  <<<<<新視窗中顯示隱藏的資料
給入.DATE1/.SEQ_NO/.COMP的值後，才會跳出新視窗顯示隱藏中的表格
"""
import requests,bs4
from selenium import webdriver
import time,sys,time
import pandas as pd
headers = {"user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.67 Safari/537.36"   }
url = "https://mops.twse.com.tw/mops/web/ajax_t108sb16"


def 查找關鍵字(s,first,last):
    try:
        start=s.index(first)+len(first)
        end=s.index(last,start)
        return s[start:end]
    except ValueError:
        return ""

    
def 單間公司(child_co_id):
    driver=webdriver.Chrome('chromedriver.exe')
    try:#嘗試在瀏覽器中打開網址
        driver.get("https://mops.twse.com.tw/mops/web/t108sb16_q1")
        time.sleep(1)
        #print(driver.page_source)#網頁程式碼
    except:#如郭發生錯誤，比如網址打不開
        driver.quit()#關閉瀏覽器
        sys.exit("查無網址，程式中止")#不跑後面的程式碼，直接離開，印出指定文字做說明

    try:#嘗試將代號丟進去搜尋
        id_name=driver.find_element_by_id("co_id")#定位搜尋的文字盒
        id_name.clear()#清空文字和原本的文字
        id_name.send_keys(child_co_id)#傳入代號
        serach_button=driver.find_element_by_css_selector("#search_bar1 > div > input[type=button]")#定位搜尋按鈕
        serach_button.click()#點擊(click)搜尋按鈕
        time.sleep(2)#延遲一秒，防止網路太慢搜尋介面還沒跳出來
        r=driver.find_element_by_css_selector("#table01 > center > form > table.hasBorder > tbody > tr.even > td:nth-child(5) > input[type=button]")#定位詳細資料按鈕
        time.sleep(2)#延遲一秒，防止網路太慢搜尋介面還沒跳出來
        屬性值=r.get_attribute("onclick")#取得指定屬性的程式碼值
        屬性值=屬性值.split(";")[3][-9:-1]#取出給資料的日期
        driver.quit()
    except:#如果搜尋不到，取不到資料會出現錯誤
        driver.quit()#關閉瀏覽器
        sys.exit("查無此資料，程式中止")   

    return 屬性值


def 單間公司資訊(child_co_id,屬性值):
    #查看隱藏表格原本設定的屬性與基本值，並將顯示在新視窗的資料的值進行更改DATE1/COMP
    payload = {'colorchg':'',
                'TYPEK':'sii',
                'id':'',
                'DATE1':屬性值,
                'SEQ_NO':'1',
                'COMP':child_co_id,
                'kind':'',
                'SKIND':'A',
                'step':'2',
                'firstin':'1',
                'CAL':'',
                'DA1':'',
                'SE1':''}
    res = requests.post(url,data=payload, headers=headers).content#得到html程式碼
    soup = bs4.BeautifulSoup(res,"lxml")#轉bs4型態
    p_tag = soup.find_all('tr', {"style":"text-align:left !important;"})#爬取所有tr標籤.style的值為text-align:left !important;的所有文字
    資料字典={}
    資料字典["公司代號"]=[child_co_id]
    for i in range(len(p_tag)):#將所有文字的列表型態進行遍歷
        if p_tag[i].text.count("停止股票過戶起訖日期：")==1:
            資料字典["停止股票過戶起訖日期"]=[查找關鍵字(p_tag[i].text,"停止股票過戶起訖日期","日")[1:]]

        if p_tag[i].text.count("本次股東常會紀念品為：")==1:
            資料字典["紀念品"]=[查找關鍵字(p_tag[i].text,"本次股東常會紀念品為：","(一)")]
        else:
            continue   
        if p_tag[i].text.count("紀念品發放原則:")==1:
            資料字典["紀念品發放原則"]=[查找關鍵字(p_tag[i].text,"紀念品發放原則:","(").strip()]
        else:
            continue
    return pd.DataFrame(資料字典)#回傳隱藏視窗的所有資訊
    
def 多間公司(list):
    driver=webdriver.Chrome('chromedriver.exe')
    屬性值字典={}
    try:
        driver.get("https://mops.twse.com.tw/mops/web/t108sb16_q1")
    except:
        driver.quit()#關閉瀏覽器
        sys.exit("查無網址，程式中止")   
    for 公司代號 in list:#查詢所有公司的按鈕設定值的資料
        try:
            id_name=driver.find_element_by_id("co_id")
            id_name.clear()
            id_name.send_keys(公司代號)
            serach_button=driver.find_element_by_css_selector("#search_bar1 > div > input[type=button]")
            serach_button.click()
            time.sleep(1)
            r=driver.find_element_by_css_selector("#table01 > center > form > table.hasBorder > tbody > tr.even > td:nth-child(5) > input[type=button]")
            屬性值=r.get_attribute("onclick")#取得指定屬性的程式碼值
            屬性值字典[公司代號]=屬性值.split(";")[3][-9:-1]#取出給資料的日期
        except:
            #print(f"查無{公司代號}公司資料")#因為是多筆公司資料的搜尋，所以還不能離開瀏覽器
            continue#繼續執行查找下一間公司資料
    driver.quit()#全部查找完畢再關掉瀏覽器
    return 屬性值字典#回傳公司代號(key)/document.fm.DATE1.value="20220504"的值(value)的字典  
                    #ex: {"2002":"20220504","2330":"20220304"}


def 多間公司資訊(屬性值字典):
    a=[]
    for 公司代號 in 屬性值字典:
        payload = {'colorchg':'',
                    'TYPEK':'sii',
                    'id':'',
                    'DATE1':屬性值字典[公司代號],
                    'SEQ_NO':'1',
                    'COMP':公司代號,
                    'kind':'',
                    'SKIND':'A',
                    'step':'2',
                    'firstin':'1',
                    'CAL':'',
                    'DA1':'',
                    'SE1':''}
        res = requests.post(url,data=payload, headers=headers).content
        soup = bs4.BeautifulSoup(res, "lxml")
        p_tag = soup.find_all('tr', {"style":"text-align:left !important;"})
        
        資料字典={}
        資料字典["公司代號"]=公司代號
        for i in range(len(p_tag)):#將所有文字的列表型態進行遍歷
            if p_tag[i].text.count("停止股票過戶起訖日期：")==1:
                資料字典["停止股票過戶起訖日期"]=[查找關鍵字(p_tag[i].text,"停止股票過戶起訖日期","日")[1:]]
            
            if p_tag[i].text.count("本次股東常會紀念品為：")==1:
                資料字典["紀念品"]=查找關鍵字(p_tag[i].text,"本次股東常會紀念品為：","(一)")
            else:
                continue   

            if p_tag[i].text.count("紀念品發放原則:")==1:
                資料字典["紀念品發放原則"]=查找關鍵字(p_tag[i].text,"紀念品發放原則:","(").strip() 
            else:
                continue   
        a.append(資料字典)  
    return a


def 查詢方式(child_co_id):
    if child_co_id.count(",")>0:#逗號分隔多間公司
        公司數量=child_co_id.split(",")#字串用符號分解成列表
        unique_set = set(公司數量)#將列表轉為集合型態，此時重複的資料就會不見
        公司數量 = list(unique_set)#將集合轉回列表型態
        公司數量.sort()#根據公司代號由小到大排序
        屬性值字典=多間公司(公司數量)#取得所有公司股東會隱藏分頁要input的資料
        開會資料=多間公司資訊(屬性值字典)
        #with open("多間公司股東會資訊.txt","w") as f:
            #f.write(多間公司資訊(屬性值字典)) #將資料給入爬蟲，即可取得股東會的隱藏分頁內容，存入記事本
        return 開會資料

    elif child_co_id.count(".")>0:#點分隔多間公司
        公司數量=child_co_id.split(".")#字串用符號分解成列表
        unique_set = set(公司數量)#將列表轉為集合型態，此時重複的資料就會不見
        公司數量 = list(unique_set)#將集合轉回列表型態
        公司數量.sort()#根據公司代號由小到大排序
        屬性值字典=多間公司(公司數量)#取得所有公司股東會隱藏分頁要input的資料
        開會資料=多間公司資訊(屬性值字典)
        #with open("多間公司股東會資訊.txt","w") as f:
            #f.write(多間公司資訊(屬性值字典)) #將資料給入爬蟲，即可取得股東會的隱藏分頁內容，存入記事本
        return 開會資料
    else:
        
        屬性值=單間公司(child_co_id)
        開會資料=單間公司資訊(child_co_id,屬性值)
        #with open(f"{child_co_id}股東會資訊.txt","w") as f:
            #f.write(單間公司資訊(child_co_id,屬性值)) 
        return 開會資料

def 製作表格(開會資料,child_co_id):
    做表格={"公司代號":[],
            '停止股票過戶起訖日期':[],
            '紀念品':[],
            '紀念品發放原則':[]}
    if child_co_id.count(",")>0 or  child_co_id.count(".")>0:
        for i in range(len(開會資料)):
            做表格["公司代號"].append(開會資料[i]["公司代號"])
            做表格["停止股票過戶起訖日期"].append(開會資料[i]["停止股票過戶起訖日期"][0])
            try:
                做表格["紀念品"].append(開會資料[i]["紀念品"])
                做表格["紀念品發放原則"].append(開會資料[i]["紀念品發放原則"])
            except:
                做表格["紀念品"].append("None")
                做表格["紀念品發放原則"].append("None")
        return pd.DataFrame(做表格)
    else:
        return 開會資料
        #if 開會資料==False:
            #print("AAA")
       # else:
            #return pd.DataFrame(開會資料)


    

#child_co_id = input("請輸入公司代號(用',' or '.'區隔):")   # 從頁面上的表格獲取的子/母公司代號
def 啟動(child_co_id):
    開會資料=查詢方式(child_co_id)
    製作表格(開會資料,child_co_id).to_excel(f"{child_co_id}股東會資訊.xlsx")

    

#啟動("2002")