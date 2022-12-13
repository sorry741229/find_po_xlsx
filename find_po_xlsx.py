import pandas as pd
import os, sys
from colorama import init, Fore, Back #字體顏色
init(autoreset=True)#字體顏色

#https://clay-atlas.com/blog/2020/03/22/python-%E8%BC%B8%E5%87%BA%E5%AD%97%E4%B8%B2%E5%9C%A8%E7%B5%82%E7%AB%AF%E6%A9%9F%E4%B8%AD%E9%A1%AF%E7%A4%BA%E9%A1%8F%E8%89%B2/
#調整視窗大小
from ctypes import windll, byref
from ctypes.wintypes import SMALL_RECT
import time
import random

WindowsSTDOUT = windll.kernel32.GetStdHandle(-11)
dimensions = SMALL_RECT(-10, -10, 120, 40) # (left, top, right, bottom)
# Width = (Right - Left) + 1; Height = (Bottom - Top) + 1
windll.kernel32.SetConsoleWindowInfo(WindowsSTDOUT, True, byref(dimensions))

pd.set_option("display.max_rows", None)    #設定最大能顯示1000rows
pd.set_option("display.max_columns", None) #設定最大能顯示1000columns
pd.set_option('display.width', 1000) # 設置打印寬度 
pd.set_option('display.max_colwidth', 180)
pd.set_option("display.colheader_justify","center") #抬頭對齊用
pd.set_option('display.unicode.ambiguous_as_wide', True) #抬頭對齊用
pd.set_option('display.unicode.east_asian_width', True) #抬頭對齊用



#讀取路徑
#project_dir = 'C:/Users/cnc-3/Desktop/新增資料夾3' 
project_dir = '//192.168.10.61/f2-cnc報表/客戶管理資料/圖'
a = os.walk(project_dir) 

print('')

print(Fore.GREEN +"{:=^100s}".format("群旭CNC_報價單查詢小幫手"))
print('')

print('檔案讀取中', end = '')
for i in range(5):
    print(".",end = '',flush = True)  #flush - 输出是否被缓存通常决定于 file，但如果 flush 关键字参数为 True，流会被强制刷新
    time.sleep(0.5)

#讀取效果	
conversation = ['正在執行一秒幾十萬上下的讀取，請稍後','莫急莫慌莫害怕，等等就好了','已經加班在趕了，等等', '痾‧‧‧忘記剛剛找到哪了，重找一下',
	'檔案有點多，你知道嗎','別看我這樣，我也是操過來的‧‧‧等等好嗎','這個需求不難，很快就', '剛剛好像夢到我找完了', '等等找給你的東西都是血跟淚換來的', 
	'認真找起來，連我自己都會怕', '快好了，稍等一下','大哥，這也太多了吧','群旭CNC部門是最棒低','終於肯來找我幫忙了喔', '謝謝你的耐心等候']
ha = random.choice(conversation)

print('')



all_path = [] #讀取路徑後，暫存放路徑清單
repeat = [] #重複檔案名
r_path = [] #重複檔案路徑名
a_xlsx = [] #所有資料夾有xlsx檔案的路徑清單
n_xlsx = [] #所有資料夾且無重複xlsx檔案的路徑清單
r_xlsx = [] #有重複xlsx檔案的路徑清單
r_path = [] #重複檔案路徑名

#列出a內裡所有檔案的路徑
# root  查找資料夾的路徑
# dirs  查找路徑內資料夾內子資料夾[清單]
# files 查找資料夾內所有檔案[清單] 資料夾不會列出

for root, dirs, files in a :
	for f in files :
		path = os.path.join(root, f) #完整路徑all_path
		all_path.append(path)

for a in all_path :
	if '.xlsx' in a :
		a_xlsx.append(a)
	else:
		continue

for check in a_xlsx :
	if check not in n_xlsx :
		n_xlsx.append(check)
	else:
		r_xlsx.append(check)			


print('')
print(Fore.YELLOW +"{:=^100s}".format(""))
print(Fore.YELLOW +"{:=^100s}".format(""))
print('')



#所有資料夾且無重複xlsx檔案的路徑清單,也已轉換斜線
newn_xlsx = []
for n in n_xlsx :
	nn = n.replace('\\','/')#把n_xlsx內所有路徑斜線轉換
	newn_xlsx.append(nn)




#用戶自己輸入搜尋相關字之xlsx檔案
print('載入完成')
print('')
while True:
	answer = [] #用戶輸入的關鍵字過濾所有的xlsx後,確認有在xlsx內的清單
	name = input('請輸入要搜尋的關鍵字(大小寫有別) :')
	
	#讀取效果
	print('')
	print(ha, end = '')
	for i in range(10):
	    print(".",end = '',flush = True)
	    time.sleep(0.5)

	for xlsx in newn_xlsx:
		try:
			df= pd.read_excel(xlsx, index_col = 0,skiprows = 11) #讀取檔案，sheet_name =None 會變成dic字典
			df.dropna(axis = 0, how = 'all', inplace = True )# 刪除空行
			df = df.reset_index(drop=False) #重設索引
			df.fillna('',inplace = True)#將nan值取代
			#print(df)
			result = df['品項'].str.contains(name, na=False)#將dataframe轉成文字比對name,得到布林值
			filter_result = df[result]
			if filter_result.empty:
			 	continue
			else:
				answer.append(xlsx)
		except :
			pass

	count = 1
	for ans_info in answer:
		df_ok = pd.read_excel(ans_info, index_col = 0,skiprows = 11) #sheet_name =None 會變成dic字典
		df_ok.fillna('',inplace = True)#將nan值取代
		df_ok.dropna(axis = 0, how = 'all', inplace = True )# 刪除空行
		final = df_ok[['品項','單價','數量','費用']]
		print("")
		print("")
		print("")
		print(Fore.CYAN +"{:=^100s}".format(""))
		print(final.to_string(index=False)) #dataframe 輸出不加索引號
		print("")
		print('此上是第',Fore.RED + str(count), '張報價單為',Fore.YELLOW + str(ans_info), '的檔案')
		count = count + 1
		print("")
		print("")

	print(Fore.CYAN +"{:=^100s}".format(""))
	print("")
	print('在剛剛的超激烈讀取中，在', Fore.GREEN + str(project_dir) ,'資料夾內的', Fore.RED + str(len(all_path)), '個檔案中')
	print('找到有', Fore.RED + str(len(answer)), '張報價單(.xlsx)內包含有', Fore.RED + str(name), '的關鍵字')
	print("")
