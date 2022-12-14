
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
print(Fore.GREEN +"{:=^100s}".format("群旭CNC_報價單查詢小幫手Ver1.1"))
print('')

print('檔案讀取中', end = '')
for i in range(6):
    print(".",end = '',flush = True)  #flush - 输出是否被缓存通常决定于 file，但如果 flush 关键字参数为 True，流会被强制刷新
    time.sleep(0.5)


print('')


all_path = [] #讀取路徑後，暫存放路徑清單
r_path = [] #重複檔案路徑名
a_xlsx = [] #所有資料夾有xlsx檔案的路徑清單
n_xlsx = [] #所有資料夾且無重複xlsx檔案的路徑清單
r_xlsx = [] #有重複xlsx檔案的路徑清單
r_path = [] #重複檔案路徑名
log = [] #搜尋紀錄

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
print('')


#所有資料夾且無重複xlsx檔案的路徑清單,也已轉換斜線 (讀取檔案用)
newn_xlsx = []
for n in n_xlsx :
	nn = n.replace('\\','/')#把n_xlsx內所有路徑斜線轉換
	newn_xlsx.append(nn)



#簡易登入

print('載入完成')
print('')

while True:
	ans = '5931'
	x = 3 #初始機會
	while x > 0 :
		x = x -1
		pwd = input('請輸入登入密碼: ')
		if pwd == ans:
			print('')
			print('')
			break
		else:
			print('密碼錯誤!')
			if x > 0:
				print('還有', x,'次機會')
			else:
				print('已輸入超過三次，程式結束')
				print('')
				print('')
				print('3秒後程式關閉', end = '')
				for i in range(6):
					print("",end = '',flush = True)  #flush - 输出是否被缓存通常决定于 file，但如果 flush 关键字参数为 True，流会被强制刷新
					time.sleep(0.5)
					print('')
				os._exit()
	break
os.system('cls') #登入後清除畫面


print('')
print(Fore.GREEN +"{:=^100s}".format("群旭CNC_報價單查詢小幫手Ver1.1"))
print('登入成功')
print('')


#用戶自己輸入搜尋相關字之xlsx檔案
#搜尋本體
while True:
	answer = [] #用戶輸入的關鍵字過濾所有的xlsx後,確認有在xlsx內的清單
	name = input('請輸入要搜尋的關鍵字(大小寫有別) :')

	


	#讀取效果	
	conversation = ['正在執行一秒幾十萬上下的讀取，請稍後','莫急莫慌莫害怕，等等就好了','已經加班在趕了，等等', '痾‧‧‧忘記剛剛找到哪了，重找一下',
	'檔案有點多，你知道嗎','別這樣，大家都是操過來的,等等','這個需求不難，很快就', '剛剛好像夢到我找完了', '找出來的東西都是血跟淚換來的', 
	'認真找起來，連我自己都會怕', '快好了，稍等一下','大哥，你這要找的也太多了吧','群旭CNC部門是最棒低','終於肯來找我幫忙了喔', '謝謝你的耐心等候','想當年找的速度超快，但現在可能要等會',
	'東西太多翻得有點亂了欸，糟糕','你...是不是在找出口','是不是...報太低了呢','已經開三班幫你找了，等等','是不是又遇到了點麻煩呢','正在從深淵搜尋中','這是小Case,真的',
	'這資料藏得有點深，要花點時間','相信我真的不會很隨便找，真的','已經在找了，別急','正在裝模作樣尋找中','是不是跟著點點點呢','喝個水，等等就好了','正在執行一天大概只有三下的讀取，慢慢等吧',
	'等得有點無聊嗎，我也沒辦法','糟糕看到不該看的東西了，可以手下留情嗎','沒有人在處理中，請稍後，', '你好，已依照您的指示 【調高薪資】 ，正確請按 1 ', '別用這眼神看我，已經在找了', 
	'已經100%全力找了，看～快到連手都看不見了','警告! 正在刪除所有檔案，請稍後','正在從總經理的資料夾搜尋中','不要羨慕我找東西的能力，真的','正在塗改檔案中，反正你也不知道數字對不對', 
	'只給' + str(name) + '這幾個字，找到都不知道民國幾年了欸','正在為你搜尋公司的內部機密中','現在流行只給' + str(name) + '這幾個字，就要給你全世界？', '資料讀取中，請稍後','馬上找給你，請等會'
	'早就知道你要找有關' + str(name) + '的東西了，放心吧，都丟了','要很認真找還是大概就可以了呢']
	ha = random.choice(conversation)

	#讀取效果
	print('')
	ha = random.choice(conversation)
	print(ha, end = '')
	for i in range(4):
	    print(".",end = '',flush = True)
	    time.sleep(0.5)
	print('')
	

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



	count = 1 #報價單計數
	for ans_info in answer:
		copy_temp = []#為了顯示後可複製用
		copy_xlsx = []#為了顯示後可複製用

		df_ok = pd.read_excel(ans_info, index_col = 0,skiprows = 11) #sheet_name =None 會變成dic字典
		df_ok.fillna('',inplace = True)#將nan值取代
		df_ok.dropna(axis = 0, how = 'all', inplace = True )# 刪除空行
		final = df_ok[['品項','單價','數量','費用']]
		print("")
		print("")
		print(Fore.CYAN +"{:=^100s}".format(""))
		print(final.to_string(index=False)) #dataframe 輸出不加索引號
		print("")

		#路徑複製轉換
		nn = ans_info.replace('//','\\\\')#路徑斜線轉換
		copy_temp.append(nn)
		for copy2 in copy_temp:
			mm = copy2.replace('/','\\') #把路徑斜線轉換
			copy_xlsx.append(mm)



		print('此是第',Fore.RED + str(count), '張報價單，為',Fore.YELLOW + str(mm), '的檔案內容')
		count = count + 1
		print("")
	

	#搜尋紀錄LOG	
	with open('//192.168.10.61/f2-cnc報表/other/po_log.csv', 'r', encoding = 'cp950') as f:
		for txtlog in f:
			log.append(str(txtlog))


	with open('//192.168.10.61/f2-cnc報表/other/po_log.csv', 'w', encoding = 'cp950') as f:
		for l in log:
			f.write(str(l))
		f.write(str(name) + '\n')




	print(Fore.CYAN +"{:=^100s}".format(""))
	print("")
	print('在剛剛的超激烈讀取中，在', Fore.GREEN + str(project_dir) ,'資料夾內的', Fore.RED + str(len(all_path)), '個檔案中，')
	print("")
	print('找到有', Fore.RED + str(len(answer)), '張報價單(.xlsx)內包含有', Fore.RED + str(name), '的關鍵字')
	print("")
