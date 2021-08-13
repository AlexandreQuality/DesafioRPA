import rpa as r
import pyautogui as aut
import pandas as pd

#--------------------------< Variaveis Globais >--------------------#

pastadoc = r"D:\Python\RPA\Udemy\docs"
pastaimg = r"D:\Python\RPA\Udemy\img"

#--------------------------<     RPA Start    >---------------------#

r.init()
janela = aut.getActiveWindow()
janela.maximize()
r.url('http://rpachallenge.com/')
aut.sleep(7)
r.download('http://rpachallenge.com/assets/downloadFiles/challenge.xlsx',fr'{pastadoc}\challenge.xlsx')
aut.sleep(2)

dados = pd.read_excel(rf'{pastadoc}\challenge.xlsx',sheet_name='Sheet1')
df = pd.DataFrame(dados,columns=["First Name","Last Name ",	"Company Name", "Role in Company",	"Address",	"Email","Phone Number"])

r.click('/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button')

for row in df.itertuples():
    r.type('//*[@ng-reflect-name="labelFirstName"]',  row[1])
    r.type('//*[@ng-reflect-name="labelLastName"]',   row[2])
    r.type('//*[@ng-reflect-name="labelCompanyName"]',row[3])
    r.type('//*[@ng-reflect-name="labelRole"]',       row[4])
    r.type('//*[@ng-reflect-name="labelAddress"]',    row[5])
    r.type('//*[@ng-reflect-name="labelEmail"]',      row[6])
    r.type('//*[@ng-reflect-name="labelPhone"]',      str(row[7]))
    r.click('/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input')
    
aut.sleep(5)
aut.screenshot(f'{pastaimg}\score.png')
r.close()


