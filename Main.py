from requests_html import HTMLSession
from bs4 import BeautifulSoup
import requests
import os
import xlwt

def login(session,user,pwd):

    #   进入登陆界面
    res = session.get("http://jwxt.yxnu.edu.cn/default2.aspx")  
    #   cookie进行处理
    Cookie=str(res.cookies)[27:69]
    #   使用BeautifulSoup进行页面处理，提取出__VIEWSTATE
    soup = BeautifulSoup(res.text,'lxml')     
    viewState = soup.find('input', attrs={'name': '__VIEWSTATE'})['value'] 
    
    #调用验证码处理函数
    checkCode = CheckImag(session,Cookie)
    
    #处理登陆请求的data
    login_info = {
            "__VIEWSTATE": viewState,
            "txtUserName": user, 
            "TextBox2": pwd,
            "txtSecretCode": checkCode,
            "RadioButtonList1": "%D1%A7%C9%FA",#学生选项
            "Button1": "",
            "lbLanguage": ""
        }
    
    #处理登陆请求的请求头
    header = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate",
    "Cookie": Cookie,
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36",
    "Referer": "http://jwxt.yxnu.edu.cn/default2.aspx",
    "Host": "jwxt.yxnu.edu.cn",
    "Origin":"http://jwxt.yxnu.edu.cn",
    "Cache-Control": "max-age=0"
}
    
    #发送登陆请求
    requests.session().post(url='http://jwxt.yxnu.edu.cn/default2.aspx', data=login_info, headers=header)

    return header
    
def CheckImag(session,Cookie):
    #获取验证码的请求头，将Cookie放入其中，Cookie中的ASP.NET_SessionId是验证码核验的标准
    headeri = {
    "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Cookie": Cookie,
    "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36",
    "Referer": "http://jwxt.yxnu.edu.cn/default2.aspx",
    "Host": "jwxt.yxnu.edu.cn",
    "Cache-Control": "max-age=0"
    }
    
    #发送请求
    resi = session.get("http://jwxt.yxnu.edu.cn/CheckCode.aspx",headers=headeri,stream=True)

    #如果验证码文件已经存在则删除它，如果不存在就直接创建一个将验证码放入其中
    if os.path.exists(r'F://FzscoreGet//yanzheng.jpg'):
        os.remove(r'F://FzscoreGet//yanzheng.jpg')
    with open(r'F://FzscoreGet//yanzheng.jpg','wb')as f:
        f.write(resi.content)
        
    #打开验证码文件要求用户进行验证（待完善，目标自动识别）
    os.startfile(r'F://FzscoreGet//yanzheng.jpg')
    checkCode = input("请输入弹出的验证码:")

    return checkCode

def lncj(session,header,xh,name):
    #历年成绩获取
    data={
        'btn_zcj':'%C0%FA%C4%EA%B3%C9%BC%A8',#学年成绩：btn_xn 历年成绩：btn_zcj
        'ddlXN':'',
        'ddlXQ':'',
        '__EVENTVALIDATION': '',
        '__EVENTTARGET':'',   
        '__EVENTARGUMENT' :'',
        '__VIEWSTATE':'',
        'hidLanguage':'',
        'ddl_kcxz':'',
    }
    
    #使用get请求先get到__VIEWSTATE
    rln = session.get('http://jwxt.yxnu.edu.cn/xscjcx.aspx?xh={}&xm={}&gnmkdm=N121605'.format(xh,name),headers=header)
    #错误处理，如果报错则说明验证码或账户密码错误，不细化分析错误，学有余力可以自己搞下
    try:
        
        soup=BeautifulSoup(rln.text,'lxml')
        value3=soup.find('input', attrs={'name': '__VIEWSTATE'})['value']
        
        data['__VIEWSTATE']=value3
        
        #post请求获取到显示界面
        lncj = session.post('http://jwxt.yxnu.edu.cn/xscjcx.aspx?xh=2018906139&xm=%CC%C6%D5%BF&gnmkdm=N121605',data=data,headers=header)
    
        soup=BeautifulSoup(lncj.text,'lxml')
        return soup
    
    except TypeError:
        print("验证码或账户密码错误请检查")
        return 0


def excelWrite(soup):
    #EXCEL写入部分
    
    #如果已经存在则删除它
    if os.path.exists(r'F://FzscoreGet//Score.xls'):
       os.remove(r'F://FzscoreGet//Score.xls')
       
    #初始化工作环境使用Xlwt
    workbook = xlwt.Workbook(encoding = 'utf-8')
    #表名
    worksheet = workbook.add_sheet('Score')
    #行数，列数初始化
    row = 0
    column = 0
    for tr in soup.find_all('tr'):
        row = row+ 1
        column = 0
        for td in tr.find_all('td'):
             worksheet.write(row,column, label = str(td.get_text()))
             column=column+1
    
     #保存excel    
    workbook.save('Score.xls')


def main():
    #count是计数，UserOut是用户手动退出检测点
    count = 1
    UserOut = False
    

    
    while True:
        print("------第{}次登陆--------".format(count))
        
        session = HTMLSession()
        #    获取用户输入的用户名和密码
        user = input("请输入用户名:")
        pwd = input("请输入密码:")
        name = input("请输入您的名字:")
        
        header=login(session,user,pwd)
        soup = lncj(session,header,user,name)
        if soup:
            excelWrite(soup)
            break
        else:
            if count>=2:
                UserContinue = input("是否继续执行(Y/N):")
                if UserContinue=="N" or UserContinue=="n":
                    UserOut = True
                    break
                else:
                    count = count +1
                    continue
            else:
                count = count +1
                continue
    
    if UserOut:
        print("Bye!")
    else:
        print("操作完成")
        OpenExcel = input("是否需要打开文件(Y/N):")
        if OpenExcel=="Y" or OpenExcel=="y":
            os.startfile(r'F://FzscoreGet//Score.xls')
        else:
            print("Bye!")
        

    
main()

