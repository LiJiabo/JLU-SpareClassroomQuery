from io import StringIO
import requests
import base64
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from PIL import Image
import time
import pandas as pd

#读取配置文件
conf_path='[uims.conf的路径(Windows路径记得加转义字符!)]'
conf=pd.read_csv(conf_path,index_col=0)
conf_dict=conf.set_index('property')['value'].to_dict()

#识别验证码
options=Options()
if(conf_dict['headless']=='True'):
    options.add_argument('--headless')
browser = webdriver.Firefox(options=options)
url='https://uims.jlu.edu.cn/ntms/'
browser.get(url)
captcha_img=browser.find_element_by_id('captcha_img')
captcha_img.screenshot(conf_dict['captcha path']+'captcha.png')
img=Image.open(conf_dict['captcha path']+'captcha.png')
img=img.convert('L')
threshold=int(conf_dict['threshold'])# 设定阈值
points=[]
for i in range(256):
    if i<threshold:
        points.append(0)
    else:
        points.append(1)
img=img.point(points,'1')
img.save(conf_dict['captcha path']+'captcha1.png')
    # 百度OCR识图
request_url = conf_dict['request url']
        # 二进制方式打开图片文件
f = open(conf_dict['captcha path']+'captcha1.png', 'rb')
img = base64.b64encode(f.read())
params = {"image":img}
        # 获取access_token
access_token_url='https://aip.baidubce.com/oauth/2.0/token'
client_id=conf_dict['API Key']
client_secret=conf_dict['Secret Key']
access_token_params={"grant_type":"client_credentials","client_id":client_id,"client_secret":client_secret}
response = requests.post(access_token_url,data=access_token_params)
if response:
    access_token_json=response.json()
    access_token=access_token_json["access_token"]

request_url = request_url + "?access_token=" + access_token
headers = {'content-type': 'application/x-www-form-urlencoded'}
response = requests.post(request_url, data=params, headers=headers)
if response:
    verifyCode=response.json()['words_result'][0]['words']

#继续操作
browser.find_element_by_id('txtUserName').send_keys(conf_dict['username'])
browser.find_element_by_id('txtPassword').send_keys(conf_dict['password'])
browser.find_element_by_id('VerifyCode').send_keys(verifyCode)
browser.find_element_by_id('VerifyCode').send_keys(Keys.ENTER)
time.sleep(5)
browser.find_element_by_id('dijit_layout_AccordionPane_1_button_title').click() # 公众功能
time.sleep(0.5)
browser.find_element_by_id('Ntms_Menu_6').click() # 教室查询及预约
time.sleep(0.5)
browser.find_element_by_id('Ntms_Menu_10').click() # 空闲教室查询及预约
time.sleep(3)
browser.find_element_by_id('dijit_form_DropDownButton_0_label').click() # 选择教学楼
time.sleep(0.5)
browser.find_element_by_id('dijit_form_TextBox_0').send_keys(conf_dict['building keyword']) # 输入教学楼名称
browser.find_element_by_id('dijit_form_Button_2_label').click() # 查询
time.sleep(0.5)
browser.find_element_by_id('dojox_grid__View_1').click() # 选择经信教学楼
time.sleep(0.5)
browser.find_element_by_id('dijit_form_Button_3_label').click() # 确定
time.sleep(0.5)
if conf_dict['date']=='Today':
    browser.find_element_by_id('dijit_form_DateTextBox_0').send_keys(time.strftime("%Y-%m-%d"))
else:
    browser.find_element_by_id('dijit_form_DateTextBox_0').send_keys(conf_dict['date'])
time.sleep(0.5)
browser.find_element_by_id('ntms_widget_ClassSetInput_0').send_keys(conf_dict['class num']) # 课节
browser.find_element_by_id('dijit_form_NumberTextBox_0').send_keys(conf_dict['min capacity']) # 最小人数
browser.find_element_by_id('ntms_teachRes_classroomTakeup_0_search_label').click() # 查询

#调整格式并写入文件
result=browser.find_element_by_id('dojox_grid_DataGrid_1').text
result=result.replace('教室编号\n类型\n容量(人)\n用途\n注释','教室编号 类型 容量(人) 用途 注释')
df=pd.read_csv(StringIO(result),sep=' ')
if(conf_dict['date']=='Today'):
    df.to_csv(conf_dict['csv base']+time.strftime('%Y-%m-%d')+"-空教室.csv")
else:
    df.to_csv(conf_dict['csv base']+conf_dict['date']+"-空教室.csv")