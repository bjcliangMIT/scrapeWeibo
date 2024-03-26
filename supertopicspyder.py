import time, xlrd, os, requests, json
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import excelSave as save

class Spider:
    def __init__(self,maxWeibo):
        self.book_name_xls = "/Users/bjcliang/Desktop/weibo/240325.xls" #填写你想存放excel的路径，没有文件会自动创建
        self.sheet_name_xls = '微博数据' #sheet表名
        self.pic_addr = '/Users/bjcliang/Desktop/weibo/img240325/' #设置自己想要放置图片的路径
        self.keyword = '龚俊'
        self.save_pic = True
        self.maxWeibo = maxWeibo; 
        self.num = 1 
        self.driver = webdriver.Chrome()#你的chromedriver的地址
        self.value_title = [["rid", "用户名称", "微博等级", "微博内容", "微博转发量","微博评论量",
                            "微博点赞","图片数量", "图片起始编号", "发布时间",
                            "话题名称","话题讨论数","话题阅读数"],]
        
        if os.path.exists(self.book_name_xls):
            print("文件已存在")
        else:
            print("文件不存在，重新创建")
            save.write_excel_xls(self.book_name_xls, self.sheet_name_xls, self.value_title)
        
    def set_driver1(self):
        self.driver.set_window_size(452, 790)
        self.driver.get('https://m.weibo.cn')
        self.driver.get("https://passport.weibo.cn/signin/login")
        print("开始自动登陆，若出现验证码手动验证")
    
    def set_driver2(self):        
        # click search box
        self.elem = self.driver.find_element("xpath","//*[@type='search']");
        self.elem.send_keys(self.keyword)
        self.elem.send_keys(Keys.ENTER)
        # wait and click
    
    def set_driver3(self):        
        yuedu_taolun = self.driver.find_element("xpath",
                "//*[@id='app']/div[1]/div[1]/div[1]/div[4]/div/div/div/a/div[2]/h4[1]").text
        self.yuedu = yuedu_taolun.split("　")[0]
        self.taolun = yuedu_taolun.split("　")[1]
        print(yuedu_taolun)
        

# 用来控制页面滚动
def Transfer_Clicks(browser):
    time.sleep(5)
    try:
        browser.execute_script("window.scrollBy(0,document.body.scrollHeight)", "")
    except:
        pass
    return "Transfer successfully \n"

#插入数据
def insert_data(elems,spider):
    for elem in elems:
        #save.write_excel_xls(spider.book_name_xls, spider.sheet_name_xls, spider.value_title);rid = 0
        workbook = xlrd.open_workbook(spider.book_name_xls)  # 打开工作簿
        sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
        worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
        rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数       
        rid = rows_old
        
        #用户名
        weibo_username = elem.find_elements(By.CSS_SELECTOR,'h3.m-text-cut')[0].text
        weibo_userlevel = "普通用户"
        #微博等级
        try: 
            weibo_userlevel_color_class = elem.find_elements(By.CSS_SELECTOR,
                                                             "i.m-icon")[0].get_attribute("class").replace("m-icon ","")
            if weibo_userlevel_color_class == "m-icon-yellowv":
                weibo_userlevel = "黄v"
            if weibo_userlevel_color_class == "m-icon-bluev":
                weibo_userlevel = "蓝v"
            if weibo_userlevel_color_class == "m-icon-goldv-static":
                weibo_userlevel = "金v"
            if weibo_userlevel_color_class == "m-icon-club":
                weibo_userlevel = "微博达人"     
        except:
            weibo_userlevel = "普通用户"
        #微博内容
        #点击“全文”，获取完整的微博文字内容
        weibo_content = get_all_text(elem, spider)
        #获取微博图片
        num_pics = 0
        if spider.save_pic:
            num_pics= get_pic(elem, spider)
        #获取分享数，评论数和点赞数               
        shares = elem.find_elements(By.CSS_SELECTOR,'i.m-font.m-font-forward + h4')[0].text
        if shares == '转发':
            shares = '0'
        comments = elem.find_elements(By.CSS_SELECTOR,'i.m-font.m-font-comment + h4')[0].text
        if comments == '评论':
            comments = '0'
        likes = elem.find_elements(By.CSS_SELECTOR,'i.m-icon.m-icon-like + h4')[0].text
        if likes == '赞':
            likes = '0'

        #发布时间
        weibo_time = elem.find_elements(By.CSS_SELECTOR,'span.time')[0].text
        '''
        print("用户名："+ weibo_username + "|"
              "微博等级："+ weibo_userlevel + "|"
              "微博内容："+ weibo_content + "|"
              "转发："+ shares + "|"
              "评论数："+ comments + "|"
              "点赞数："+ likes + "|"
              "图片数量："+ num_pics + "|"
              "图片起始编号："+ num-num_pic + "|"
              "发布时间："+ weibo_time + "|"
              "话题名称" + name + "|" 
              "话题讨论数" + yuedu + "|"
              "话题阅读数" + taolun)
        '''
        value1 = [[rid, weibo_username, weibo_userlevel,weibo_content,
                   shares,comments,likes, num_pics, spider.num-num_pics, weibo_time, 
                   spider.keyword, spider.yuedu, spider.taolun],]
        if rid % 50 == 0:
            print("当前插入第%d条数据" % rid)
        save.write_excel_xls_append_norepeat(spider.book_name_xls, value1)


#获取“全文”内容
def get_all_text(elem, spider):
    try:
        #判断是否有“全文内容”，若有则将内容存储在weibo_content中
        href = elem.find_element_by_link_text('全文').get_attribute('href')
        spider.driver.execute_script('window.open("{}")'.format(href))
        spider.driver.switch_to.window(driver.window_handles[1])
        weibo_content = spider.driver.find_element(By.CLASS_NAME, 'weibo-text').text
        spider.driver.close()
        spider.driver.switch_to.window(driver.window_handles[0])
    except:
        weibo_content = elem.find_elements(By.CSS_SELECTOR,'div.weibo-text')\
                        [0].text
    return weibo_content

def get_pic(elem, spider):
    try:
        #获取该条微博中的图片元素,之后遍历每个图片元素，获取图片链接并下载图片
        pic_links = elem.find_elements(By.CSS_SELECTOR, 
                                       'div > div > article > div > div:nth-child(2) > div > ul > li')
        num_pic = len(pic_links)
        for pic_link in pic_links:
            pic_link = pic_link.find_element(By.CSS_SELECTOR, 
                                                  'div > img').get_attribute('src')
            response = requests.get(pic_link)
            pic = response.content
            with open(spider.pic_addr + str(spider.num) + '.jpg', 'wb') as file:
                file.write(pic)
                spider.num += 1
        
    except Exception as e: print(e)
    return num_pic
    
    
#获取当前页面的数据
def get_current_weibo_data(spider):
    #开始爬取数据
        before = 0 
        after = 0
        n = 0 
        timeToSleep = 100
        while True:
            before = after
            Transfer_Clicks(spider.driver)
            time.sleep(3)
            elems = spider.driver.find_elements(By.CSS_SELECTOR,'div.card.m-panel.card9')
            print("当前包含微博最大数量：%d,n当前的值为：%d, n值到5说明已无法解析出新的微博" % (len(elems),n))
            after = len(elems)
            if after > before:
                n = 0
            if after == before:        
                n = n + 1
            if n == 5:
                print("当前关键词最大微博数为：%d" % after)
                insert_data(elems,spider)
                break
            if len(elems)>spider.maxWeibo:
                print("当前微博数以达到%d条"%spider.maxWeibo)
                insert_data(elems,spider)
                break

            if after > timeToSleep:
                print("抓取到%d多条，插入当前新抓取数据并休眠5秒" % timeToSleep)
                timeToSleep = timeToSleep + 100
                #insert_data(elems,spider) 
                time.sleep(5) 

