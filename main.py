# This is a sample Python script.

# 原始思路和代码于CSDN博主「Parzival_」的原创文章中提出，链接：https://blog.csdn.net/Parzival_/article/details/122360528
# 本代码由lijinhai0804进行优化，主要是升级使用为selenium 4版本的包，且适配了更通用的Chrome浏览器，版本118.0.5993.70（正式版本）（64 位）。
# 交流讨论：lijinhai0804@whu.edu.cn;https://github.com/lijinhai0804
# 使用前需要安装selenium 4和 webdriver-manager, openpyxl, xlrd, pandas. 在pycharm中可以直接安装最新版，非常方便！
# 转载请附上原文出处链接及上述声明。

from selenium import webdriver
import os
from selenium.webdriver.common.by import By
import time

# 强烈建议用校园网的IP登录，非常方便不用进行login。login函数默认注释掉，有需要的人可以启用。
# def login(driver):
#     '''登录wos'''
#     # 通过CHINA CERNET Federation登录
#     driver.find_element(By.CSS_SELECTOR, '.mat-select-arrow').click()
#     driver.find_element(By.CSS_SELECTOR, '#mat-option-9 span:nth-child(1)').click()
#     driver.find_element(By.CSS_SELECTOR,
#                         'button.wui-btn--login:nth-child(4) span:nth-child(1) span:nth-child(1)').click()
#     time.sleep(3)
#     login = driver.find_element(By.CSS_SELECTOR, '#show')
#     login.send_keys('武汉大学')  # 改成你的学校名
#     time.sleep(0.5)
#     driver.find_element(By.CSS_SELECTOR, '.dropdown-item strong:nth-child(1)').click()
#     driver.find_element(By.CSS_SELECTOR, '#idpSkipButton').click()
#     time.sleep(1)
#     # ! 跳转到学校的统一身份验证(想自动输入账号密码就把下面两行注释解除,按照自己学校的网址修改一下css选择器路径)
#     # driver.find_element(By.CSS_SELECTOR, 'input#un').send_keys('你的学号') # 改成你的学号/账号
#     # driver.find_element(By.CSS_SELECTOR, 'input#pd').send_keys('你的密码') # 改成你的密码
#     time.sleep(20)  # ! 手动输入账号、密码、验证码，点登录


def send_key(driver, path, value):
    '''driver -> driver;\n
       path -> css选择器;\n
       value -> 填入值
    '''
    markto = driver.find_element(By.CSS_SELECTOR, path)
    markto.clear()
    markto.send_keys(value)


def rename_file(SAVE_TO_DIRECTORY, name, record_format='excel'):
    '''导出文件重命名 \n
       SAVE_TO_DIRECTORY -> 导出记录存储位置(文件夹)；\n
       name -> 重命名为
    '''
    # files = list(filter(lambda x:'savedrecs' in x and len(x.split('.'))==2,os.listdir(SAVE_TO_DIRECTORY)))
    while True:
        files = list(filter(lambda x: 'savedrecs' in x and len(x.split('.')) == 2, os.listdir(SAVE_TO_DIRECTORY)))
        if len(files) > 0:
            break

    files = [os.path.join(SAVE_TO_DIRECTORY, f) for f in files]  # add path to each file
    files.sort(key=lambda x: os.path.getctime(x))
    newest_file = files[-1]
    # newest_file=os.path.join(SAVE_TO_DIRECTORY,'savedrecs.txt')
    if record_format == 'excel':
        os.rename(newest_file, os.path.join(SAVE_TO_DIRECTORY, name + ".xls"))
    elif record_format == 'bib':
        os.rename(newest_file, os.path.join(SAVE_TO_DIRECTORY, name + ".bib"))
    else:
        os.rename(newest_file, os.path.join(SAVE_TO_DIRECTORY, name + ".txt"))


# 判断页面元素是否存在。element处传入XPATH路径。
def isElementExist(driver, element):
    try:
        driver.find_element(By.XPATH, element)
        return True
    except:
        return False


# 以下是主要调用的核心函数
def startdownload(url, record_num, SAVE_TO_DIRECTORY, record_format='excel', reverse=False):
    '''url -> 检索结果网址; \n
       record_num -> 需要导出的记录条数(检索结果数); \n
       SAVE_TO_DIRECTORY -> 记录导出存储路径(文件夹)，在代码末尾设置;\n
       reverse -> 是否设置检索结果降序排列, default=False \n
       ----------------------------------------------------
       tip1:首次打开wos必须登录,在学校统一身份认证处需要手动输入验证码并点击登录，需要用login函数;IP登录请忽视，强烈建议校园网IP登录！！！
       tip2:第一次导出时，需要手动在10秒内（下文可修改）设置好定制的内容，后续都会直接点击定制好的导出字段
       tip3:建议以下都用完整XPATH进行元素寻找，绝对不会找不到！！！
    '''

    # 创建一个Chrome配置文件对象
    ch = webdriver.ChromeOptions()
    # 设置下载目录为SAVE_TO_DIRECTORY变量的值，这应该是一个有效的本地路径,最新版selenium4无executable_path
    ch.add_experimental_option("prefs", {
        "download.default_directory": SAVE_TO_DIRECTORY,
        "download.prompt_for_download": False, })
    # 创建一个Chrome驱动对象，使用chromedriver的可执行路径和自定义的配置文件
    driver = webdriver.Chrome(options=ch)
    # 导航到url变量指定的网址，这应该是一个有效的网址
    driver.get(url)
    driver.maximize_window()  # 窗口最大化
    # 暂停脚本执行4秒，等待网页加载或渲染
    time.sleep(4)
    # 调用函数，判断是否有弹窗存在，如果有，就关闭它
    isElementExist(driver,
                   '/ html / body / div[3] / div[2] / div / div[1] / div / div[2] / div / button[2]')  # 调用函数，判断“接受Cookies”是否出现
    if isElementExist:
        print("Ture")
        driver.find_element(By.XPATH,
                            '/ html / body / div[3] / div[2] / div / div[1] / div / div[2] / div / button[2]').click()  # 按掉接受Cookies按钮
    time.sleep(6)  # 这次等待是为了让右下角的弹窗加载出来
    isElementExist(driver, '/ html / body / div[5] / div / div[1] / button')  # 调用函数，判断“右下角弹窗”是否出现
    if isElementExist:
        print("Ture")
        driver.find_element(By.XPATH, '/ html / body / div[5] / div / div[1] / button').click()  # 关闭接受COOKIES后右下角的弹窗
    time.sleep(0.5)

    # 获取需要导出的文献数量
    # record_num = int(driver.find_element(By.CSS_SELECTOR, '.brand-blue').text)
    # 按时间降序排列，本功能默认关闭
    if reverse:
        driver.find_element(By.CSS_SELECTOR,
                            '.top-toolbar wos-select:nth-child(1) button:nth-child(1) span:nth-child(2)').click()
        driver.find_element(By.CSS_SELECTOR, "div.wrap-mode:nth-child(2) span:nth-child(1)").click()
        time.sleep(3)

    # 开始导出
    start = 1  # 起始记录
    i = 0  # 导出记录的数字框id随导出次数递增
    flag = 0  # mac文件夹默认有一个'.DS_Store'文件,win系统改成0，原本是1
    while start < record_num:
        print(start, '-', start + 999, ' ', 'start')
        driver.find_element(By.XPATH,
                            '/html/body/app-wos/main/div/div/div[2]/div/div/div['
                            '2]/app-input-route/app-base-summary-component/div/div[2]/app-page-controls['
                            '1]/div/app-export-option/div/app-export-menu/div/button').click()  # 点击导出按钮
        # driver.find_element(By.CSS_SELECTOR, 'button.cdx-but-md:nth-child(2) span:nth-child(1)').click()  # 点击导出按钮
        time.sleep(1)
        if record_format == 'excel':
            time.sleep(0.5)
            driver.find_element(By.CSS_SELECTOR, '#exportToExcelButton').click()  # 选择导出格式为excel
            time.sleep(0.5)
            RecordOptions = driver.find_element(By.XPATH,
                                                '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                                '1]/app-export-overlay/div/div[3]/div['
                                                '2]/app-export-out-details/div/div['
                                                '2]/form/div/fieldset/mat-radio-group/div[3]/mat-radio-button/label')
            driver.execute_script("arguments[0].click();", RecordOptions)  # 不需要加载出来就能点击,选择自定义记录条数
            time.sleep(1)
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                          '1]/app-export-overlay/div/div[3]/div[2]/app-export-out-details/div/div['
                                          '2]/form/div/fieldset/mat-radio-group/div[3]/mat-form-field[1]/div/div['
                                          '1]/div[3]/input').clear()  # 清除自定义记录条数起点
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                          '1]/app-export-overlay/div/div[3]/div[2]/app-export-out-details/div/div['
                                          '2]/form/div/fieldset/mat-radio-group/div[3]/mat-form-field[2]/div/div['
                                          '1]/div[3]/input').clear()  # 清除自定义记录条数终点
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                          '1]/app-export-overlay/div/div[3]/div[2]/app-export-out-details/div/div['
                                          '2]/form/div/fieldset/mat-radio-group/div[3]/mat-form-field[1]/div/div['
                                          '1]/div[3]/input').send_keys(start)  # 输入自定义记录条数起点
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                          '1]/app-export-overlay/div/div[3]/div[2]/app-export-out-details/div/div['
                                          '2]/form/div/fieldset/mat-radio-group/div[3]/mat-form-field[2]/div/div['
                                          '1]/div[3]/input').send_keys(start + 999)  # 输入自定义记录条数终点
            if i == 0:
                time.sleep(10)  # 在这10秒内，手动edit合适的字段，然后静候花开.后面程序会自动选择定制好的导出选项（前提是不重启浏览器或重新进入界面）
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div['
                                          '2]/app-input-route[1]/app-export-overlay/div/div[3]/div['
                                          '2]/app-export-out-details/div/div[2]/form/div/div['
                                          '1]/wos-select/button').click()  # 更改导出字段
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                          '1]/app-export-overlay/div/div[3]/div[2]/app-export-out-details/div/div['
                                          '2]/form/div/div[1]/wos-select/div/div/div/div[4]/span').click()  # 选择修改好的自定义字段
            # 'div.wrap-mode:nth-child(3) span:nth-child(1)').click()  # 选择所需字段(excel:3完整/4自定义; txt:3完整/4完整+引文)
            # time.sleep(0.5)
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                          '1]/app-export-overlay/div/div[3]/div[2]/app-export-out-details/div/div['
                                          '2]/form/div/div[2]/button[1]').click()  # 点击导出按钮
            # driver.find_element(By.CSS_SELECTOR, 'div.flex-align:nth-child(3) button:nth-child(1)').click()  # 点击导出
            time.sleep(1)
            # 下面的意思是检测有没有成功生成文件，成功生成就继续
            while len(os.listdir(SAVE_TO_DIRECTORY)) == flag:
                time.sleep(10)  # 等待下载完毕
            # 导出文件按照包含的记录编号重命名
            rename_file(SAVE_TO_DIRECTORY, 'record-' + str(start) + '-' + str(start + 999), record_format=record_format)
            start = start + 1000
            print(start)
        # elif内是导出bib，如果需要别的格式就修改路径，注意bib默认的是导出所有记录+参考文献，一次只能500篇
        elif record_format == 'bib':
            time.sleep(0.5)
            driver.find_element(By.XPATH, '/html/body/div[5]/div[2]/div/div/div/div/div[7]/button').click()  # 选择导出格式为bib
            time.sleep(0.5)
            RecordOptions = driver.find_element(By.XPATH,
                                                '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                                '1]/app-export-overlay/div/div[3]/div['
                                                '2]/app-export-out-details/div/div['
                                                '2]/form/div/fieldset/mat-radio-group/div[3]/mat-radio-button/label')
            driver.execute_script("arguments[0].click();", RecordOptions)  # 不需要加载出来就能点击,选择自定义记录条数
            time.sleep(1)
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                          '1]/app-export-overlay/div/div[3]/div[2]/app-export-out-details/div/div['
                                          '2]/form/div/fieldset/mat-radio-group/div[3]/mat-form-field[1]/div/div['
                                          '1]/div[3]/input').clear()  # 清除自定义记录条数起点
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                          '1]/app-export-overlay/div/div[3]/div[2]/app-export-out-details/div/div['
                                          '2]/form/div/fieldset/mat-radio-group/div[3]/mat-form-field[2]/div/div['
                                          '1]/div[3]/input').clear()  # 清除自定义记录条数终点
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                          '1]/app-export-overlay/div/div[3]/div[2]/app-export-out-details/div/div['
                                          '2]/form/div/fieldset/mat-radio-group/div[3]/mat-form-field[1]/div/div['
                                          '1]/div[3]/input').send_keys(start)  # 输入自定义记录条数起点
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                          '1]/app-export-overlay/div/div[3]/div[2]/app-export-out-details/div/div['
                                          '2]/form/div/fieldset/mat-radio-group/div[3]/mat-form-field[2]/div/div['
                                          '1]/div[3]/input').send_keys(start + 499)  # 输入自定义记录条数终点,bib全记录和参考文献一次只能500篇
            if i == 0:
                time.sleep(10)  # 在这10秒内，手动edit合适的字段，然后静候花开.后面程序会自动选择定制好的导出选项（前提是不重启浏览器或重新进入界面）
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div['
                                          '2]/app-input-route[1]/app-export-overlay/div/div[3]/div['
                                          '2]/app-export-out-details/div/div[2]/form/div/div['
                                          '1]/wos-select/button').click()  # 更改导出字段
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                          '1]/app-export-overlay/div/div[3]/div[2]/app-export-out-details/div/div['
                                          '2]/form/div/div[1]/wos-select/div/div/div/div[4]/span').click()  # 选择修改好的自定义字段
            driver.find_element(By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route['
                                          '1]/app-export-overlay/div/div[3]/div[2]/app-export-out-details/div/div['
                                          '2]/form/div/div[2]/button[1]').click()  # 点击导出按钮
            time.sleep(1)
            # 下面的意思是检测有没有成功生成文件，成功生成就继续
            while len(os.listdir(SAVE_TO_DIRECTORY)) == flag:
                time.sleep(10)  # 等待下载完毕
            # 导出文件按照包含的记录编号重命名
            rename_file(SAVE_TO_DIRECTORY, 'record-' + str(start) + '-' + str(start + 499), record_format=record_format)
            start = start + 500
            print(start)
        i = i + 2
        flag = flag + 1

    time.sleep(10)
    driver.quit()


# 主要参数修改在此
# 所参考的原创代码出自于CSDN博主「Parzival_」的原创文章,转载请附上原文出处链接及本声明。
# 原创文章链接：https://blog.csdn.net/Parzival_/article/details/122360528
# 本代码由@lijinhai0804(https://github.com/lijinhai0804)进行优化，主要是升级使用为selenium 4版本的包，且适配了更通用的Chrome浏览器。

if __name__ == '__main__':
    # WOS“检索结果”页面的网址
    url = 'https://webofscience.clarivate.cn/wos/woscc/summary/832beac5-03b8-401d-8076-2d91f76fdfb3-ac500860/relevance/1'
    # 导出到本地的存储路径(自行修改)
    download_path = r'D:\test'  # 最前面加r，就可以避免保留字符冲突问题，注意斜杠是向右下
    startdownload(url, 1500, download_path, record_format='bib', reverse=False)  # 主要函数的参数在这设定
    print('Done')
