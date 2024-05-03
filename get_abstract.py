import time
import  selenium
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from urllib.parse import urljoin
import pandas as pd

#get直接返回，不再等待界面加载完成
desired_capabilities = DesiredCapabilities.CHROME
desired_capabilities["pageLoadStrategy"] = "none"

# 设置谷歌驱动器的环境
options = webdriver.ChromeOptions()
# 设置chrome不加载图片，提高速度
#options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
# 设置不显示窗口
#options.add_argument('--headless')
# 创建一个谷歌驱动器
driver = webdriver.Chrome(options=options)
# 设置搜索主题
theme = "可持续发展"

# 打开页面
driver.get("https://www.cnki.net")
# 传入关键字
WebDriverWait( driver, 100 ).until( EC.presence_of_element_located( (By.XPATH ,'''//*[@id="txt_SearchText"]''') ) ).send_keys(theme)
# 点击搜索
WebDriverWait( driver, 100 ).until( EC.presence_of_element_located( (By.CLASS_NAME ,"search-btn") ) ).click()
time.sleep(3)

# 点击切换中文文献
WebDriverWait( driver, 100 ).until( EC.presence_of_element_located( (By.CLASS_NAME ,"ch") ) ).click()
time.sleep(1)
#按照引用量排序
WebDriverWait( driver, 100 ).until( EC.element_to_be_clickable( (By.ID ,"CF") ) ).click()
time.sleep(1)
# 获取总文献数和页数
res_unm = WebDriverWait( driver, 100 ).until( EC.presence_of_element_located( (By.XPATH ,"//span[@class='pagerTitleCell']/em") ) ).text
# 去除千分位里的逗号
res_unm = int(res_unm.replace(",",''))
page_unm = int(res_unm/20) + 1

print(f"共找到 {res_unm} 条结果, {page_unm} 页。")

#从某一页开始爬取
page_start = 52
now_page = driver.find_element(By.CLASS_NAME, "countPageMark").text
while now_page != '79/300':
    time.sleep(1)
    print(now_page)
    try:
        WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "PageNext"))).click()
        time.sleep(1)
        now_page = driver.find_element(By.CLASS_NAME, "countPageMark").text
    except Exception as e:
        print(f"因为{e},点击不成功")

# 设置所需篇数
papers_need = res_unm

# 赋值序号, 控制爬取的文章数量
count = 1
# 当，爬取数量小于需求时，循环网页页码
rows = []
while count <= papers_need:
    # 等待加载完全，休眠3S
    time.sleep(3)
    title_list = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "fz14")))
    tbody = driver.find_element(By.TAG_NAME,'tbody')
    # 循环网页一页中的条目
    tr_elements = tbody.find_elements(By.TAG_NAME, 'tr')
    print(len(tr_elements))
    i=0
    for tr in tr_elements:
        try:
            seq_td = tr.find_element(By.CLASS_NAME, "seq")
            if seq_td:  # 确保我们找到了class='seq'的<td>
                # 获取seq_td元素的下一个同级<td>元素，即class='name'的<td>
                seq=int(seq_td.text.strip())
                title_td = seq_td.find_element(By.XPATH, "following-sibling::td[1]")
                title = title_td.text.strip()  # 移除可能的前后空格
                authors_td = title_td.find_element(By.XPATH, "following-sibling::td[1]")
                authors = authors_td.text.strip()  # 移除可能的前后空格
                source_td = authors_td.find_element(By.XPATH, "following-sibling::td[1]")
                source = source_td.text.strip()  # 移除可能的前后空格
                date_td = source_td.find_element(By.XPATH, "following-sibling::td[1]")
                date = date_td.text.strip()  # 移除可能的前后空格
                data_td = date_td.find_element(By.XPATH, "following-sibling::td[1]")
                data = data_td.text.strip()  # 移除可能的前后空格
                quote_td = data_td.find_element(By.XPATH, "following-sibling::td[1]")
                quote = quote_td.text.strip()  # 移除可能的前后空格
                downloadcnt = driver.find_element(By.CLASS_NAME, "downloadCnt").text
                print(data)
                if data =='期刊' or data =='博士' or data =='硕士':
                    # 点击条目
                    print(f'点击条目{seq}',end='\t')
                    title_list[i].click()
                    # 获取driver的句柄
                    print("获取driver的句柄",end='\t')
                    n = driver.window_handles
                    # driver切换至最新生产的页面
                    print("driver切换至最新生产的页面", end='\t')
                    driver.switch_to.window(n[-1])
                    # 开始获取页面信息
                    print('开始获取页面信息', end='\t')
                    # title = WebDriverWait( driver, 10 ).until( EC.presence_of_element_located((By.XPATH ,"/html/body/div[2]/div[1]/div[3]/div/div/div[3]/div/h1") ) ).text
                    # authors = WebDriverWait( driver, 10 ).until( EC.presence_of_element_located((By.XPATH ,"/html/body/div[2]/div[1]/div[3]/div/div/div[3]/div/h3[1]") ) ).text
                    try:
                        span_element= WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH ,"//span[contains(text(), '专题：')]")))
                        rowtit_p = span_element.find_element(By.XPATH, "following-sibling::p[1]")
                        rowtit = rowtit_p.text
                        print(rowtit)
                    except:
                        rowtit = '无'
                    abstract = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "abstract-text"))).text
                    try:
                        print('开始获取关键词', end='\t')
                        keywords = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CLASS_NAME, "keywords"))).text[:-1]
                    except:
                        keywords = '无'
                    url = driver.current_url
                    # 获取下载链接
                    # link = WebDriverWait( driver, 10 ).until( EC.presence_of_all_elements_located((By.CLASS_NAME  ,"btn-dlcaj") ) )[0].get_attribute('href')
                    # link = urljoin(driver.current_url, link)
                    # 写入文件
                    print("写入文件")
                    res = [count,title,authors,rowtit,date,source,data,quote,downloadcnt,keywords,abstract,url]
                    rows.append(res)
            i+=1
        except:
            print(f" 第{count} 条爬取失败\n")
            # 跳过本条，接着下一个
            i+=1
            continue
        finally:
            # 如果有多个窗口，关闭第二个窗口， 切换回主页
            n2 = driver.window_handles
            if len(n2) > 1:
                driver.close()
                driver.switch_to.window(n2[0])
            # 计数,判断需求是否足够
            count += 1
            if count == papers_need: break
    # 写入 Excel 文件
    try:
        existing_df = pd.read_excel(r'C:\Users\27612\Desktop\abstract.xlsx')
        # 追加新的数据到原有的DataFrame中
        # 注意：这里假设 rows 中的数据格式与 existing_df 中的列相匹配
        # 将 rows 转换为 DataFrame
        new_rows = pd.DataFrame(rows,columns=['序号', '标题', '作者', '专题', '日期', '来源', '数据','被引','下载量', '关键词', '摘要', 'URL'])
        existing_df = pd.concat([existing_df,new_rows], ignore_index=True)
        # 将更新后的DataFrame保存回Excel文件
        existing_df.to_excel(r'C:\Users\27612\Desktop\abstract.xlsx', index=False)
        print("数据已成功追加到Excel文件中。")
    except Exception as e:
        print(f"在追加数据到Excel文件时发生错误：{e}")
    rows.clear()
    # 切换到下一页
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@id='PageNext']"))).click()
# 关闭浏览器
driver.close()


