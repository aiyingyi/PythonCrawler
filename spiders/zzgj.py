from utils import getHTMLText, str_format, item_filter, write_excel_xlsx
from bs4 import BeautifulSoup
import re
import time

from utils.email_utils import send_email

"""
    http://www.365trade.com.cn/  中招国际
"""


# 爬取数据并返回结果
def parse(page):
    # 定义list，存储爬取的结果
    result = []
    soup = BeautifulSoup(page, "html.parser")

    # 获取table标签
    for table in soup.find_all("table", attrs={"class": "table_text"}):
        # 获取父tr节点
        for tr in table.find_all("tr"):
            # 存储一条爬取的数据
            td_list = tr.find_all("td")
            if len(td_list) == 6:
                # 获取td中子标签的文本
                title = str_format(td_list[0].find("a").text)
                # 匹配网址链接
                hrefs = re.findall('https?.*\.html', td_list[0].find("a").attrs["href"])
                update_time = str_format(td_list[4].text)
                # 将元组数据加入列表
                if len(hrefs) >= 0:
                    result.append((title, hrefs[0], update_time))
    return result


def get_sntba_info():
    """
    获取爬取数据并过滤
    :return: 符合条件的数据列表
    """
    words = ["信息", "系统", "数字化", "数据", "软件"]
    url = "http://bulletin.sntba.com/xxfbcmses/search/bulletin.html?searchDate=1996-04-20&dates=300&word=&categoryId=88&industryName=&area=&status=01&publishMedia=&sourceInfo=&showStatus=,lt&page="

    page_index = 1
    res_set = set()
    set_size = 0
    while True:
        # 分页查询的url
        request_url = url + str(page_index)

        print(request_url)

        page_index = page_index + 1
        page = getHTMLText(request_url)
        items = parse(page)
        set_size = len(res_set)

        res_set = res_set | set(items)
        # 假定总数不会再发生变化
        if set_size == len(res_set):
            break

    return item_filter(list(res_set), words)


if __name__ == "__main__":
    res = get_sntba_info()

    prefix = time.strftime("%Y-%m-%d", time.localtime()) + "_"

    file_path = '招标项目.xlsx'

    sheet_name = '招标项目'

    # 设置excel表头
    header = ("项目名称", "查看链接", "更新日期")
    res.insert(0, header)

    # 添加文件路径
    file_path = prefix + file_path
    write_excel_xlsx(file_path, sheet_name, res)

    # 发送邮件
    receiver = ['1309961163@qq.com', '1394783493@qq.com']
    send_email(file_path, receiver)
