import time
import pywinauto
import requests
import re
import subprocess
import win32gui
from bs4 import BeautifulSoup
from urllib.parse import unquote
from pywinauto import Application

def retry(times=5):
    def decorator(func):
        def wrapper(*args, **kwargs):
            last_exception = None
            for i in range(times):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    last_exception = e
                    print(f"Trying the {i+1} time")
            print(f"Failed after {times} attempts")
            raise last_exception
        return wrapper
    return decorator


# IDM下载函数
def download_with_idm(url, file_name, download_path):
    idm_path = "C:\\Program Files (x86)\\Internet Download Manager\\idman.exe"
    download_command = f"{idm_path} /d {url} /p {download_path} /f {file_name} /n /a /q"

    try:
        result = subprocess.run(download_command, check=True, capture_output=True, text=True)
        print(f"IDM输出：\n{result.stdout}")
        print(f"下载已开始：{url}\n文件名：{file_name}\n下载路径：{download_path}")
    except subprocess.CalledProcessError as error:
        print(f"下载失败，错误信息：{error}")
        print(f"IDM输出：\n{error.stdout}")


# 获取链接函数
@retry
def Get_Title(given_link):
    response = requests.get(given_link)
    soup = BeautifulSoup(response.text, "html.parser")
    title = soup.title.string
    print(title)

# 处理对应链接的函数
def Download_ASMR(given_link,download_path1,download_path2):
    response = requests.get(given_link)
    soup = BeautifulSoup(response.text, "html.parser")
    title = soup.title.string

    # 去除后缀
    text = title
    clean_text = re.sub(r'\[RJ\d+\]\s*', '', text)  # 删除[RJ******]和空格
    clean_text = re.sub(r'\s*-\s*Hentai\s+ASMR', '', clean_text)  # 删除 - Hentai ASMR并去除空格
    clean_text = re.sub(r'\s+', '', clean_text)  # 删除所有空格
    new_title = clean_text + '.mp3'  # 添加.mp3扩展名

    print(new_title)
    rj_number = re.search(r"\d+", given_link, re.IGNORECASE).group(0)

    if __name__ == "__main__":
        url1 = f"https://cdn.hentaiasmr.moe/asmr/RJ{rj_number}.mp3"
        url2 = f"https://cdn.hentaiasmr.moe/asmr4/RJ{rj_number}.mp3"
        file_name = new_title
        download_with_idm(url1, file_name, download_path1)
        download_with_idm(url2, file_name, download_path2)


# 提取tag
def extract_text_from_url(url):
    base_url = 'https://www.hentaiasmr.moe/tag/'
    if url.startswith(base_url):
        encoded_text = url[len(base_url):]
        # print(f"{encoded_text=}")
        decoded_text = unquote(encoded_text)
        return decoded_text
    else:
        return None


# 获取所有链接
def get_links(urls):
    response = requests.get(urls)
    html_content = response.content
    soup = BeautifulSoup(html_content, 'html.parser')
    all_link = soup.find_all('a')
    return all_link


# 选择页码
def get_selection(prompt):
    while True:
        number = input(prompt)
        try:
            number = int(number)
            if number >= 0:
                break
            else:
                print("无效数字，请重新输入")
        except ValueError:
            print("无效数字，请重新输入")
    return number


# 检测输入是否为tag
def get_tag(taglist):
    while True:
        tag = input("请输入想查看的tag")
        if tag in taglist:
            break
        else:
            print("无效tag，请重新输入")
    return tag


# 获取打开窗口句柄
def wait_for_window(title):
    while True:
        needed_hwnd = win32gui.FindWindow(None, title)
        if needed_hwnd != 0 and win32gui.GetWindowText(needed_hwnd) == title:
            return needed_hwnd
        time.sleep(0.1)


# 获取操作对象
def get_dlg(title=None, hwnd=None):
    if hwnd is None:
        hwnd = win32gui.FindWindow(None, title)
    app = pywinauto.Application().connect(handle=hwnd)
    return app.window(handle=hwnd)
def main():
    download_path1 = f"D:\\asmr\\"
    download_path2 = f"D:\\asmr4\\"
    tag_chart = []
    page_list = ["most-viewed", "popular", "latest", "random"]
    page_number = 0
    now_url = ""
    url_tags = f'https://www.hentaiasmr.moe/tags/'
    pattern_ASMR = re.compile(r'^https://www\.hentaiasmr\.moe/rj\d+\.html', re.I)
    pattern_tag = re.compile(r"tag")
    pattern_title = re.compile(r'\[[^\]]+\] .*')
    page_prompt = "请选择想要的页面.0:最多观看/1:近期热门/2:近期更新/3:随机/大于3:选择tag"

    # 选择页面
    selection = get_selection(page_prompt)
    if int(selection) < 4:
        now_url = f'https://www.hentaiasmr.moe/page/{page_number}/?filter={page_list[int(selection)]}'
    elif int(selection) >= 4:
        # 以下为选择tag
        links = get_links(url_tags)
        # print(f"{links=}")
        for link in links:
            link = link.get('href')
            if pattern_tag.search(str(link)):
                # print(str(link))
                # 以下是符合要求的链接：
                tag_chart.append(extract_text_from_url(link))
        tag_chart = list(set(tag_chart))
        tag_chart.remove(None)
        tag_chart = sorted(tag_chart)
        # print(tag_chart)
        for i in range(round(len(tag_chart)/10)):
            print(tag_chart[10*i:10*i +10])

        print("===================================")
        now_tag = get_tag(tag_chart)
        now_url = f"https://www.hentaiasmr.moe/tag/{now_tag}/page/{page_number}/"

    page_number = get_selection("请选择要查看的页码")
    links = get_links(now_url)
    # print(f"{links=}")
    result = []
    for link in links:
        href = link.get('href')
        text = link.get_text(strip=True)
        text_match = pattern_title.search(text)
        if text_match:
            text = text_match.group(0)
        result.append({'href': href, 'text': text})
        # print(f"{result}")
        if href and pattern_ASMR.search(href):
            # 以下是符合要求的链接：
            print("===================================")
            print(f"{href} {text}")

    # 判断是否继续
    while True:
        user_input = input("是否继续？（Y/N）")
        if user_input.upper() == "Y":
            # 用户选择继续，执行相应的代码
            print("继续执行...")
            break
        elif user_input.upper() == "N":
            # 用户选择不继续，退出程序
            print("退出程序...")
            exit()
        else:
            # 用户输入无效选项，提示重新输入
            print("无效选项，请重新输入。")

    # 继续进行
    for link in links:
        link = link.get('href')
        if pattern_ASMR.search(str(link)):
            # 以下是符合要求的链接，开始下载
            Download_ASMR(link,download_path1,download_path2)

    # IDM 确认
    Application(backend="uia").start(r"C:\Program Files (x86)\Internet Download Manager\IDMan.exe")
    hwnd = wait_for_window("Internet Download Manager 6.41")
    dlg = get_dlg(hwnd=hwnd)
    Toolbar = dlg.child_window(class_name="ToolbarWindow32")
    button = Toolbar.button('开始队列')
    time.sleep(0.3)
    button.click()


if __name__ == "__main__":
    main()