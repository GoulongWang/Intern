import splunklib.client as client
import splunklib.results as results
from dotenv import load_dotenv
import os
import xml.etree.ElementTree as ET
import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def connect_to_splunk(username, password, host='localhost', port='8089', owner='admin', app='search', sharing='user'):
    try:
        service = client.connect(host=host, port=port, username=username, password=password, owner=owner, app=app,
                                 sharing=sharing)
        if service:
            print("Splunk service created successfully")
            print("------------------------------------")
        return service
    except Exception as e:
        print(e)


def output_excel(jobResult):
    '''
    xmlData = str(jobResult)
    # 解析 XML
    root = ET.fromstring(xmlData)


    data = []
    for result in root.findall('result'):
        row = {}
        for field in result.findall('field'):
            # 欄位名稱
            key = field.get('k')

            # 欄位值
            value = field.find('value/text').text
            row[key] = value
            # 如果有多個 ip 只抓到第一個 ip

        # 將每列資料加入 data
        if row:
            data.append(row)

    df = pd.DataFrame(data)
    df.to_excel('output.xlsx', index=False)
    '''

    # 解析 XML 資料
    xmlData = str(jobResult)
    root = ET.fromstring(xmlData)

    data = []
    for result in root.findall('result'):
        src_ip = None
        dest_ips = []
        dest_ports = []
        counts = []

        for field in result.findall('field'):
            # 欄位名稱
            key = field.get('k')

            # 欄位值
            if key == 'src_ip':
                src_ip = field.find('value/text').text
            elif key == 'dest_ip':
                for value in field.findall('value'):
                    text_element = value.find('text')
                    if text_element is not None:
                        dest_ips.append(text_element.text)
            elif key == 'dest_port':
                for value in field.findall('value'):
                    text_element = value.find('text')
                    if text_element is not None:
                        dest_ports.append(text_element.text)
            elif key == 'count':
                for value in field.findall('value'):
                    text_element = value.find('text')
                    if text_element is not None:
                        counts.append(text_element.text)
        '''
        print(f"src_ip: {src_ip}")
        print(f"dest_ip: {dest_ips}")
        print(f"dest_port: {dest_ports}")
        print(f"counts: {counts}")
        print("\n")
        '''

        for dest_ip, dest_port, count in zip(dest_ips, dest_ports, counts):
            data.append([src_ip, dest_ip, dest_port, count])

    df = pd.DataFrame(data, columns=['src_ip', 'dest_ip', 'dest_port', 'count'])
    fileName = 'output.xlsx'
    df.to_excel(fileName, index=False, engine='openpyxl')

    # 合併 src_ip 儲存格
    wb = load_workbook(fileName)
    ws = wb.active

    # 從第 2 列開始處理，第 1 列是欄位名稱
    start_row = 2

    # 逐行遍歷 Dataframe，df.iterrows() 每次回傳一組包含列索引和列數據的資料。
    for index, row in df.iterrows():
        # 第一列為標題
        if index == 0:
            prev_src_ip = row['src_ip']
            continue

        # 合併儲存格
        # 如果目前此格的 src_ip 等於上一列的 src_ip 就合併
        if row['src_ip'] == prev_src_ip:
            # start_column 和 end_column 皆為 1，因為合併儲存格都在第一行
            # ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row + 1, end_column=1)
            start_row += 1
        # 相鄰兩列 src_ip 不同，就將 prev_src_ip 設為目前的 src_ip，並將 start_row 往下移動
        else:
            prev_src_ip = row['src_ip']
            start_row = index + 1
        # start_row += 1

    wb.save(fileName)
    print("Excel 文件已生成：output.xlsx")


def run_normal_mode_search(splunk_service, search_string, payload={}):
    try:
        job = splunk_service.jobs.create(search_string, **payload)
        while True:
            while not job.is_ready():
                pass
            if job["isDone"] == "1":
                break

        # print(job.results())

        '''
        # 印出 search id
        xmlData = str(job.results())
        # 解析 XML
        root = ET.fromstring(xmlData)

        # 找到 search_id 的 <field> 標籤
        search_id = None
        for field in root.findall('.//field[@k="search_id"]'):
            text_element = field.find('./value/text')
            if text_element is not None:
                search_id = text_element.text
                break

        print(f'Search ID: {search_id}')

        # 從 search_id 中擷取 sid
        # 從字串尾開始找出第一次出現 '_' 的 index，令為 A
        # 在從 Index 0 ~ Index A 的範圍內，從尾部開始找出 '_'，即可找出倒數第二個 '_'
        # 最後從倒數第二個 '_' 的下一個 index 開始印到字串尾就可以抓出 1723662300_20573
        last_underline_index = search_id.rfind('_')
        # penultimate 倒數第二
        penultimate_index = search_id.rfind('_', 0, last_underline_index)
        sid_underline = search_id[penultimate_index + 1:]
        sid_str = sid_underline.replace('_', '.')
        print(f'sid: {sid_str}')
        '''

        # 產出 Excel 檔案
        output_excel(job.results())

    except Exception as e:
        print(e)


def main():
    try:
        load_dotenv()
        Username = os.getenv("username")
        Password = os.getenv("password")
        Host = os.getenv("host")

        splunk_service = connect_to_splunk(username=Username, password=Password, host=Host)

        # Daily Report SID
        search_string = ("search index=_audit (info=granted OR info=completed) "
                         "savedsearch_name=\"Checkpoint內對外阻擋TOP50: daily report\" "
                         "search_id=\"*scheduler*\" "
                         "| head 1"
                         "| table search_id, savedsearch_name, _time")
        payload = {"exec_mode": "normal", "earliest_time": "-24h", "latest_time": "now"}
        # run_normal_mode_search(splunk_service, search_string, payload)

        # Daily Report 的語法
        search_string2 = (
            "search index=checkpoint action=Drop src_ip IN (10.*,172.16.*,172.18.*,172.19.*,172.17.*,192.168.*) dest!=10.*"
            "| eval dest_port=if(dest_port>1025,\"High Port(1025-65535)\",dest_port) "
            "| stats count by src_ip dest_port dest_ip "
            "| sort - count "
            "| stats list(dest_port) as dest_port list(dest_ip) as dest_ip list(count) as count by src_ip "
            "| sort - count "
            "| table src_ip dest_ip dest_port count "
            "| head 50")

        payload2 = {"exec_mode": "normal", "earliest_time": "-1d@d", "latest_time": "@d"}
        run_normal_mode_search(splunk_service, search_string2, payload2)
    except Exception as e:
        print(e)


if __name__ == "__main__":
    main()