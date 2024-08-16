import splunklib.client as client
from dotenv import load_dotenv
import os
import xml.etree.ElementTree as ET
import pandas as pd


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
    xmlData = str(jobResult)
    # 解析 XML 資料
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

        for dest_ip, dest_port, count in zip(dest_ips, dest_ports, counts):
            data.append([src_ip, dest_ip, dest_port, count])

    df = pd.DataFrame(data, columns=['src_ip', 'dest_ip', 'dest_port', 'count'])
    # 去除 src_ip 重複儲存格
    df['src_ip'] = df['src_ip'].where(~df['src_ip'].duplicated())
    df.to_excel('output.xlsx', index=False)
    print("Excel 文件已生成：output.xlsx")


def run_normal_mode_search(splunk_service, search_string, payload={}):
    try:
        job = splunk_service.jobs.create(search_string, **payload)
        while True:
            while not job.is_ready():
                pass
            if job["isDone"] == "1":
                break

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

        # Daily Report 的語法
        search_string = (
            "search index=checkpoint action=Drop src_ip IN (10.*,172.16.*,172.18.*,172.19.*,172.17.*,192.168.*) dest!=10.*"
            "| eval dest_port=if(dest_port>1025,\"High Port(1025-65535)\",dest_port) "
            "| stats count by src_ip dest_port dest_ip "
            "| sort - count "
            "| stats list(dest_port) as dest_port list(dest_ip) as dest_ip list(count) as count by src_ip "
            "| sort - count "
            "| table src_ip dest_ip dest_port count "
            "| head 50")

        payload = {"exec_mode": "normal", "earliest_time": "-1d@d", "latest_time": "@d"}
        run_normal_mode_search(splunk_service, search_string, payload)
    except Exception as e:
        print(e)


if __name__ == "__main__":
    main()