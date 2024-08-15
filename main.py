import splunklib.client as client
import splunklib.results as results
from dotenv import load_dotenv
import os
import xml.etree.ElementTree as ET


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


def run_normal_mode_search(splunk_service, search_string, payload={}):
    try:
        job = splunk_service.jobs.create(search_string, **payload)
        while True:
            while not job.is_ready():
                pass
            if job["isDone"] == "1":
                break

        # print(job.content)
        # print(job.results())
        # print(type(job.results()))

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

    except Exception as e:
        print(e)


def main():
    try:
        load_dotenv()
        Username = os.getenv("username")
        Password = os.getenv("password")
        Host = os.getenv("host")

        splunk_service = connect_to_splunk(username=Username, password=Password, host=Host)
        search_string = ("search index=_audit (info=granted OR info=completed) "
                         "savedsearch_name=\"Checkpoint內對外阻擋TOP50: daily report\" "
                         "search_id=\"*scheduler*\" "
                         "| head 1"
                         "| table search_id, savedsearch_name, _time")
        payload = {"exec_mode": "normal", "earliest_time": "-24h", "latest_time": "now"}
        run_normal_mode_search(splunk_service, search_string, payload)
        print("結束")
    except Exception as e:
        print(e)


if __name__ == "__main__":
    main()