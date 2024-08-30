## Splunk 資安報表自動檢核工具
### 簡介
在 Splunk 中抓取被防火牆阻擋的內網連外網連線次數最多的前 50 大 IP 相關資訊。

### 安裝環境
- Python 3.12.4
- Splunk Enterprise 9.2.2
- [Splunk SDK for Python](https://pypi.org/project/splunk-sdk/)
### 使用說明
1. 新增 .env 檔，填寫內容如下，與 main.py 放在同一目錄下
```python
# Splunk host (default: localhost)
host=[Splunk host]
# Splunk admin port (default: 8089)
port=8089
username=[Splunk username]
password=[Splunk password]
scheme=https
# Your version of Splunk (default: 6.2)
version=[Splunk Enterprise Version]
```
2. 執行 main.py 會產出 Excel 報表，如圖所示。(執行時間約一個多小時)
![report.png](/img/report.png)

### 程式流程
1. 在Splunk 上抓取「Checkpoint內對外阻擋TOP50: daily report」
2. 輸出原始「Checkpoint內對外阻擋TOP50: daily report」Excel 檔 
![origin_report.png](/img/origin_report.png)
3. 在 dest_ip 欄位的右邊增加一個 **fqdn 欄位（Fully Qualified Domain Name）**，顯示目標 IP 的域名
4. 透過 Splunk 查詢每個 dest_ip 的 DNS query
   1. 在公司的 DNS Server 查詢 Log <br>
      Splunk 語法: **search index={DNS Server index} {ip} | head 1 | table dns_answer_name** <br>
      ![splunk.png](/img/splunk.png)
   2. 若 Splunk 中無紀錄，使用 nslookup 反查域名
   3. 過濾掉私有 IP 及無意義的域名，這些域名不需顯示在報表。 <br>
      例如 : 包含目標 IP 的域名
      host-**219-68-146-45**.dynamic.kbtelecom.net <br>
	  guestnat-**104-133-122-101**.corp.google.com <br>
      rate-limited-proxy-**66-249-92-143**.google.com <br>
      cm-**84.215.131.230**.get.no <br>
      NK**219-91-5-144**.adsl.dynamic.apol.com.tw <br>