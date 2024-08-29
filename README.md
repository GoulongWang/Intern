## Splunk 資安報表自動檢核工具
### 簡介
在 Splunk 中抓取被防火牆阻擋的內網連外網連線次數最多的前 50 大 IP 相關資訊。

### 安裝環境
- Python 3.12.4
- Splunk Enterprise 9.2.2
- [Splunk SDK for Python](https://pypi.org/project/splunk-sdk/)
### 使用說明
1. 新增 .env 檔，填寫內容如下，與 main.py 放在同一目錄下
```
host=[Splunk IP]
port=8089
username=[Splunk username]
password=[Splunk password]
scheme=https
version=[Splunk Enterprise Version]
```
2. 執行 main.py 會產出 Excel 報表，如圖所示。
![report.png](/report.png)