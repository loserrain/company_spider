# 104公司資料爬蟲

這是一個用於爬取 104 人力銀行公司資訊的爬蟲程式，可以將公司資訊匯出成多種格式（JSON、CSV、Markdown、Word）。

## 輸出範例(友嘉實業)

- 在output資料夾裡有友嘉實業股份有限公司的相關輸出資料


## 功能說明

- 支援多種輸出格式：
  - JSON：適合程式化處理
  - CSV：適合資料分析
  - Markdown：適合閱讀和網頁展示
  - Word：適合列印和分享
- 自動建立 output 資料夾存放所有輸出檔案
- 使用公司名稱作為檔名
- 完整的公司資訊擷取，包括：
  - 基本資訊（公司名稱、統一編號等）
  - 公司簡介
  - 主要商品/服務
  - 福利制度
  - 經營理念
  - 公司標籤
  - 最新消息
  - 公司發展歷程

## 安裝說明

1. 複製專案到本地：
```bash
git clone https://github.com/loserrain/comoany_spider.git
cd comoany_spider
```

2. 建立虛擬環境（建議）：
```bash
# Windows
python -m venv venv
.\venv\Scripts\activate

# MacOS/Linux
python -m venv venv
source venv/bin/activate
```

3. 安裝必要套件：
```bash
pip install -r requirements.txt
```

## 使用方法

1. 修改爬蟲目標：
   在 `company/spiders/company_spider.py` 中修改 `company_id`：
```python
def start_requests(self):
    company_id = 'e6o7g3l'  # 修改為目標公司的 ID
```

2. 執行爬蟲：
```bash
scrapy crawl company_spider
```

3. 檢查輸出結果：
   所有檔案都會儲存在 `output` 資料夾中，檔名格式為 `[公司名稱].[格式]`

## 檔案結構

```
104-company-crawler/
│  README.md
│  requirements.txt
│  scrapy.cfg
│
├─company/
│  │  __init__.py
│  │  items.py
│  │  middlewares.py
│  │  pipelines.py
│  │  settings.py
│  │
│  └─spiders/
│     │  __init__.py
│     │  company_spider.py
│
└─output/
    │  公司名稱.json
    │  公司名稱.csv
    │  公司名稱.md
    │  公司名稱.docx
```

## 如何取得公司 ID

1. 前往 104 人力銀行公司頁面
2. 從網址中取得公司 ID，例如：
   - 網址：`https://www.104.com.tw/company/e6o7g3l`
   - 公司 ID：`e6o7g3l`
