# 臺灣電子發票自動下載工具

## 主要功能
- 自動化登入：自動填寫帳號、密碼及統一編號，登入財政部電子發票整合服務平台。

- reCAPTCHA 驗證：內建使用 OpenAI Whisper 進行處理。

- 批次下載：下載當月折讓單發票為 Excel 檔案，每月 7日 前同時下載上月折讓單發票。


## 環境要求

- Python 3.x

- Google Chrome 瀏覽器

- ChromeDriver：需與 Chrome 瀏覽器版本相符。

## 安裝與設定

1. 使用 pip 安裝專案所需的套件：<br/> ```pip install selenium aiohttp whisper```

2. 下載 ChromeDriver：<br/> 預設路徑為 ```chromedriver-win64\chromedriver.exe```

3. 設定 ```loginInfo.py``` 登入資訊：
   
```
ban = "統一編號"
user_id = "帳號"
password = "密碼"
User = "Windows使用者名稱"
```

## 使用方法

開啟終端並輸入 ```python InvoiceDownloader.py```

※ 程式會自動開啟 Chrome 瀏覽器，執行登入、導航及下載流程，並在完成後關閉瀏覽器。
