# 実行環境・構成ガイド

## 1. 実行環境

- Python バージョン: 3.13.3  
- 使用パッケージ:  
  `pip install pywin32`（`win32com` を使用するため）

## 2. 実行手順

1. ターミナルを開く  
2. 以下のコマンドを実行 : python main.py

## 3. ファイル構成

- `main.py`：エンドポイント（起動用スクリプト）  
- `DailyReport.py`：Excelマクロファイル `VBA02_20220528_SalesReportDaily_ver5.0.xlsm` に対応  
- `ReportGen_SendMail.py`：Excelマクロファイル `VBA05_20190709_SalesReport_SendMail_ver2.0.xlsm` に対応  

## 4. フォルダ構成

- main.py  
- DailyReport.py  
- ReportGen_SendMail.py  
- VBA02_20220528_SalesReportDaily_ver5.0.xlsm  
- VBA05_20190709_SalesReport_SendMail_ver2.0.xlsm  

## 5. 重要なポイント

- メール送信機能を使用するには、Outlook へのサインインが必要です。
