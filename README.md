# excel_reporting_utils

## Overview

excel形式でのレポート作成用ライブラリです。
表形式のデータに対してつぎのようなレポートをexcel形式で作成可能です。

- セグメント別集計
- 期間別集計

---

## Requirement

- openxlsx 4.1.3
- stringr 1.4.0

---

## Reference

### `excel_reporting_manager$updateReports(wb, raw_data)`

説明

Workbookオブジェクトにレポートを追加します。

引数 

- *wb* : Workbookオブジェクト
- *raw_data* : 元データ

### `excel_reporting_manager$addSummaryReport(metric, calculated_values, segment_column)`

説明

サマリーレポートの定義を行います。

引数 

- *metric* : 合計対象の列
- *calculated_values* : 計算指標の
- *segment_column* : セグメントに利用する列

```r
source('./excel_reporting_utils/excel_reporting_manager.R', encoding="utf-8")

# 日付, 参照元, セッション, ユーザー数からなるアクセスデータ
access <- data.frame(
    date=c(20200101:20200130), 
    referer=rep(c('google', 'yahoo', '(direct)'), 10),
    session=rep(c(1:5, 6)),
    users=rep(c(1:3), 10)
)

erm <- excel_reporting_manager()

# アクセスデータの列名を設定する
erm$setRawDataColumnsMapping(
    list(
        date='date',# 日付(yyyymmdd)
        referer='referer',
        session='session',
        users='users'
    )
)

# サマリーレポートの定義
erm$addSummaryReport(
    metric=c('session', 'users'),
    calculated_values=list(
        'session_per_user'=function(row) {
            return(str_glue('=IFERROR(B{row} / C{row}, "-")'))
        }
    ),
    segment_column='referer'
)

erm$addRawData()

# Workbookオブジェクトを作成
wb <- openxlsx::createWorkbook()

erm$updateReports(wb, access)
openxlsx::saveWorkbook(wb, 'report.xlsx', overwrite = TRUE)
```

report.xlsx 

| referer	| session	| users | session_per_user | 
| :---: | :---: | :---: | :---: | 
| (direct)	| 45 | 	30	| 1.5 |
| google	| 25 | 	10	| 2.5 |
| yahoo	| 35 | 	20 | 1.75 |
| total	| 105 | 60	| 1.75 |


### `excel_reporting_manager$addDuringReport(metric, calculated_values, during_list, filter_column)`

説明

期間毎レポートを定義します。

引数

- *metric* : 合計対象の列
- *calculated_values* : 計算指標
- *during_list* : 期間
- *filter_column* : フィルターに使用する列

```r
source('./excel_reporting_utils/excel_reporting_manager.R', encoding="utf-8")

# 日付, 参照元, セッション, ユーザー数からなるアクセスデータ
access <- data.frame(
    date=c(20200101:20200130), 
    referer=rep(c('google', 'yahoo', '(direct)'), 10),
    session=rep(c(1:5, 6)),
    users=rep(c(1:3), 10)
)

erm <- excel_reporting_manager()

# アクセスデータの列名を設定する
erm$setRawDataColumnsMapping(
    list(
        date='date',# 日付(yyyymmdd)
        referer='referer',
        session='session',
        users='users'
    )
)

# 期間レポートの定義
erm$addDuringReport(
    metric=c('session'),
    calculated_values=list(
        'accumulation_session'=function(row) {
            return(str_glue('=IFERROR(SUM(B3:B{row}), "-")'))
        }
    ),
    during_list=list(
        during1=c('20200101', '20200107'),
        during2=c('20200108', '20200114'),
        during3=c('20200115', '20200121'),
        during4=c('20200122', '20200128'),
        during5=c('20200129', '20200131')
    ),
    filter_column='referer'
)

erm$addRawData()

# Workbookオブジェクトを作成
wb <- openxlsx::createWorkbook()

erm$updateReports(wb, access)
openxlsx::saveWorkbook(wb, 'report.xlsx', overwrite = TRUE)
```

report.xlsx

| 期間レポート |	total	| | 
| :---:|:---: |:---: |
| 期間 |	session	| accumulation_session |
| 20200101_20200107 | 22	| 22 |
| 20200108_20200114	| 23	| 45 |
| 20200115_20200121	| 24	| 69 |
| 20200122_20200128	| 25	| 94 |
| 20200129_20200131	| 11	| 105 |

---

## Install

つぎのコマンドをターミナル上で実行してください。

```bash
git clone https://github.com/gradus-AP/excel_reporting_utils.git
```

---

## License

google_analytics_reporting under [MIT license](https://en.wikipedia.org/wiki/MIT_License), see LICENSE.txt.