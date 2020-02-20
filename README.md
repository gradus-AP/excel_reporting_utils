# excel_reporting_utils

## Overview

excel形式でのレポート作成用ライブラリです。
表形式のデータに対してつぎのようなレポートをexcel形式で作成可能です。

- セグメント別集計
- 期間別集計

---

## Install

つぎのコマンドをターミナル上で実行してください。

```bash
git clone https://github.com/gradus-AP/excel_reporting_utils.git
```
---

## Requirement

- [openxlsx 4.1.3](https://www.rdocumentation.org/packages/openxlsx/versions/4.1.3)
- [stringr 1.4.0](https://github.com/tidyverse/stringr)

---

## Reference

#### `excel_reporting_manager$addSummaryReport`

説明

サマリーレポートの定義を行います。

引数 

- *name* : (既定値 summary)集計結果シート名を指定します.
- *raw_data_sheet* : (既定値 raw_data)集計対象シート名を指定します.
- *metric* : (必須)合計対象の列名を指定します.
- *calculated_values* : (省略可)計算指標を指定します.
- *segment_column* : (省略可)セグメント別に分ける識別子に用いる列名を指定します.
- *filter_column* : (省略可)集計行のフィルター列を指定します.
- *filter_value* : (既定値 *)集計行のフィルター値を指定します.

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

# nameは元データの名前(省略時raw_data)
erm$addRawData(
    name='access',
    mapping=list(
        date='date',# 日付(yyyymmdd)
        referer='referer',
        session='session',
        users='users'
    )
)

# サマリーレポートの定義
erm$addSummaryReport(
    raw_data_sheet='access',
    metric=c('session', 'users'),
    calculated_values=list(
        'session_per_user'=function(row) {
            return(stringr::str_glue('=IFERROR(B{row} / C{row}, "-")'))
        }
    ),
    segment_column='referer'
)

# Workbookオブジェクトを作成
wb <- openxlsx::createWorkbook()

# 元データをバインド 
erm$bindRawData(wb, list(access=access))

openxlsx::saveWorkbook(wb, 'report.xlsx', overwrite = TRUE)

```

report.xlsx 

| referer	| session	| users | session_per_user | 
| :---: | :---: | :---: | :---: | 
| (direct)	| 45 | 	30	| 1.5 |
| google	| 25 | 	10	| 2.5 |
| yahoo	| 35 | 	20 | 1.75 |
| total	| 105 | 60	| 1.75 |

#### `excel_reporting_manager$addDuringReport`

説明

期間毎レポートを定義します.

引数

- *raw_data_sheet* : (既定値 raw_data)集計するシート名を指定します.
- *metric* : (必須)合計対象の列名を指定します.
- *calculated_values* : (省略可)計算指標を指定します.
- *during_list* : (必須) 集計期間を指定します.
- *filter_column* : (省略可)フィルターに使用する列名を指定します.

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

# nameは元データの名前(省略時raw_data)
erm$addRawData(
    name='access',
    mapping=list(
        date='date',# 日付(yyyymmdd)
        referer='referer',
        session='session',
        users='users'
    )
)

# 期間レポートの定義
erm$addDuringReport(
    raw_data_sheet='access',
    metric=c('session'),
    calculated_values=list(
        'accumulation_session'=function(row) {
            return(stringr::str_glue('=IFERROR(SUM(B3:B{row}), "-")'))
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

# Workbookオブジェクトを作成
wb <- openxlsx::createWorkbook()

# 元データをバインド 
erm$bindRawData(wb, list(access=access))

# report.xlsxが出力される
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

#### `excel_reporting_manager$addPivotTable`

ピボット集計を定義します.

引数

- *raw_data_sheet* : (既定値 raw_data)集計対象シート名を指定します.
- *metric* : (必須)合計対象の列名を指定します.
- *calculated_values* : (省略可)計算指標を指定します.
- *segment_column* : (必須)集計行の列名を指定します.
- *filter_column* : (必須)集計列の列名を指定します.
- *filter_values* : (必須)列方向の識別子を指定します.


#### `excel_reporting_manager$addRawData`

元データを定義します.

引数

- *name* : (既定値 raw_data)元データの名前を指定します.
- *mapping* : (必須)元データの列名とその識別子の対応付けを行います.

#### `excel_reporting_manager$bindRawData`

説明

Workbookオブジェクトに元データをバインドし、レポートを追加します.

引数 

- *wb* : (必須)Workbookオブジェクト
- *rawData* : (必須)元データ

---

## License

excel_reporting_utils under [MIT license](https://en.wikipedia.org/wiki/MIT_License), see [LICENSE.txt](./LICENSE.txt).