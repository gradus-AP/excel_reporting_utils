#' ---------------------------------------------------------------------------
#' @description 
#' excel レポート作成用R Script
#' 
#' @import 
#' - `openxlsx 4.1.3`
#' - `stringr 1.4.0`
#' ---------------------------------------------------------------------------
library(tidyverse)
config <- list(
    # 親ディレクトリ
    wd = 'PATH_TO_WORK_DIRECTORY',
    access=list(
        # xlsx形式の元データのパス
        RAW_DATA.PATH='./sample/blog_access.xlsx',
        # 出力先xlsxパス
        REPORT_PATH='./sample/report.xlsx'
    )
)

source('./excel_reporting_utils/excel_reporting_manager.R', encoding="utf-8")

config <- config[['access']]
# Workbookオブジェクトを作成
init <- function() {
    if (file.exists(config$REPORT_PATH)) {
        return(openxlsx::loadWorkbook(config$REPORT_PATH))
    } 
    wb <- openxlsx::createWorkbook()
    return(wb)
}

erm <- excel_reporting_manager()

# nameは元データの名前(省略時raw_data)
erm$addRawData(
    mapping=list(
        date='日付',# 日付(yyyymmdd)
        referer='参照元',
        session='セッション',
        users='ユーザー',
        page_views='ページビュー数'
    )
)

# サマリーレポートの定義
erm$addSummaryReport(
    metric=c('session', 'users', 'page_views'),
    calculated_values=list(
        'page_view_per_session'=function(row) {
            return(str_glue('=IFERROR(D{row} / B{row}, "-")'))
        },
        'page_view_per_user'=function(row) {
            return(str_glue('=IFERROR(D{row} / C{row}, "-")'))
        }
    ),
    segment_column='referer'
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

# Workbookオブジェクトをxlsx形式で保存
save <- function(wb) {
    # フォント設定
    openxlsx::modifyBaseFont(wb, fontSize = 10.5, fontColour = "black", fontName = "Meiryo")
    openxlsx::saveWorkbook(wb, file=config$REPORT_PATH, overwrite = TRUE)
}

reporting <- function(raw_data, init, updates, save) {
    # Workbookオブジェクトを生成
    wb <- init()
    
    # シートを更新
    erm$bindRawData(wb, raw_data)
    
    # 保存
    save(wb)
}

ga.raw_data <- readxl::read_excel(config$RAW_DATA.PATH, sheet=2) %>% 
    dplyr::mutate(日付 = as.integer(format(.$日付)))

reporting(ga.raw_data, init, updates, save)
