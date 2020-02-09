#'-------------------------------------------------------------------------------
#' @description excel形式でレポート出力
#' @author masaya.genshin@gmail.com
#'-------------------------------------------------------------------------------
excel_reporting_manager <- function() {
    START_ROW = 3
    startCol = 2
    MAPPING <- NULL # 元データの列名のmappingを行う
    COLUMNS <- NULL # 元データの列
    updates <- list()
    
    # 表の先頭行を書き込む
    writeHeader <- function(sheetName, header) {
        return(function(wb) {
            df <- as.list(rep(0, length(header)))
            names(df) <- header
            openxlsx::writeData(wb, sheetName, data.frame(df), startRow = START_ROW - 1)
        })
    }
    
    # シートにタイトルを記入する
    writeTitle <- function(sheetName, title) {
        return(function(wb) {
            openxlsx::writeData(wb, sheetName, title, startCol = 1, startRow = 1)
        })
    }
    
    # 元データの列定義を設定する
    setRawDataColumnsMapping <- function(mapping) {
        MAPPING <<- mapping
        COLUMNS <<- lapply(c(1:length(mapping)), function(i) LETTERS[[i]])
        names(COLUMNS) <<- names(mapping)
    }
    
    # xlsxのsum関数を作成する
    xlsx.SUMIFS <- function(target, filters=NULL) {
        params <- Reduce(
            function(acc, cur) paste0(acc, ', ', cur[1], ', ', cur[2])
            , x=filters, init=target)
        return(stringr::str_glue('{ifelse(is.null(filters), "=SUM(", "=SUMIFS(")}{params})'))
    }
    
    # サマリーレポートを作成する
    addSummaryReport <- function(metric, calculated_values=list(), segment_column=NULL) {
        updateSummaryReport <- function(wb, raw_data) {
            if(!('summary' %in% names(wb))) {
                openxlsx::addWorksheet(wb, 'summary')
            }
            raw_data.ROWS <- nrow(raw_data)
            calculated_columns <- names(calculated_values)
            
            segment_id_list <- NULL
            if (!is.null(segment_column)) {
                # セグメント識別子(リスト)
                segment_id_list <- as.character(sort(unlist(unique(raw_data[,MAPPING[[segment_column]]]))))
            }
            segment_id_list <- c(segment_id_list, 'total')
            
            #タイトル
            writeTitle('summary', 'サマリー')(wb)
            # ヘッダーを記入
            header <- c(segment_column, metric, calculated_columns)
            writeHeader('summary',  header)(wb)
            
            # 集計表
            formula_list <- function(row) {
                return(
                    append(
                        lapply(COLUMNS[metric], function(column) {
                            filters <- NULL
                            if(!is.null(segment_column)) {
                                filters <- list(c('raw_data!${COLUMNS[[segment_column]]}2:${COLUMNS[[segment_column]]}{raw_data.ROWS + 1}', 'IF($A{row}="total", "*", $A{row})'))
                            }
                            return(stringr::str_glue(xlsx.SUMIFS('raw_data!{column}$2:{column}${raw_data.ROWS + 1}', filters)))
                        }), 
                        lapply(calculated_values, function(fml) {
                            return(fml(row))
                        })
                    )
                )
            }
            
            for (fml in formula_list(c(START_ROW:(length(segment_id_list) + START_ROW - 1)))) {
                openxlsx::writeFormula(wb, 'summary', fml, startCol=startCol, startRow=START_ROW)
                startCol = startCol + 1
            }
            
            # セグメント識別子をA列に記入
            openxlsx::writeData(wb, 'summary', segment_id_list, startCol = 1, startRow = START_ROW)
        }
        updates <<- append(updates, list(updateSummaryReport=updateSummaryReport))
    }
    
    # 期間レポートを作成する
    addDuringReport <- function(metric, calculated_values=list(), during_list, filter_column=NULL) {
        updateDuringReport <- function(wb, raw_data) {
            if(!('transition' %in% names(wb))) {
                openxlsx::addWorksheet(wb, 'transition')
            }
            raw_data.ROWS <- nrow(raw_data)
            calculated_columns <- names(calculated_values)
            # タイトル記入
            writeTitle('transition', '期間レポート')(wb)
            # ヘッダーを記入
            writeHeader('transition',  c('期間', metric, calculated_columns))(wb)
            
            # 集計表
            formula_list <- function(row) {
                return(
                    append(
                        lapply(COLUMNS[metric], function(column){
                            filters <- list(
                                c('raw_data!${COLUMNS$date}2:${COLUMNS$date}{raw_data.ROWS + 1}', '">= " & LEFT($A{row},8)'),
                                c('raw_data!${COLUMNS$date}2:${COLUMNS$date}{raw_data.ROWS + 1}', '"<= " & RIGHT($A{row},8)')
                            )
                            if (!is.null(filter_column)) {
                                filters <- append(filters, list(c('raw_data!${COLUMNS[[filter_column]]}2:${COLUMNS[[filter_column]]}{raw_data.ROWS + 1}', 'IF($B$1="total", "*", $B$1)')))
                            }
                            return(stringr::str_glue(xlsx.SUMIFS('raw_data!{column}$2:{column}${raw_data.ROWS + 1}', filters)))
                        }), 
                        lapply(calculated_values, function(fml) {
                            return(fml(row))
                        })
                    )
                )
            }
            openxlsx::writeData(wb, 'transition', 'total', startCol = 2, startRow = 1)
            
            for (fml in formula_list(c(START_ROW:(length(names(during_list)) + START_ROW - 1)))) {
                openxlsx::writeFormula(wb, 'transition', fml, startCol = startCol, startRow = START_ROW)
                startCol = startCol + 1
            }
            
            # 期間列をA列に記入
            during_list_str <- sapply(during_list, function(during) {
                    return(stringr::str_glue('{during[1]}_{during[2]}'))
                })
            openxlsx::writeData(wb, 'transition', during_list_str, startCol = 1, startRow = START_ROW)
        }
        updates <<- append(updates, list(updateDuringReport=updateDuringReport))
    }
    
    addRawData <- function() {
        updates <<- append(updates, list(
            updateRawData=function(wb, raw_data) {
                if(!('raw_data' %in% names(wb))) {
                    openxlsx::addWorksheet(wb, 'raw_data')
                }
                # データ書き込み
                openxlsx::writeData(wb, raw_data, sheet='raw_data')
            }
        ))
    }
    
    updateReports <- function(wb, raw_data) {
        for (update in updates) {
            update(wb, raw_data)
        }
    }

    return(list(
        setRawDataColumnsMapping=setRawDataColumnsMapping, 
        addSummaryReport=addSummaryReport,
        addDuringReport=addDuringReport,
        addRawData=addRawData,
        updateReports=updateReports
    ))
}
