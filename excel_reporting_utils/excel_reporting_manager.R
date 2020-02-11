#'-------------------------------------------------------------------------------
#' @description excel形式でレポート出力
#' @author masaya.genshin@gmail.com
#'-------------------------------------------------------------------------------
excel_reporting_manager <- function() {
    START_ROW = 3
    startCol = 2
    mapping_list <- list() # 元データの列名のmappingを行う
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
    
    # xlsxのsum関数を作成する
    xlsx.SUMIFS <- function(target, filters=NULL) {
        params <- Reduce(
            function(acc, cur) paste0(acc, ', ', cur[1], ', ', cur[2])
            , x=filters, init=target)
        return(stringr::str_glue('{ifelse(is.null(filters), "=SUM(", "=SUMIFS(")}{params})'))
    }
    
    # リストに要素を追加する
    add <- function(l, new) {
        tmp <- c(l, list(0))
        tmp[[length(l) + 1]] <- new
        return(tmp)
    }
    
    # サマリーレポートを作成する
    addSummaryReport <- function(raw_data_sheet='raw_data', metric, calculated_values=list(), segment_column=NULL, filter_column=NULL) {
        # 元データの対応(リスト)を取得
        mapping <- mapping_list[[raw_data_sheet]]
        COLUMNS <- lapply(c(1:length(mapping)), function(i) LETTERS[[i]])
        names(COLUMNS) <- names(mapping)
        
        updateSummaryReport <- function(wb, raw_data) {
            if(!('summary' %in% names(wb))) {
                openxlsx::addWorksheet(wb, 'summary')
            }
            raw_data.ROWS <- nrow(raw_data)
            calculated_columns <- names(calculated_values)
            
            segment_id_list <- NULL
            if (!is.null(segment_column)) {
                # セグメント識別子(リスト)
                segment_id_list <- as.character(sort(unlist(unique(raw_data[,mapping[[segment_column]]]))))
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
                                filters <- list(c('{raw_data_sheet}!${COLUMNS[[segment_column]]}2:${COLUMNS[[segment_column]]}{raw_data.ROWS + 1}', 'IF($A{row}="total", "*", $A{row})'))
                            }
                            if(!is.null(filter_column)) {
                                filters <- append(filters, list(c('{raw_data_sheet}!${COLUMNS[[filter_column]]}2:${COLUMNS[[filter_column]]}{raw_data.ROWS + 1}', 'B1')))
                            }
                            return(stringr::str_glue(xlsx.SUMIFS('{raw_data_sheet}!{column}$2:{column}${raw_data.ROWS + 1}', filters)))
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
            # filter
            openxlsx::writeData(wb, 'summary', '*', startCol = 2, startRow = 1)
        }
        updates <<- add(updates, list(f=updateSummaryReport, raw_data_sheet=raw_data_sheet))
    }
    
    # 期間レポートを作成する
    addDuringReport <- function(raw_data_sheet='raw_data', metric, calculated_values=list(), during_list, filter_column=NULL) {
        # 元データの対応(リスト)を取得
        mapping <- mapping_list[[raw_data_sheet]]
        COLUMNS <- lapply(c(1:length(mapping)), function(i) LETTERS[[i]])
        names(COLUMNS) <- names(mapping)
        
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
                                c('{raw_data_sheet}!${COLUMNS$date}2:${COLUMNS$date}{raw_data.ROWS + 1}', '">= " & LEFT($A{row},8)'),
                                c('{raw_data_sheet}!${COLUMNS$date}2:${COLUMNS$date}{raw_data.ROWS + 1}', '"<= " & RIGHT($A{row},8)')
                            )
                            if (!is.null(filter_column)) {
                                filters <- append(filters, list(c('{raw_data_sheet}!${COLUMNS[[filter_column]]}2:${COLUMNS[[filter_column]]}{raw_data.ROWS + 1}', 'IF($B$1="total", "*", $B$1)')))
                            }
                            return(stringr::str_glue(xlsx.SUMIFS('{raw_data_sheet}!{column}$2:{column}${raw_data.ROWS + 1}', filters)))
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
        updates <<- add(updates, list(f=updateDuringReport, raw_data_sheet=raw_data_sheet))
    }
    
    addRawData <- function(name='raw_data', mapping) {
        .mapping <- list(1)
        names(.mapping) <- name
        .mapping[[name]] <- mapping 
        mapping_list <<- append(mapping_list, .mapping)
        
        updateRaeData <- function(wb, raw_data) {
            if(!(name %in% names(wb))) {
                openxlsx::addWorksheet(wb, name)
            }
            # データ書き込み
            openxlsx::writeData(wb, raw_data, sheet=name)
        }
        
        updates <<- add(updates, list(f=updateRaeData, raw_data_sheet=name))
    }
    
    bindRawData <- function(wb, rawData) {
        if(is.data.frame(rawData)) {
            .rawData <- list(raw_data=rawData)
        } else {
            .rawData <- rawData
        }
        
        for (update in updates) {
            (update[['f']])(wb, .rawData[[update[['raw_data_sheet']]]])
        }
    }

    return(list(
        addSummaryReport=addSummaryReport,
        addDuringReport=addDuringReport,
        addRawData=addRawData,
        bindRawData=bindRawData
    ))
}
