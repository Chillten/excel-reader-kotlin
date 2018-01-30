package com.bogovich.excel

data class ExcelFileReaderError(val code: String,
                                val defaultMessage: String,
                                val sheetName: String,
                                val rowIndex: Int,
                                val columnReference: String)