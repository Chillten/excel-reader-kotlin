package com.bogovich.excel.callback

import com.bogovich.excel.ExcelItemWriter
import com.bogovich.excel.ExecutionContext
import com.bogovich.excel.ExecutionContextAware

class DefaultExcelReadCallback<T> : ExcelReadCallback,
        ExecutionContextAware {

    override var executionContext: ExecutionContext? = null

    var writer: ExcelItemWriter<T>? = null

    override fun beforeReading() {
        executionContext?.clean()
    }

    override fun afterReading() {
        this.writeLastItem()
    }

    private fun writeLastItem() {
        if(executionContext?.chunkItems?.isNotEmpty()!!) {
            executionContext?.let { writer?.writeLastItems(it) }
        }
    }
}