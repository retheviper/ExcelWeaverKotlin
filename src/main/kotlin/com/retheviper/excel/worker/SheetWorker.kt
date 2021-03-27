package com.retheviper.excel.worker

import com.retheviper.excel.definition.RowDef
import com.retheviper.excel.definition.SheetDef
import org.apache.poi.ss.usermodel.Sheet
import java.time.LocalDate
import java.util.*


/**
 * Worker for manipulate sheet.
 */
class SheetWorker(

    /** Definition of sheet. */
    private val sheetDef: SheetDef,

    /** Apache POI's sheet class. */
    private val sheet: Sheet
) {
    /** Current index of row. Increases when [SheetWorker.nextRow] called. */
    private var currentRowIndex: Int = sheetDef.dataStartIndex

    /**
     * Get next [RowWorker].
     *
     * @return
     */
    private fun nextRow() = getRowWorker(currentRowIndex++)

    /**
     * Get [RowWorker] with current index.
     *
     * @param rowIndex
     * @return
     */
    private fun getRowWorker(rowIndex: Int) = RowWorker(rowDef, getRow(rowIndex))

    /**
     * Get Row.
     *
     * @param rowIndex
     * @return
     */
    private fun getRow(rowIndex: Int) = if (sheet.getRow(rowIndex) != null)
        sheet.getRow(rowIndex)
    else
        sheet.createRow(rowIndex).apply {
            sheet.getRow(sheetDef.dataStartIndex)
                .forEach {
                    createCell(it.columnIndex)
                        .also { cell -> cell.cellStyle = it.cellStyle }
                }
        }
            .also { it.rowStyle = sheet.getRow(sheetDef.dataStartIndex).rowStyle }

    /**
     * Get definition of row.
     *
     * @return
     */
    private val rowDef: RowDef
        get() = sheetDef.rowDef

    /** d
     * Read objects value and write into sheet.
     *
     * @param data
     * @param <T>
    </T> */
    fun <T> listToSheet(data: List<T>) = data.forEach { nextRow().objToRow(it) }

    /**
     * Read sheet's value and write into objects.
     *
     * @param <T>
     * @return
    </T> */
    fun <T> sheetToList(): List<T> =
        generateSequence(sheetDef.dataStartIndex) { it + 1 }.takeWhile { it <= sheet.lastRowNum }
            .map {
                val constructor = sheetDef.dataClass.declaredConstructors.last()
                val args: Array<*> = constructor.parameters
                    .map {
                        when (it.type) {
                            Integer.TYPE, java.lang.Long.TYPE, java.lang.Double.TYPE -> 0
                            java.lang.Boolean.TYPE -> true
                            String::class.java -> ""
                            Date::class.java -> Date()
                            LocalDate::class.java -> LocalDate.now()
                            else -> it.type.constructors.firstOrNull { c -> !c.isVarArgs }?.let { c -> c.newInstance() }
                        }
                    }.toTypedArray()
                constructor.newInstance(*args).apply {
                    getRowWorker(it).rowToObj(this)
                } as T
            }.toList()
}