package com.retheviper.excel.worker

import com.retheviper.excel.definition.CellDef
import com.retheviper.excel.definition.RowDef
import org.apache.poi.ss.usermodel.Row


/**
 * Worker for manipulate row.
 */
class RowWorker(

    /** Definition of row. */
    private val rowDef: RowDef,

    /** Apache POI's row class. */
    private val row: Row
) {
    /** List of [CellWorker]. */
    private val cellWorkers: List<CellWorker> = rowDef.cellDefs.map { getCellWorker(it.cellIndex) }

    /** Current index of cell. Increases when [RowWorker.nextCell] called. */
    private var currentCellIndex: Int = rowDef.cellDefs[0].cellIndex

    /**
     * Get next [CellWorker].
     *
     * @return
     */
    private fun nextCell() = getCellWorker(currentCellIndex++)

    /**
     * Get [CellWorker] with current index.
     *
     * @param cellIndex
     * @return
     */
    private fun getCellWorker(cellIndex: Int) = CellWorker(getCellDef(cellIndex), row.getCell(cellIndex))

    /**
     * Get [CellDef] with current index.
     *
     * @param cellIndex
     * @return
     */
    private fun getCellDef(cellIndex: Int) = rowDef.cellDefs.first { it.cellIndex == cellIndex }

    /**
     * Read row's value and write into object's field.
     *
     * @param obj
     * @param <T>
    </T> */
    fun <T> rowToObj(obj: T) = cellWorkers.forEach { it.cellToObj(obj) }

    /**
     * Read object's field value and write into cell's field.
     *
     * @param obj
     * @param <T>
    </T> */
    fun <T> objToRow(obj: T) = cellWorkers.forEach { it.objToCell(obj) }
}