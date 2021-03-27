package com.retheviper.excel.worker

import com.retheviper.excel.definition.CellDef
import com.retheviper.excel.definition.type.DataType
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DateUtil
import java.lang.reflect.Field
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.*


/**
 * Worker for manipulate cell.
 */
class CellWorker(

    /** Definition of cell. */
    private val cellDef: CellDef,

    /** Apache POI's cell class. */
    private val cell: Cell
) {

    /**
     * Read cell's value and write into object's field.
     *
     * @param obj
     * @param <T>
     * @return
    </T> */
    fun <T> cellToObj(obj: T): T? {
        try {
            getObjectField(obj).let {
                when (cellDef.type) {
                    DataType.INTEGER -> it.set(obj, cell.numericCellValue.toInt())
                    DataType.LONG -> it.set(obj, cell.numericCellValue.toLong())
                    DataType.DOUBLE -> it.set(obj, cell.numericCellValue)
                    DataType.STRING -> it.set(
                        obj,
                        if (cell.stringCellValue.isBlank()) null else cell.stringCellValue
                    )
                    DataType.BOOLEAN -> it.set(obj, cell.booleanCellValue)
                    DataType.LOCAL_DATE_TIME -> it.set(obj, cell.localDateTimeCellValue)
                    DataType.LOCAL_DATE -> it.set(obj, cell.localDateTimeCellValue.toLocalDate())
                    DataType.DATE -> it.set(obj, DateUtil.getJavaDate(cell.numericCellValue))
                }
            }
        } catch (e: Exception) {
            return null
        }
        return obj
    }

    /**
     * Read object's field value and write into cell's field.
     *
     * @param obj
     * @param <T>
     * @return
    </T> */
    fun <T> objToCell(obj: T) {
        try {
            getObjectField(obj).get(obj).let {
                when (cellDef.type) {
                    DataType.INTEGER -> cell.setCellValue((it as Int).toDouble())
                    DataType.LONG -> cell.setCellValue((it as Long).toDouble())
                    DataType.DOUBLE -> cell.setCellValue(it as Double)
                    DataType.STRING -> cell.setCellValue(it as String)
                    DataType.BOOLEAN -> cell.setCellValue(it as Boolean)
                    DataType.LOCAL_DATE_TIME -> cell.setCellValue(it as LocalDateTime)
                    DataType.LOCAL_DATE -> cell.setCellValue(it as LocalDate)
                    DataType.DATE -> cell.setCellValue(DateUtil.getExcelDate(it as Date))
                }
            }
        } catch (e: Exception) {
            return
        }
    }

    /**
     * Get field from obj.
     *
     * @param obj
     * @param <T>
     * @return
    </T> */
    private fun <T> getObjectField(obj: T): Field =
        obj!!::class.java.getDeclaredField(cellDef.name).apply {
            isAccessible = true
        }
}