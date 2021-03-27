package com.retheviper.excel.definition

import com.retheviper.excel.definition.annotation.Column
import com.retheviper.excel.definition.annotation.Sheet
import com.retheviper.excel.definition.type.DataType
import com.retheviper.excel.worker.BookWorker
import org.apache.poi.ss.util.CellAddress
import java.lang.reflect.Field
import java.time.LocalDateTime
import java.util.*

/**
 * Definition of entire excel file, book.
 */
open class BookDef {

    /** Collection of Sheet definitions. */
    val sheetDefs: MutableList<SheetDef> = ArrayList()

    /**
     * Add definition of sheet based on current data class. The data class must be contains [Sheet] and [Column].
     *
     * @param dataClass
     */
    fun addSheet(dataClass: Class<*>) =
        sheetDefs.add(dataClassToSheetDef(dataClass))

    /**
     * Remove definition of sheet based on current data class.
     *
     * @param dataClass
     */
    fun removeSheet(dataClass: Class<*>?) =
        sheetDefs.removeIf { it.dataClass == dataClass }

    /**
     * Find and returns definition of sheet based on current data class.
     * If null, [NullPointerException] will be thrown.
     *
     * @param dataClass
     * @return
     */
    fun getSheetDef(dataClass: Class<*>?) =
        sheetDefs.first { it.dataClass == dataClass }

    /**
     * Create BookWorker for read data from excel file.
     *
     * @param path
     * @return
     */
    fun openBook(path: String) =
        BookWorker(
            bookDef = this,
            inputPath = path
        )

    /**
     * Create BookWorker for write data with template excel file.
     *
     * @param templatePath
     * @param outputPath
     * @return
     */
    fun openBook(templatePath: String, outputPath: String) =
        BookWorker(
            bookDef = this,
            inputPath = templatePath,
            outputPath = outputPath
        )

    /**
     * Create [SheetDef] from data class.
     *
     * @param dataClass
     * @return
     */
    private fun dataClassToSheetDef(dataClass: Class<*>): SheetDef {
        require(dataClass.isAnnotationPresent(Sheet::class.java)) {
            "Missing annotation: $dataClass.simpleName"
        }
        val sheet: Sheet = dataClass.getAnnotation(Sheet::class.java)
        return SheetDef(
            rowDef = RowDef(dataClass.declaredFields
                .filter {
                    it.isAnnotationPresent(Column::class.java)
                }.map { fieldToCellDef(it) }),
            name = if (sheet.name.isNotBlank()) sheet.name else dataClass.simpleName,
            index = sheet.index,
            dataStartIndex = sheet.dataStartIndex,
            dataClass = dataClass
        )
    }

    /**
     * Create [CellDef] from field.
     *
     * @param field
     * @return
     */
    private fun fieldToCellDef(field: Field) =
        field.apply { isAccessible = true }.let {
            val column: Column = it.getAnnotation(Column::class.java)
            CellDef(
                name = if (column.name.isNotBlank()) column.name else it.name,
                type = defineFieldType(it.type),
                cellIndex = CellAddress(column.position + 1).column
            )
        }

    /**
     * Define field's data type.
     *
     * @param fieldType
     * @return
     */
    private fun defineFieldType(fieldType: Class<*>) =
        when (fieldType) {
            Integer.TYPE, Int::class.java -> DataType.INTEGER
            java.lang.Long.TYPE, Long::class.java -> DataType.LONG
            java.lang.Double.TYPE, Double::class.java -> DataType.DOUBLE
            Date::class.java -> DataType.DATE
            LocalDateTime::class.java -> DataType.LOCAL_DATE_TIME
            java.lang.Boolean.TYPE, Boolean::class.java -> DataType.BOOLEAN
            else -> DataType.STRING
        }


    companion object {
        /**
         * Add definition of sheet based on current data classes. The data class must be contains [Sheet] and [Column].
         *
         * @param dataClasses
         * @return BookDef
         */
        fun of(dataClasses: Iterable<Class<*>>) =
            BookDef().apply {
                dataClasses.forEach { addSheet(it) }
            }


        /**
         * Add definition of sheet based on current data classes. The data class must be contains [Sheet] and [Column].
         *
         * @param dataClasses
         * @return BookDef
         */
        fun of(vararg dataClasses: Class<*>) =
            BookDef().apply {
                dataClasses.forEach { addSheet(it) }
            }
    }
}