package com.retheviper.excel.worker


import com.retheviper.excel.definition.BookDef
import com.retheviper.excel.definition.SheetDef
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.nio.file.Files
import java.nio.file.Path


/**
 * Worker for manipulate excel file.
 */
open class BookWorker(

    /**
     * Definition of book.
     */
    private val bookDef: BookDef,

    /**
     * Path for input.
     */
    private val inputPath: String,

    /**
     * Path for output.
     */
    private var outputPath: String? = null
) : AutoCloseable {

    /**
     * Apache POI's workbook class.
     */
    private val workbook: Workbook = WorkbookFactory.create(File(inputPath))

    /**
     * Path for output.
     */
    private var output: Path? = null

    init {
        outputPath?.let { output = Path.of(it) }
    }

    /**
     * Get [SheetWorker].
     *
     * @param sheetName
     * @return SheetWorker
     */
    private fun getSheetWorker(sheetName: String?) = searchSheetDef { it.name == sheetName }

    /**
     * Get [SheetWorker].
     *
     * @param index
     * @return SheetWorker
     */
    private fun getSheetWorker(index: Int) = searchSheetDef { it.index == index }

    /**
     * Get [SheetWorker].
     *
     * @param dataClass
     * @return SheetWorker
     */
    private fun getSheetWorker(dataClass: Class<*>) = searchSheetDef { it.dataClass == dataClass }

    /**
     * Search [SheetDef] with condition and map to [SheetWorker].
     *
     * @param condition
     * @return
     */
    private fun searchSheetDef(condition: (target: SheetDef) -> Boolean) =
        bookDef.sheetDefs.first(condition).let { getSheetWorker(it) }

    /**
     * Get [SheetWorker].
     *
     * @param sheetDef
     * @return
     */
    private fun getSheetWorker(sheetDef: SheetDef) =
        SheetWorker(
            sheetDef,
            if (sheetDef.index > -1) workbook.getSheetAt(sheetDef.index) else workbook.getSheet(sheetDef.name)
        )

    /**
     * Write data into sheet.
     *
     * @param data
     * @param <T>
    </T> */
    fun <T> write(data: List<T>) =
        getSheetWorker(data.first()!!::class.java).listToSheet(data)

    /**
     * Read data from sheet as list.
     *
     * @param sheetName
     * @param <T>
     * @return
    </T> */
    fun <T> read(sheetName: String?): List<T> {
        return getSheetWorker(sheetName).sheetToList()
    }

    /**
     * Read data from sheet as list.
     *
     * @param index
     * @param <T>
     * @return
    </T> */
    fun <T> read(index: Int): List<T> {
        return getSheetWorker(index).sheetToList()
    }

    /**
     * Read data from sheet as list.
     *
     * @param dataClass
     * @param <T>
     * @return
    </T> */
    fun <T> read(dataClass: Class<T>): List<T> =
        getSheetWorker(dataClass).sheetToList()

    override fun close() {
        if (output != null) {
            output!!.let {
                if (Files.notExists(it.parent)) {
                    Files.createDirectories(it.parent)
                }
                Files.newOutputStream(it).use { stream -> workbook.write(stream) }
            }
        } else {
            workbook.close()
        }
    }
}