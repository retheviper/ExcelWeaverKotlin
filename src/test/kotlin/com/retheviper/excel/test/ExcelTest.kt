package com.retheviper.excel.test

import com.retheviper.excel.test.header.Contract
import com.retheviper.excel.test.testbase.TestBase
import com.retheviper.excel.definition.BookDef
import io.kotlintest.matchers.shouldBe
import io.kotlintest.specs.FreeSpec
import java.util.stream.IntStream


class ExcelTest : FreeSpec() {

    private val templatePath = "src/test/resources/template/tf03783921_win321.xlsx"

    private val outputPath = "src/test/resources/temp/output.xlsx"

    private val bookDef = BookDef().apply { addSheet(Contract::class.java) }

    private val testData: List<Contract> = TestBase.testData

    init {
        "writeTest" {
            bookDef.openBook(templatePath, outputPath).use {
                it.write(testData)
            }
        }

        "readTest" {
            val list: List<Contract>
            bookDef.openBook(outputPath).use { list = it.read(Contract::class.java) }
            IntStream.range(0, testData.size).forEach { index: Int ->
                val expected = testData[index]
                val actual = list[index]
                actual.name shouldBe expected.name
                actual.companyPhone shouldBe expected.companyPhone
                actual.cellPhone shouldBe expected.cellPhone
                actual.homePhone shouldBe expected.homePhone
                actual.email shouldBe expected.email
                actual.birth shouldBe expected.birth
                actual.address shouldBe expected.address
                actual.city shouldBe expected.city
                actual.province shouldBe expected.province
                actual.post shouldBe expected.post
                actual.memo shouldBe expected.memo
            }
        }
    }
}