package com.retheviper.excel.test.testbase

import com.retheviper.excel.test.header.Contract
import java.text.DateFormat
import java.text.DecimalFormat
import java.text.SimpleDateFormat
import java.util.*
import java.util.concurrent.ThreadLocalRandom


object TestBase {

    private val NUMBER_FORMAT = DecimalFormat("0000")

    private val DATE_FORMAT: DateFormat = SimpleDateFormat("yyyy/MM/dd")

    private const val ADDRESS = "somewhere"

    private const val CITY = "some city"

    private const val POST = 90000

    private const val DATA_AMOUNT = 10000

    private const val START_DATE = "1970/01/01"

    private const val END_DATE = "2000/12/31"

    val testData: List<Contract> = generateSequence(0) { it + 1 }
        .takeWhile { it < DATA_AMOUNT }
        .map {
            val formatted = NUMBER_FORMAT.format(it)
            val phone = "(000) 0000-$formatted"
            Contract(
                name = "People$formatted",
                companyPhone = phone,
                cellPhone = phone,
                homePhone = phone,
                email = "$formatted@example.com",
                birth = createRandomDate(),
                address = ADDRESS,
                city = CITY,
                province = ADDRESS,
                post = POST + it
            )
        }.toList()

    private fun createRandomDate() = Calendar.getInstance().apply {
        time = Date(
            ThreadLocalRandom
                .current()
                .nextLong(DATE_FORMAT.parse(START_DATE).time, DATE_FORMAT.parse(END_DATE).time)
        )
        this[Calendar.HOUR] = 0
        this[Calendar.MINUTE] = 0
        this[Calendar.SECOND] = 0
        this[Calendar.HOUR_OF_DAY] = 0
    }.time

}