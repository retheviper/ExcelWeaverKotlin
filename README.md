# ExcelWeaverKotlin

![excelweaver_kotlin](./misc/excelweaver_kotlin.jpg)

A simple library for manipulate excel file in Kotlin.

basically, this is just another implementation version of [ExcelWeaver](https://github.com/retheviper/ExcelWeaver).

## TL;DR

`*.xlsx` → `List`

`List` → `*.xlsx`

## To build

`./gradlew shadowJar`

(And Jar will be found in `build/libs`)

## To Use

### Make definition of sheet

```kotlin
@Sheet(dataStartIndex = 2)
data class Contract (// Class name will be sheets name

    @Column(position = "B")
    val name: String,

    @Column(position = "C")
    val cellPhone: String,

    @Column(position = "D")
    val postCode: Int
)
```

### Make definition of book

```kotlin
// Create from array or list of SheetDef classes
val bookDef = BookDef.of(Contract::class.java, Message::class.java);

// add more sheet
bookDef.addSheet(Salary::class.java);
```

### Read file

```kotlin
val list: List<Contract>
bookDef.openBook(outputPath).use { list = it.read(Contract::class.java) }
```

### Write file

```kotlin
val list = createData()
bookDef.openBook(templatePath, outputPath).use { it.write(list) }
```
