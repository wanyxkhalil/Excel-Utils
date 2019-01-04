package io.wanyxkhalil.excel.utils

import io.wanyxkhalil.excel.utils.domain.SheetObject
import io.wanyxkhalil.excel.utils.dynamic.DynamicExcelBuilder
import io.wanyxkhalil.excel.utils.dynamic.GroovyDynamicData
import io.wanyxkhalil.excel.utils.exception.EuException
import org.apache.commons.collections4.CollectionUtils
import org.apache.poi.xssf.streaming.SXSSFWorkbook

import java.nio.file.Files
import java.nio.file.Path

class DynamicExcelUtil extends ExcelUtil {

    static DynamicExcelBuilder builder() {
        new DynamicExcelBuilder()
    }

    static Path write2File(Path path, List<SheetObject> sheets) {

        if (!Files.exists(path)) {
            path = Files.createFile(path)
        }

        if (!Files.isWritable(path)) {
            throw EuException.build("该地址不可访问")
        }

        if (CollectionUtils.sizeIsEmpty(sheets)) {
            throw EuException.build(("生成动态Excel时不可为空"))
        }

        if (!Objects.equals(sheets.get(0).clz, GroovyDynamicData.class)) {
            throw EuException.build(("使用动态数据生成Excel时不可为空"))
        }

        // Excel生成
        def book = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE)

        sheets.forEach {
            sheetWrite(book, it)
        }

        book.write(Files.newOutputStream(path))

        path
    }

    @SuppressWarnings("deprecation")
    private static void sheetWrite(SXSSFWorkbook book, SheetObject sheetObj) {
        def sheet = book.createSheet(sheetObj.sheetName)
        def headRow = sheet.createRow(0)

        def lst = sheetObj.data

        // 获取属性
        def d = lst.max { it.map.size() }
        def fields = d?.map?.keySet()

        lst.each { data ->
            data.map.keySet().each {
                if (!fields.contains(it)) {
                    fields += it
                }
            }
        }

        // 设置文件第一行
        fields.eachWithIndex { field, i ->
            def cell = headRow.createCell(i)
            cell.setCellValue(field)
        }

        // 设置值
        lst.eachWithIndex { elem, i ->
            def nextRow = sheet.createRow(i + 1)

            fields.eachWithIndex { field, index ->
                def cell = nextRow.createCell(index)
                def var = elem."$field"
                if (Objects.nonNull(var)) {
                    if (var instanceof Number) {
                        cell.setCellValue(var)
                    } else {
                        cell.setCellValue(var.toString())
                    }
                }
            }
        }
    }
}
