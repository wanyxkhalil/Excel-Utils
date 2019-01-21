package io.wanyxkhalil.excel.utils.dynamicwriter

import io.wanyxkhalil.excel.utils.domain.SheetInfo
import org.apache.commons.collections4.CollectionUtils
import org.apache.poi.xssf.streaming.SXSSFWorkbook

import java.nio.file.Files
import java.nio.file.Path

class DynamicWriterUtils {

    /**
     * 内存驻留行数
     */
    protected static final int ROW_ACCESS_WINDOW_SIZE = 1000

    /**
     * 生成excel文件
     * @param path 文件路径
     * @param sheets 表数据
     * @return 文件路径
     */
    protected static Path write2File(Path path, List<SheetInfo> sheets) {

        if (!Files.exists(path)) {
            path = Files.createFile(path)
        }

        if (CollectionUtils.isEmpty(sheets)) {
            return path
        }

        // Excel生成
        def book = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE)

        sheets.forEach {
            sheetWrite(book, it)
        }

        book.write(Files.newOutputStream(path))

        path
    }

    /**
     * 写入各表数据
     * @param sheetInfo 表数据
     * @param book excel文件
     */
    @SuppressWarnings("all")
    private static void sheetWrite(SXSSFWorkbook book, SheetInfo sheetInfo) {
        def sheet = book.createSheet(sheetInfo.name)

        def lst = sheetInfo.data

        // 获取字段信息
        def fields = retrieveFields(lst)

        // 添加标题行
        def headRow = sheet.createRow(0)
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

    /**
     * 以最大属性数量为基础，添加其他
     * @param lst
     */
    private static retrieveFields(List<DynamicData> lst) {
        def d = lst.max { it.map.size() }
        def fields = d?.map?.keySet()

        lst.each { data ->
            data.map.keySet().each {
                if (!fields.contains(it)) {
                    fields += it
                }
            }
        }

        fields
    }
}
