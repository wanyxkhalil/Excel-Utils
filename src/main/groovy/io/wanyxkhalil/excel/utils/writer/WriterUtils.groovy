package io.wanyxkhalil.excel.utils.writer

import io.wanyxkhalil.excel.utils.domain.ExcelField
import io.wanyxkhalil.excel.utils.domain.FieldInfo
import io.wanyxkhalil.excel.utils.domain.SheetInfo
import org.apache.commons.collections4.CollectionUtils
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.streaming.SXSSFWorkbook

import java.nio.file.Files
import java.nio.file.Path

class WriterUtils {

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
    static Path write2File(Path path, List<SheetInfo> sheets) {
        if (!Files.exists(path)) {
            path = Files.createFile(path)
        }

        if (CollectionUtils.isEmpty(sheets)) {
            return path
        }

        def book = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE)

        sheets.each {
            sheetWrite(it, book)
        }

        book.write(Files.newOutputStream(path))

        path
    }

    /**
     * 写入各表数据
     * @param sheetInfo 表数据
     * @param book excel文件
     */
    private static void sheetWrite(SheetInfo sheetInfo, Workbook book) {
        Sheet sheet = book.createSheet(sheetInfo.name)
        def clz = sheetInfo.clz
        def data = sheetInfo.data

        if (!data) {
            return
        }

        // 获取字段信息
        def fieldInfo = listFieldInfo(clz)

        // 添加标题行
        def headRow = sheet.createRow(0)
        fieldInfo.eachWithIndex { it, i ->
            def cell = headRow.createCell(i)
            cell.cellValue = it.annotationValue
        }

        // 添加行
        data.eachWithIndex { Object obj, int i ->
            def row = sheet.createRow(i + 1)

            // 每一行设置各单元格
            fieldInfo.eachWithIndex { FieldInfo field, int j ->
                def cell = row.createCell(j)

                def v = obj.invokeMethod(field.getFieldGetter(), null)

                // 非数值则使用String
                def s = v?.toString()
                if (s && !s.isNumber()) {
                    v = s
                }

                cell.cellValue = v
            }
        }
    }

    /**
     * 获取字段信息，父类字段在前，子类字段在后
     * @param clz 类
     * @return 字段信息列表
     */
    private static List<FieldInfo> listFieldInfo(Class clz) {
        def list = new ArrayList<FieldInfo>()

        // 遍历类继承关系
        for (; clz != Object.class; clz = clz.getSuperclass()) {
            def fields = clz.declaredFields

            // 筛选excel字段
            def annotationFields = fields.findAll {
                it.getAnnotation(ExcelField.class) != null
            }

            // 获取字段注解和字段get方法
            def fieldInfo = annotationFields.collect {
                def field = new FieldInfo()

                def annotation = it.getAnnotation(ExcelField.class)
                field.annotationValue = annotation.value()
                field.fieldGetter = retrieveFieldGetter(it.name)

                field
            }

            list.addAll(0, fieldInfo)
        }

        list
    }

    private static String retrieveFieldGetter(String filed) {
        return "get" + filed.substring(0, 1).toUpperCase() + filed.substring(1)
    }
}
