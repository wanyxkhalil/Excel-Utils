package io.wanyxkhalil.excel.utils.normal

import io.wanyxkhalil.excel.utils.domain.FieldInfo
import io.wanyxkhalil.excel.utils.domain.SheetObject
import io.wanyxkhalil.excel.utils.util.ExcelBuilder
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.streaming.SXSSFWorkbook

import java.nio.file.Files
import java.nio.file.Path

class ExcelUtil {

    /**
     * 内存驻留行数
     */
    protected static final int ROW_ACCESS_WINDOW_SIZE = 1000

    /**
     * 新建ExcelBuilder，用以创建ExcelObject
     * @return ExcelBuilder对象
     */
    static ExcelBuilder builder() {
        new ExcelBuilder()
    }

    /**
     * 生成excel文件
     * @param path 文件路径
     * @param sheets 表数据
     * @return 文件路径
     */
    static Path write2File(Path path, List<SheetObject> sheets) {
        if (!Files.exists(path)) {
            path = Files.createFile(path)
        }

        if (!Files.isWritable(path)) {
            throw new RuntimeException("该地址不可访问")
        }

        if (Objects.isNull(sheets) || sheets.size() == 0) {
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
     * @param sheetObject 表数据
     * @param book excel文件
     */
    private static void sheetWrite(SheetObject sheetObject, Workbook book) {
        Sheet sheet = book.createSheet(sheetObject.sheetName)
        def clz = sheetObject.clz
        def data = sheetObject.data

        if (!data) {
            return
        }

        // 获取字段信息
        def fieldInfo = listFieldInfo(clz)

        // 添加标题行
        def headRow = sheet.createRow(0)
        fieldInfo.eachWithIndex { it, i ->
            def cell = headRow.createCell(i)
            cell.cellValue = it.annotationName
        }

        // 添加行
        data.eachWithIndex { Object obj, int i ->
            def row = sheet.createRow(i + 1)

            // 每一行设置各单元格
            fieldInfo.eachWithIndex { FieldInfo field, int j ->
                def cell = row.createCell(j)

                def v = obj.invokeMethod(field.fieldNameGetter, null)

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
                field.annotationName = annotation.value()

                field.fieldNameGetter = methodNameGet(it.name)

                field
            }

            list.addAll(0, fieldInfo)
        }

        list
    }

    private static String methodNameGet(String filed) {
        return "get" + filed.substring(0, 1).toUpperCase() + filed.substring(1)
    }
}
