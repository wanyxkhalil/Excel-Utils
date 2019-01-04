package io.wanyxkhalil.excel.utils.normal;

import io.wanyxkhalil.excel.utils.domain.ExcelObject;
import io.wanyxkhalil.excel.utils.domain.SheetObject;
import io.wanyxkhalil.excel.utils.util.ExcelFileUtils;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class ExcelBuilder {

    /**
     * 空字符串
     */
    private static final String BLANK_STRING = "";


    private static final String DEFAULT_SHEET_NAME = "sheet";

    /**
     * 文件名
     */
    protected String fileName;

    /**
     * 表单数据
     */
    protected List<SheetObject> sheets;

    public ExcelBuilder fileName() {
        return fileName(BLANK_STRING);
    }

    public ExcelBuilder fileName(String fileName) {
        this.fileName = fileName;
        return this;
    }

    public ExcelBuilder sheet(List data) {
        StringBuilder sheetName = new StringBuilder(DEFAULT_SHEET_NAME);
        if (Objects.nonNull(sheets)) {
            sheetName.append(sheets.size() + 1);
        }

        return sheet(sheetName.toString(), data);
    }

    public ExcelBuilder sheet(String sheetName, List data) {
        Class clz;
        if (Objects.isNull(data) || data.size() == 0) {
            clz = Object.class;
        } else {
            clz = data.get(0).getClass();
        }

        return sheet(sheetName, clz, data);
    }

    public ExcelBuilder sheet(String sheetName, Class clz, List data) {
        if (Objects.isNull(sheets)) {
            sheets = new ArrayList<>();
        }

        SheetObject obj = new SheetObject();
        obj.setSheetName(sheetName);
        obj.setClz(clz);
        obj.setData(data);

        sheets.add(obj);

        return this;
    }

    public ExcelObject build() {
        ExcelObject obj = new ExcelObject();
        obj.setSheets(this.sheets);

        Path path = ExcelFileUtils.getXlsxPath(this.fileName);
        obj.setPath(path);

        return obj;
    }
}