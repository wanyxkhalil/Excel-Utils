package io.wanyxkhalil.excel.utils.writer;

import io.wanyxkhalil.excel.utils.constant.Constants;
import io.wanyxkhalil.excel.utils.domain.SheetInfo;
import io.wanyxkhalil.excel.utils.util.FileUtils;
import org.apache.commons.collections4.CollectionUtils;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class WriterBuilder {

    /**
     * 文件名
     */
    private String fileName;

    /**
     * 表单
     */
    private List<SheetInfo> sheets;

    public WriterBuilder fileName() {
        return fileName(Constants.BLANK_STRING);
    }

    public WriterBuilder fileName(String fileName) {
        this.fileName = fileName;
        return this;
    }

    public WriterBuilder sheet(List list) {
        int sheetNum = retrieveSheetNum();
        return sheet(Constants.DEFAULT_SHEET_NAME + sheetNum, list);
    }

    public WriterBuilder sheet(String sheetName, List list) {
        Class clz = retrieveClass(list);
        return sheet(sheetName, clz, list);
    }

    public WriterBuilder sheet(String sheetName, Class clz, List list) {
        if (Objects.isNull(sheets)) {
            sheets = new ArrayList<>();
        }

        SheetInfo sheet = new SheetInfo();
        sheet.setName(sheetName);
        sheet.setClz(clz);
        sheet.setData(list);

        sheets.add(sheet);

        return this;
    }

    public Path build() {
        Path path = FileUtils.retrievePath(this.fileName);

        return WriterUtils.write2File(path, this.sheets);
    }

    /**
     * 获取默认的sheet数。如1
     */
    private int retrieveSheetNum() {
        if (CollectionUtils.isEmpty(sheets)) {
            return 1;
        }
        return sheets.size() + 1;
    }

    /**
     * 根据list获取class
     */
    private Class retrieveClass(List list) {
        if (Objects.isNull(list) || list.size() == 0) {
            return Object.class;
        }
        return list.get(0).getClass();
    }
}
