package io.wanyxkhalil.excel.utils.dynamicwriter;

import io.wanyxkhalil.excel.utils.constant.Constants;
import io.wanyxkhalil.excel.utils.domain.SheetInfo;
import io.wanyxkhalil.excel.utils.util.FileUtils;
import org.apache.commons.collections4.CollectionUtils;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class DynamicWriterBuilder {

    /**
     * 文件名
     */
    private String fileName;

    /**
     * 表单
     */
    private List<SheetInfo> sheets;

    public DynamicWriterBuilder fileName() {
        return fileName(Constants.BLANK_STRING);
    }

    public DynamicWriterBuilder fileName(String fileName) {
        this.fileName = fileName;
        return this;
    }

    public DynamicWriterBuilder sheet(List<DynamicData> list) {
        int sheetNum = retrieveSheetNum();
        return sheet(Constants.DEFAULT_SHEET_NAME + sheetNum, list);
    }

    public DynamicWriterBuilder sheet(String sheetName, List<DynamicData> list) {
        if (Objects.isNull(sheets)) {
            sheets = new ArrayList<>();
        }

        SheetInfo sheet = new SheetInfo();
        sheet.setName(sheetName);
        sheet.setData(list);

        sheets.add(sheet);

        return this;
    }

    public Path build() {
        Path path = FileUtils.retrievePath(this.fileName);

        return DynamicWriterUtils.write2File(path, this.sheets);
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
}
