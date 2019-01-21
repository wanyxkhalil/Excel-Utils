package io.wanyxkhalil.excel.utils.domain;

/**
 * excel utils 异常
 */
public class ExcelException extends RuntimeException {

    private ExcelException(String message) {
        super(message);
    }

    public static ExcelException build(String message) {
        return new ExcelException(message);
    }
}
