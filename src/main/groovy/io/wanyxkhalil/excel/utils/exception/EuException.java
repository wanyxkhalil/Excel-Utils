package io.wanyxkhalil.excel.utils.exception;

/**
 * 系统通用异常
 */
public class EuException extends RuntimeException {

    private EuException(String message) {
        super(message);
    }

    public static EuException build(String messgae) {
        return new EuException(messgae);
    }
}
