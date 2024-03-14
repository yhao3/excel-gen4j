package com.github.yhao3.excelpro.error;

/**
 * The common exception to ExcelPro SDK.
 */
public class ExcelProException extends RuntimeException {

    public ExcelProException(String message) {
        super(message);
    }

    public ExcelProException(String message, Throwable cause) {
        super(message, cause);
    }
}
