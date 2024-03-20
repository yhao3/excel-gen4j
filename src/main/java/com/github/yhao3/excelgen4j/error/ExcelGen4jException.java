package com.github.yhao3.excelgen4j.error;

/**
 * The common exception to Excel Generator for Java.
 */
public class ExcelGen4jException extends RuntimeException {

    public ExcelGen4jException(String message) {
        super(message);
    }

    public ExcelGen4jException(String message, Throwable cause) {
        super(message, cause);
    }
}
