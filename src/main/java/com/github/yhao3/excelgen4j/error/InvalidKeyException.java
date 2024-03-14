package com.github.yhao3.excelgen4j.error;

/**
 * The invalid key exception, e.g. {{}}.
 */
public class InvalidKeyException extends ExcelGen4jException {

    public InvalidKeyException(String message) {
        super(message);
    }

    public InvalidKeyException(String message, Throwable cause) {
        super(message, cause);
    }
}
