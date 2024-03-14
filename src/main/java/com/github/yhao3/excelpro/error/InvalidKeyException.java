package com.github.yhao3.excelpro.error;

/**
 * The common exception to ExcelPro SDK.
 */
public class InvalidKeyException extends ExcelProException {

    public InvalidKeyException(String message) {
        super(message);
    }

    public InvalidKeyException(String message, Throwable cause) {
        super(message, cause);
    }
}
