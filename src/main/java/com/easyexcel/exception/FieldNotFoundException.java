package com.easyexcel.exception;

/**
 * @author wrh
 * @date 2021/8/25
 */
public class FieldNotFoundException extends RuntimeException {

    public FieldNotFoundException(String message) {
        super(message);
    }
}
