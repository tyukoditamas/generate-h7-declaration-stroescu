package org.example;

public class InvalidExcelFileException  extends RuntimeException {
    public InvalidExcelFileException(String msg) {
        super(msg);
    }
}
