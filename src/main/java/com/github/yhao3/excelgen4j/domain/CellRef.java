package com.github.yhao3.excelgen4j.domain;

public class CellRef {

    private int rowNum;
    private int colIndex;
    private int mergeColumnAmount;
    private Object value;

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public int getColIndex() {
        return colIndex;
    }

    public void setColIndex(int colIndex) {
        this.colIndex = colIndex;
    }

    public int getMergeColumnAmount() {
        return mergeColumnAmount;
    }

    public void setMergeColumnAmount(int mergeColumnAmount) {
        this.mergeColumnAmount = mergeColumnAmount;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }
}
