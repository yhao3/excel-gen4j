package com.github.yhao3.excelgen4j.processor;

import com.github.yhao3.excelgen4j.domain.CellRef;
import com.github.yhao3.excelgen4j.domain.DataRef;
import com.github.yhao3.excelgen4j.domain.ParameterRef;
import com.github.yhao3.excelgen4j.error.ExcelGen4jException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;

public class CsvToExcelProcessor {

    public static void processCellRef(CellRef cellRef, SXSSFSheet sheet, Integer minRowNumInDataList, Integer maxRowNumInDataList) {
        Object cellRefValue = cellRef.getValue();
        int rowNum = shiftDownRowNumIfNecessary(cellRef, minRowNumInDataList, maxRowNumInDataList);
        if (cellRefValue instanceof String) {
            String value = ((String) cellRefValue).replaceAll("_", "");
            createCellAndValue(cellRef.getColIndex(), value, sheet, rowNum);
            mergeCell(cellRef.getMergeColumnAmount(), sheet, rowNum, cellRef.getColIndex());
        } else {
            throw new ExcelGen4jException("The cell value is only allowed to be String type!");
        }
    }

    private static int shiftDownRowNumIfNecessary(CellRef cellRef, Integer minRowNumInDataList, Integer maxRowNumInDataList) {
        int rowNum = cellRef.getRowNum();
        if (rowNum >= minRowNumInDataList) {
            // 在 minRowNumInDataList 之後的 row 都要下移
            // if rowNum greater than or equal to minRowNumInDataList, then rowNum should be added by maxRowNumInDataList - minRowNumInDataList
            return rowNum + maxRowNumInDataList - minRowNumInDataList;
        } else {
            return rowNum;
        }
    }

    public static void processDataRef(List<?> dataList, Class<?> clazz, DataRef dataRef, SXSSFSheet sheet) {
        int rowNum = dataRef.getRowNum();
        // iterate the dataList provided by the user, and map the data to the cell
        for (int i = 0; i < dataList.size(); i++) {
            Object data = dataList.get(i);
            // 1. Get the value of the field by reflection
            for (Field declaredField : clazz.getDeclaredFields()) {
                declaredField.setAccessible(true);
                Object value;
                try {
                    value = declaredField.get(data);
                } catch (Exception e) {
                    throw new ExcelGen4jException("Maybe the data list and the class " + clazz.getName() + " are not matched!");
                }
                if (Objects.equals(declaredField.getName().toLowerCase(), dataRef.getKey().toLowerCase())) {
                    processDataRefVal(value, sheet, rowNum + i, dataRef.getColIndex(), dataRef.getMergeColumnAmount());
                }
            }
        }
    }

    public static void processDataRefVal(Object value, SXSSFSheet sheet, int rowNum, int colIndex, int mergeCount) {
        value = processNullValue(value);
        if (value instanceof String) {
            createCellAndValue(colIndex, (String) value, sheet, rowNum);
            mergeCell(mergeCount, sheet, rowNum, colIndex);
        } else if (value instanceof Double) {
            createCellAndValue(colIndex, (Double) value, sheet, rowNum);
            mergeCell(mergeCount, sheet, rowNum, colIndex);
        } else if (value instanceof BigDecimal) {
            createCellAndValue(colIndex, ((BigDecimal) value).doubleValue(), sheet, rowNum);
            mergeCell(mergeCount, sheet, rowNum, colIndex);
        } else if (value instanceof Integer) {
            createCellAndValue(colIndex, (Integer) value, sheet, rowNum);
            mergeCell(mergeCount, sheet, rowNum, colIndex);
        } else {
            throw new IllegalArgumentException("Unsupported data type: " + value.getClass());
        }
    }

    private static Object processNullValue(Object value) {
        return value == null ? "" : value;
    }

    public static void processParameter(
        Map<String, Object> parametersMap,
        ParameterRef parameterRef,
        SXSSFSheet sheet,
        Integer minRowNumInDataList,
        Integer maxRowNumInDataList
    ) {
        int rowNum = shiftDownRowNumIfNecessary(parameterRef, minRowNumInDataList, maxRowNumInDataList);

        Object obj = parametersMap.get(parameterRef.getKey());
        if (obj == null) {
            throw new ExcelGen4jException("The parameter \"" + parameterRef.getKey() + "\" is not found!");
        }
        processDataRefVal(obj, sheet, rowNum, parameterRef.getColIndex(), parameterRef.getMergeColumnAmount());
    }

    private static void mergeCell(int mergeCount, SXSSFSheet sheet, int rowNum, int colIndex) {
        if (mergeCount > 0) {
            sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, colIndex, colIndex + mergeCount - 1));
        }
    }

    private static <T> void createCellAndValue(int colIndex, T value, SXSSFSheet sheet, int rowNum) {
        SXSSFRow row = getSingletonRow(sheet, rowNum);
        if (value instanceof Integer) {
            row.createCell(colIndex).setCellValue((Integer) value);
        } else if (value instanceof Double) {
            row.createCell(colIndex).setCellValue((Double) value);
        } else if (value instanceof String) {
            row.createCell(colIndex).setCellValue((String) value);
        }
    }

    private static SXSSFRow getSingletonRow(SXSSFSheet sheet, int rowNum) {
        SXSSFRow row = sheet.getRow(rowNum);
        if (row == null) {
            return sheet.createRow(rowNum);
        } else {
            return sheet.getRow(rowNum);
        }
    }
}
