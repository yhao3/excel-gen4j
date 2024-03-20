package com.github.yhao3.excelgen4j.csv;

import com.github.yhao3.excelgen4j.domain.CellRef;
import com.github.yhao3.excelgen4j.domain.CsvRef;
import com.github.yhao3.excelgen4j.domain.DataRef;
import com.github.yhao3.excelgen4j.domain.ParameterRef;
import com.github.yhao3.excelgen4j.domain.syntax.DataPropertySyntax;
import com.github.yhao3.excelgen4j.domain.syntax.ParameterSyntax;
import com.github.yhao3.excelgen4j.error.InvalidKeyException;
import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.List;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.commons.lang3.StringUtils;

public class CsvReader {

    public CsvRef read(InputStream inputStream) throws IOException, CsvException {
        // Issue: 當使用 java.io.InputStreamReader 來讀取 CSV 要特別處理 UTF-8 with BOM 編碼，否則 CSV 第一個欄位 always 會是 null
        // see: https://stackoverflow.com/a/44281660
        // Solution: 使用 org.apache.commons.io.input.BOMInputStream 包裝 java.io.InputStream 來處理 UTF-8 with BOM 編碼
        return read(BOMInputStream.builder().setInputStream(inputStream).get());
    }

    public CsvRef read(BOMInputStream inputStream) throws IOException, CsvException {
        InputStreamReader reader = new InputStreamReader(inputStream);
        return buildCsvRef(reader);
    }

    public static CsvRef buildCsvRef(InputStreamReader reader) throws IOException, CsvException {
        CSVReader csvReader = new CSVReader(reader);
        List<String[]> strArrList = csvReader.readAll();
        // parse the csv file to CsvDTO object
        CsvRef csvRef = new CsvRef();
        for (int row = 0; row < strArrList.size(); row++) {
            String[] strArr = strArrList.get(row);
            for (int col = 0; col < strArr.length; col++) {
                String str = strArr[col];
                String[] splitArr = str.split(",");
                for (String cell : splitArr) {
                    if (cell.startsWith(ParameterSyntax.getPrefix()) && cell.endsWith(ParameterSyntax.getSuffix())) {
                        String key = extractKey(cell);
                        int underscoresCount = countCollapsibleSyntaxUnderscoresInTheEnd(cell);
                        ParameterRef paramRef = new ParameterRef();
                        paramRef.setRowNum(row);
                        paramRef.setColIndex(col);
                        paramRef.setKey(key);
                        paramRef.setMergeColumnAmount(underscoresCount);
                        csvRef.addParameterKey(paramRef);
                    } else if (cell.startsWith(DataPropertySyntax.getPrefix()) && cell.endsWith(DataPropertySyntax.getSuffix())) {
                        String key = extractKey(cell);
                        int underscoresCount = countCollapsibleSyntaxUnderscoresInTheEnd(cell);
                        DataRef dataRef = new DataRef();
                        dataRef.setRowNum(row);
                        dataRef.setColIndex(col);
                        dataRef.setKey(key);
                        dataRef.setMergeColumnAmount(underscoresCount);
                        csvRef.addDataListKey(dataRef);
                    } else {
                        // else is the plain text cell
                        if (StringUtils.isBlank(cell)) {
                            continue;
                        }
                        int underscoresCount = countUnderscoresInTheEnd(cell);
                        CellRef cellRef = new CellRef();
                        cellRef.setRowNum(row);
                        cellRef.setColIndex(col);
                        cellRef.setValue(cell);
                        cellRef.setMergeColumnAmount(underscoresCount);
                        csvRef.addCellRef(cellRef);
                    }
                }
            }
        }
        return csvRef;
    }

    static String extractKey(String cell) {
        String keyWithUnderscoresIfExists = cell.substring(2, cell.length() - 2);
        if (StringUtils.isBlank(keyWithUnderscoresIfExists)) {
            throw new InvalidKeyException("The key in the cell is blank!");
        }
        // replace the underscores in the end of the key
        int underscoresCount = countUnderscoresInTheEnd(keyWithUnderscoresIfExists);
        if (underscoresCount > 0) {
            return keyWithUnderscoresIfExists.substring(0, keyWithUnderscoresIfExists.length() - underscoresCount);
        }
        return keyWithUnderscoresIfExists;
    }

    static int countCollapsibleSyntaxUnderscoresInTheEnd(String cell) {
        int count = 0;
        for (int i = cell.length() - 3; i >= 0; i--) {
            if (cell.charAt(i) == '_') {
                count++;
            } else {
                break;
            }
        }
        return count;
    }

    static int countUnderscoresInTheEnd(String cell) {
        int count = 0;
        for (int i = cell.length() - 1; i >= 0; i--) {
            if (cell.charAt(i) == '_') {
                count++;
            } else {
                break;
            }
        }
        return count;
    }
}
