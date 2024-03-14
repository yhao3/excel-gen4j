package com.github.yhao3.excelpro.excel;

import com.github.yhao3.excelpro.csv.CsvReader;
import com.github.yhao3.excelpro.domain.CellRef;
import com.github.yhao3.excelpro.domain.CsvRef;
import com.github.yhao3.excelpro.domain.DataRef;
import com.github.yhao3.excelpro.domain.ParameterRef;
import com.github.yhao3.excelpro.processor.CsvToExcelProcessor;
import com.opencsv.exceptions.CsvException;
import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class ExcelGenerator {

    public SXSSFWorkbook generate(String path, Map<String, Object> parametersMap, List<?> dataList, Class<?> clazz)
        throws IOException, CsvException {
        // 1. Read the CSV template as CsvRef object
        BufferedInputStream inputStream = new BufferedInputStream(new FileInputStream(path));
        CsvReader csvReader = new CsvReader();
        CsvRef csvRef = csvReader.read(inputStream);

        Set<DataRef> dataListKeySet = csvRef.getDataListKeySet();
        Integer minRowNumInDataList = dataListKeySet.stream().map(DataRef::getRowNum).min(Integer::compareTo).orElse(0);
        Integer maxRowNumInDataList = minRowNumInDataList + dataList.size() - 1;

        // 2. Generate the Excel file
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        SXSSFSheet sheet = workbook.createSheet();
        // 2.1 Process the parameters
        for (ParameterRef parameterRef : csvRef.getParametersKeySet()) {
            CsvToExcelProcessor.processParameter(parametersMap, parameterRef, sheet, minRowNumInDataList, maxRowNumInDataList);
        }
        // 2.2 Process the cellRefs
        for (CellRef cellRef : csvRef.getCellRefSet()) {
            CsvToExcelProcessor.processCellRef(cellRef, sheet, minRowNumInDataList, maxRowNumInDataList);
        }
        // 2.3 Process the dataList
        for (DataRef dataRef : dataListKeySet) {
            CsvToExcelProcessor.processDataRef(dataList, clazz, dataRef, sheet);
        }
        return workbook;
    }
}
