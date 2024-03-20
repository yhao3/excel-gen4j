package com.github.yhao3.excelgen4j;

import static org.junit.jupiter.api.Assertions.assertThrows;

import com.github.yhao3.excelgen4j.error.InvalidKeyException;
import com.github.yhao3.excelgen4j.excel.ExcelGenerator;
import com.opencsv.exceptions.CsvException;
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Test;

public class IntegrationTests {

    /**
     * Simple Template__________
     * {{bodyHeader1}},{{bodyHeader2_______}},,,,,,,{{bodyHeader3__}}
     * [[bodyValue1]],[[bodyValue2_______]],,,,,,,[[bodyValue3__]]
     */
    @Test
    void givenSimpleTemplate_whenGenerate_thenSuccess() throws IOException, CsvException {
        ExcelGenerator generator = new ExcelGenerator();
        Map<String, Object> parametersMap = Map.of("bodyHeader1", "Product ID", "bodyHeader2", "Product Name", "bodyHeader3", "Quantity");
        List<SimpleTemplateData> dataList = List.of(
            new SimpleTemplateData("1", "Product 1", 10),
            new SimpleTemplateData("2", "Product 2", 20)
        );
        try (
            SXSSFWorkbook workbook = generator.generate(
                "src/test/resources/simple-template.csv",
                parametersMap,
                dataList,
                SimpleTemplateData.class
            )
        ) {
            FileOutputStream fileOutputStream = new FileOutputStream("src/test/resources/simple-template.xlsx");
            workbook.write(fileOutputStream);
        }
    }

    static class SimpleTemplateData {

        private String bodyValue1;
        private String bodyValue2;
        private int bodyValue3;

        public SimpleTemplateData(String bodyValue1, String bodyValue2, int bodyValue3) {
            this.bodyValue1 = bodyValue1;
            this.bodyValue2 = bodyValue2;
            this.bodyValue3 = bodyValue3;
        }

        public String getBodyValue1() {
            return bodyValue1;
        }

        public void setBodyValue1(String bodyValue1) {
            this.bodyValue1 = bodyValue1;
        }

        public String getBodyValue2() {
            return bodyValue2;
        }

        public void setBodyValue2(String bodyValue2) {
            this.bodyValue2 = bodyValue2;
        }

        public int getBodyValue3() {
            return bodyValue3;
        }

        public void setBodyValue3(int bodyValue3) {
            this.bodyValue3 = bodyValue3;
        }
    }

    /**
     * Explain the placeholder syntax:
     * ,    : Cell separator
     * {{}} : Parameters
     * [[]] : Data List
     * _    : Merge ranges
     * Input CSV Template:
     * <pre>
     * <code>
     * ```csv
     * ,,,{{compName}},,,,,,,,,,,,,,
     * ,,{{reportName}},,,,,,,,,,,,,,,
     * {{header1}},,,,{{header2}},,,{{header3}},,,,{{header4}},,,{{header5}},,,{{header6}}
     * {{value1}},,,,{{value2}},,,{{value3}},,,,{{value4}},,,{{value5}},,,{{value6}}
     * {{bodyHeader1}},{{bodyHeader2}},,,,,,,{{bodyHeader3}},,{{bodyHeader4}},,{{bodyHeader5}},,{{bodyHeader6}},,,{{bodyHeader7}}
     * [[bodyValue1]],[[bodyValue2]],,,,,,,[[bodyValue3]],,[[bodyValue4]],,[[bodyValue5]],,[[bodyValue6]],,,[[bodyValue7]]
     * ```
     * </code>
     * <code>
     * ```csv
     * ,,,{{compName}}______________
     * ,,{{reportName}}_______________
     * {{header1}}____,{{header2}}___,{{header3}}____,{{header4}}___,{{header5}}___,{{header6}}
     * {{value1}}____,{{value2}}___,{{value3}}____,{{value4}}___,{{value5}}___,{{value6}}
     * {{bodyHeader1}},{{bodyHeader2}}_______,{{bodyHeader3}}__,{{bodyHeader4}}__,{{bodyHeader5}}__,{{bodyHeader6}}___,{{bodyHeader7}}
     * [[bodyValue1]],[[bodyValue2]]_______,[[bodyValue3]]__,[[bodyValue4]]__,[[bodyValue5]]__,[[bodyValue6]]___,[[bodyValue7]]
     * ```
     * </code>
     * </pre>
     * <p>
     * <p>
     * Input Values:
     *
     * <pre>
     *     compName: "Demo Inc."
     *     reportName: "Sales Report"
     *     header1: "Order ID"
     *     header2: "Total Price"
     *     header3: "Tax"
     *     header4: "Total Price with Tax"
     *     header5: "Date"
     *     header6: "Memo"
     *     value1: "1"
     *     value2: 100
     *     value3: 10
     *     value4: 110
     *     value5: "2021-01-01"
     *     value6: "Sold to customer 1"
     *     bodyHeader1: "Product ID"
     *     bodyHeader2: "Product Name"
     *     bodyHeader3: "Quantity"
     *     bodyHeader4: "Unit Price"
     *     bodyHeader5: "Total Price"
     *     bodyHeader6: "Tax"
     *     bodyHeader7: "Total Price with Tax"
     *
     *     bodyValueList: [
     *         {
     *             bodyValue1: "1"
     *             bodyValue2: "Product 1"
     *             bodyValue3: 10
     *             bodyValue4: 10
     *             bodyValue5: 100
     *             bodyValue6: 10
     *             bodyValue7: 110
     *         },
     *         {
     *             bodyValue1: "2"
     *             bodyValue2: "Product 2"
     *             bodyValue3: 20
     *             bodyValue4: 20
     *             bodyValue5: 400
     *             bodyValue6: 40
     *             bodyValue7: 440
     *         }
     *     ]
     * </pre>
     * <p>
     * Expected Output:
     */
    @Test
    void test() throws IOException, CsvException {
        ExcelGenerator generator = new ExcelGenerator();
        Map<String, Object> parametersMap = prepareMockParams();
        List<ExampleData> dataList = prepareMockDataList();
        try (
            SXSSFWorkbook workbook = generator.generate("src/test/resources/input-template.csv", parametersMap, dataList, ExampleData.class)
        ) {
            FileOutputStream fileOutputStream = new FileOutputStream("src/test/resources/test.xlsx");
            workbook.write(fileOutputStream);
        }
    }

    @Test
    void test2() throws IOException, CsvException {
        ExcelGenerator generator = new ExcelGenerator();
        Map<String, Object> parametersMap = prepareMockParams();
        parametersMap.put("compName", "測試公司");
        parametersMap.put("reportName", "領料單");
        parametersMap.put("param1", "測試公司");
        parametersMap.put("param2", "1234567890");
        parametersMap.put("param3", "測試公司");
        parametersMap.put("param4", "2021-01-01");
        parametersMap.put("param5", "");
        parametersMap.put("param6", "");
        parametersMap.put("param7", "test1");
        parametersMap.put("param8", "test2");

        List<Master> dataList = prepareMockMasterData();
        try (SXSSFWorkbook workbook = generator.generate("src/test/resources/input-template2.csv", parametersMap, dataList, Master.class)) {
            FileOutputStream fileOutputStream = new FileOutputStream("src/test/resources/test2.xlsx");
            workbook.write(fileOutputStream);
        }
    }

    @Test
    void test3() throws IOException, CsvException {
        ExcelGenerator generator = new ExcelGenerator();
        Map<String, Object> parametersMap = prepareMockParams();
        parametersMap.put("compName", "測試公司");
        parametersMap.put("reportName", "領料單");
        parametersMap.put("param1", "測試公司");
        parametersMap.put("param2", "1234567890");
        parametersMap.put("param3", "測試公司");
        parametersMap.put("param4", "2021-01-01");
        parametersMap.put("param5", "");
        parametersMap.put("param6", "");
        parametersMap.put("param7", "test1");
        parametersMap.put("param8", "test2");

        List<Master> dataList = new ArrayList<>();
        Master master1 = new Master();
        master1.setString01("001");
        master1.setString02("台北市");
        master1.setDecimal01(new BigDecimal(23));
        master1.setDecimal02(new BigDecimal(100));
        master1.setString03("噸");
        master1.setDecimal03(new BigDecimal(2300));
        master1.setString04("去化作業");
        Master master2 = new Master();
        master2.setString01("002");
        master2.setString02("嘉義市");
        master2.setDecimal01(null);
        master2.setDecimal02(new BigDecimal(100));
        master2.setString03("噸");
        master2.setDecimal03(new BigDecimal(1300));
        master2.setString04("去化作業");
        Master master3 = new Master();
        master3.setString01("003");
        master3.setString02("台北市");
        master3.setDecimal01(new BigDecimal(23));
        master3.setDecimal02(new BigDecimal(100));
        master3.setString03("噸");
        master3.setDecimal03(new BigDecimal(2300));
        master3.setString04("去化作業");
        dataList.add(master1);
        dataList.add(master2);
        dataList.add(master3);
        try (SXSSFWorkbook workbook = generator.generate("src/test/resources/input-template2.csv", parametersMap, dataList, Master.class)) {
            FileOutputStream fileOutputStream = new FileOutputStream("src/test/resources/test2.xlsx");
            workbook.write(fileOutputStream);
        }
    }

    @Test
    void givenTemplateWithBlankKey_whenGenerate_thenThrowException() throws IOException, CsvException {
        ExcelGenerator generator = new ExcelGenerator();

        // expect to throw an exception
        assertThrows(InvalidKeyException.class, () -> {
            try (
                SXSSFWorkbook workbook = generator.generate(
                    "src/test/resources/invalid-template-with-blank-key.csv",
                    new HashMap<>(),
                    new ArrayList<>(),
                    Object.class
                )
            ) {
                OutputStream outputStream = new BufferedOutputStream(
                    new FileOutputStream("src/test/resources/invalid-template-with-blank-key.xlsx")
                );
                workbook.write(outputStream);
            }
        });
    }

    private static List<Master> prepareMockMasterData() {
        List<Master> dataList = new ArrayList<>();
        Master master1 = new Master();
        master1.setString01("001");
        master1.setString02("台北市");
        master1.setDecimal01(new BigDecimal(23));
        master1.setDecimal02(new BigDecimal(100));
        master1.setString03("噸");
        master1.setDecimal03(new BigDecimal(2300));
        master1.setString04("去化作業");
        Master master2 = new Master();
        master2.setString01("002");
        master2.setString02("嘉義市");
        master2.setDecimal01(new BigDecimal(13));
        master2.setDecimal02(new BigDecimal(100));
        master2.setString03("噸");
        master2.setDecimal03(new BigDecimal(1300));
        master2.setString04("去化作業");
        dataList.add(master1);
        dataList.add(master2);
        return dataList;
    }

    private static List<ExampleData> prepareMockDataList() {
        List<ExampleData> dataList = new ArrayList<>();
        ExampleData exampleData1 = new ExampleData();
        exampleData1.setBodyValue1("1");
        exampleData1.setBodyValue2("Product 1");
        exampleData1.setBodyValue3(10);
        exampleData1.setBodyValue4(10);
        exampleData1.setBodyValue5(100);
        exampleData1.setBodyValue6(10);
        exampleData1.setBodyValue7(110);
        ExampleData exampleData2 = new ExampleData();
        exampleData2.setBodyValue1("2");
        exampleData2.setBodyValue2("Product 2");
        exampleData2.setBodyValue3(20);
        exampleData2.setBodyValue4(20);
        exampleData2.setBodyValue5(400);
        exampleData2.setBodyValue6(40);
        exampleData2.setBodyValue7(440);
        dataList.add(exampleData1);
        dataList.add(exampleData2);
        return dataList;
    }

    private static Map<String, Object> prepareMockParams() {
        Map<String, Object> parametersMap = new HashMap<>();
        parametersMap.put("compName", "Demo Inc.");
        parametersMap.put("reportName", "Sales Report");
        parametersMap.put("header1", "Order ID");
        parametersMap.put("header2", "Total Price");
        parametersMap.put("header3", "Tax");
        parametersMap.put("header4", "Total Price with Tax");
        parametersMap.put("header5", "Date");
        parametersMap.put("header6", "Memo");
        parametersMap.put("value1", "1");
        parametersMap.put("value2", 100);
        parametersMap.put("value3", 10);
        parametersMap.put("value4", 110);
        parametersMap.put("value5", "2021-01-01");
        parametersMap.put("value6", "Sold to customer 1");
        parametersMap.put("bodyHeader1", "Product ID");
        parametersMap.put("bodyHeader2", "Product Name");
        parametersMap.put("bodyHeader3", "Quantity");
        parametersMap.put("bodyHeader4", "Unit Price");
        parametersMap.put("bodyHeader5", "Total Price");
        parametersMap.put("bodyHeader6", "Tax");
        parametersMap.put("bodyHeader7", "Total Price with Tax");
        return parametersMap;
    }
}
