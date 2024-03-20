package com.github.yhao3.excelgen4j.excel;

import static org.junit.jupiter.api.Assertions.*;

import com.github.yhao3.excelgen4j.csv.CsvReader;
import com.github.yhao3.excelgen4j.domain.CsvRef;
import com.opencsv.exceptions.CsvException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.junit.jupiter.api.Test;

class CsvReaderTest {

    @Test
    void read() throws IOException, CsvException {
        CsvReader csvReader = new CsvReader();
        File file = new File("src/test/resources/test.csv");
        FileInputStream fis = new FileInputStream(file);
        CsvRef read = csvReader.read(fis);
        assertNotNull(read);
    }
}
