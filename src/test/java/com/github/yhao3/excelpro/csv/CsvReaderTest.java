package com.github.yhao3.excelpro.csv;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import com.github.yhao3.excelpro.error.InvalidKeyException;
import org.junit.jupiter.api.Test;

class CsvReaderTest {

    @Test
    void countUnderscoresInTheEnd_returnsCorrectCount_whenStringEndsWithUnderscores() {
        String cell = "__test___";
        int result = CsvReader.countUnderscoresInTheEnd(cell);
        assertEquals(3, result);
    }

    @Test
    void countUnderscoresInTheEnd_returnsZero_whenStringDoesNotEndWithUnderscores() {
        String cell = "__t__est";
        int result = CsvReader.countUnderscoresInTheEnd(cell);
        assertEquals(0, result);
    }

    @Test
    void countUnderscoresInTheEnd_returnsZero_whenStringIsEmpty() {
        String cell = "";
        int result = CsvReader.countUnderscoresInTheEnd(cell);
        assertEquals(0, result);
    }

    @Test
    void countUnderscoresInTheEnd_returnsCorrectCount_whenStringIsOnlyUnderscores() {
        String cell = "____";
        int result = CsvReader.countUnderscoresInTheEnd(cell);
        assertEquals(4, result);
    }

    @Test
    void extractKey_returnsKeyWithoutUnderscores_whenKeyEndsWithUnderscoresAndKeyContainsUnderscores() {
        String cell = "{{k__ey___}}";
        String result = CsvReader.extractKey(cell);
        assertEquals("k__ey", result);
    }

    @Test
    void extractKey_returnsKeyWithoutUnderscores_whenKeyEndsWithUnderscores() {
        String cell = "{{key___}}";
        String result = CsvReader.extractKey(cell);
        assertEquals("key", result);
    }

    @Test
    void extractKey_returnsKeyAsIs_whenKeyDoesNotEndWithUnderscores() {
        String cell = "{{key}}";
        String result = CsvReader.extractKey(cell);
        assertEquals("key", result);
    }

    @Test
    void extractKey_throwsInvalidKeyException_whenKeyIsEmpty() {
        String cell = "{{}}";
        assertThrows(InvalidKeyException.class, () -> CsvReader.extractKey(cell));
    }

    @Test
    void extractKey_returnsKeyWithoutUnderscores_whenKeyIsOnlyUnderscores() {
        String cell = "{{____}}";
        String result = CsvReader.extractKey(cell);
        assertEquals("", result);
    }
}
