package com.github.yhao3.excelgen4j.domain;

import java.util.LinkedHashSet;
import java.util.Set;

public class CsvRef {

    private Set<ParameterRef> parametersKeySet = new LinkedHashSet<>();
    private Set<DataRef> dataListKeySet = new LinkedHashSet<>();
    private Set<CellRef> cellRefSet = new LinkedHashSet<>();

    public Set<ParameterRef> getParametersKeySet() {
        return parametersKeySet;
    }

    public void setParametersKeySet(Set<ParameterRef> parametersKeySet) {
        this.parametersKeySet = parametersKeySet;
    }

    public void addParameterKey(ParameterRef parameterRef) {
        this.parametersKeySet.add(parameterRef);
    }

    public Set<DataRef> getDataListKeySet() {
        return dataListKeySet;
    }

    public void setDataListKeySet(Set<DataRef> dataListKeySet) {
        this.dataListKeySet = dataListKeySet;
    }

    public void addDataListKey(DataRef dataRef) {
        this.dataListKeySet.add(dataRef);
    }

    public Set<CellRef> getCellRefSet() {
        return cellRefSet;
    }

    public void setCellRefSet(Set<CellRef> cellRefSet) {
        this.cellRefSet = cellRefSet;
    }

    public void addCellRef(CellRef cellRef) {
        this.cellRefSet.add(cellRef);
    }
}
