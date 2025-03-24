package com.github.srteixeiradias.explode_excel_formatter.domain;

import java.util.List;

public class ExcelRow {

    private final List<String> columns;

    public ExcelRow(List<String> columns) {
        this.columns = columns;
    }

    public List<String> getColumns() {
        return columns;
    }
}
