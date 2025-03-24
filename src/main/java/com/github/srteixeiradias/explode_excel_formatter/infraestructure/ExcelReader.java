package com.github.srteixeiradias.explode_excel_formatter.infraestructure;


import com.github.srteixeiradias.explode_excel_formatter.domain.ExcelRow;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;



public class ExcelReader {
    // essa classe vai ser responsável por ler a planilha e extrair os dados sem células mescladas.

    public List<ExcelRow> readExcel(File file) throws IOException {
        List<ExcelRow> rows = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Pega a primeira aba

            for (Row row : sheet) {
                List<String> cellValues = new ArrayList<>();
                for (Cell cell : row) {
                    cellValues.add(cell.toString()); // Pega o valor da célula
                }
                rows.add(new ExcelRow(cellValues));
            }
        }
        return rows;
    }
}
