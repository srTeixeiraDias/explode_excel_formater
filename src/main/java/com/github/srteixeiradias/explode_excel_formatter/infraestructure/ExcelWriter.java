package com.github.srteixeiradias.explode_excel_formatter.infraestructure;

import com.github.srteixeiradias.explode_excel_formatter.domain.ExcelRow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelWriter {
    //essa classe vai ser responsável por escrever os dados formatados em uma planilha nova

    public void writeExcel(File file, List<ExcelRow> rows) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(file)) {

            Sheet sheet = workbook.createSheet("Formatted");

            int rowIndex = 0;
            for (ExcelRow rowData : rows) {
                Row row = sheet.createRow(rowIndex++);
                int colIndex = 0;

                for (String value : rowData.getColumns()) {
                    row.createCell(colIndex++).setCellValue(value);
                }
            }

            workbook.write(fos);
        }
    }
}
