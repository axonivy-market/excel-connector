package com.axonivy.connector.excel;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import ch.ivyteam.ivy.scripting.objects.Recordset;

public class ExcelWriter {

  public void execute(Recordset recordset, String filePath) {
    try (OutputStream out = new FileOutputStream(filePath);
            HSSFWorkbook workbook = new HSSFWorkbook()) {
      HSSFSheet sheet = workbook.createSheet();
      HSSFCellStyle cs = workbook.createCellStyle();

      if (recordset != null) {
        setHeaderCells(sheet, recordset, cs);
        setContentCells(sheet, recordset, cs);
      }

      workbook.write(out);
    } catch (Exception ex) {
      throw new RuntimeException("Could not write excel file to " + filePath, ex);
    }
  }

  private static void setHeaderCells(HSSFSheet sheet, Recordset recordset, HSSFCellStyle cellstyle) {
    HSSFRow row = sheet.createRow(0);
    List<String> colnames = recordset.getKeys();
    for (int i = 0; i < colnames.size(); i++) {
      HSSFCell cell = row.createCell(i);
      cell.setCellStyle(cellstyle);
      cell.setCellValue(colnames.get(i));
    }
  }

  private static void setContentCells(HSSFSheet sheet, Recordset recordset, HSSFCellStyle cellstyle) {
    for (int i = 0; i < recordset.size(); i++) {
      HSSFRow row = sheet.createRow(i + 1);

      for (int j = 0; j < recordset.columnCount(); j++) {
        HSSFCell cell = row.createCell(j);
        cell.setCellStyle(cellstyle);

        Object obj = recordset.getField(i, j);
        if (obj != null) {
          if (obj instanceof Number) {
            cell.setCellValue(Double.valueOf(obj.toString())
                    .doubleValue());
          } else {
            cell.setCellValue(obj.toString());
          }

        }
      }
    }
  }
}
