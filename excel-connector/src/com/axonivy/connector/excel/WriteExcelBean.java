package com.axonivy.connector.excel;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import ch.ivyteam.ivy.environment.Ivy;
import ch.ivyteam.ivy.scripting.objects.Recordset;

public class WriteExcelBean {

  public void perform(Recordset rs, String filePath) throws Exception {

    try (OutputStream out = new FileOutputStream(filePath)) {
      @SuppressWarnings("resource")
      HSSFWorkbook wb = new HSSFWorkbook();
      HSSFSheet s = wb.createSheet();
      HSSFRow r = null;
      HSSFCell c = null;
      HSSFCellStyle cs = wb.createCellStyle();
      r = s.createRow(0);
      List<String> colnames = rs.getKeys();
      for (int i = 0; i < colnames.size(); i++) {
        c = r.createCell(i);
        c.setCellStyle(cs);
        c.setCellValue(colnames.get(i));
      }
      for (int j = 0; j < rs.size(); j++) {
        int rownum = (j + 1);
        r = s.createRow(rownum);

        for (int i = 0; i < colnames.size(); i++) {
          int cellnum = (i);

          c = r.createCell(cellnum);
          c.setCellStyle(cs);
          Object obj = rs.getField(j, colnames.get(i));
          if (obj != null) {
            try {
              if (!Double.valueOf(obj.toString()).isNaN()) {
                c.setCellValue(Double.valueOf(obj.toString())
                        .doubleValue());
              }
            } catch (Exception ex) {
              c.setCellValue(obj.toString());
            }
          }
        }
      }

      wb.write(out);
      // wb.close();

    } catch (Exception ex) {
      Ivy.log().error(
              "WriteExcelBean failed to write output file!  Output file path =" + filePath + ".");
      ex.printStackTrace();
      throw ex;
    }
  }
}
