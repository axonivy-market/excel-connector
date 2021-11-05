package com.axonivy.connector.excel;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import ch.ivyteam.ivy.environment.Ivy;
import ch.ivyteam.ivy.process.engine.IRequestId;
import ch.ivyteam.ivy.scripting.exceptions.IvyScriptException;
import ch.ivyteam.ivy.scripting.language.IIvyScriptContext;
import ch.ivyteam.ivy.scripting.objects.CompositeObject;
import ch.ivyteam.ivy.scripting.objects.Record;
import ch.ivyteam.ivy.scripting.objects.Recordset;

public class ReadExcelBean {

  private boolean addTitleRow = false;

  public CompositeObject perform(IRequestId requestId, CompositeObject in,
          IIvyScriptContext context, String rsParam, String filepathParam) throws Exception {
    Recordset rs = null; // (Recordset) getProcessDataField(context, rsParam);
    String filePath = getFilePath(context, filepathParam);

    try (InputStream input = new FileInputStream(filePath)) {
      POIFSFileSystem fs = new POIFSFileSystem(input);
      @SuppressWarnings("resource")
      HSSFWorkbook wb = new HSSFWorkbook(fs);
      HSSFSheet sheet = wb.getSheetAt(0);
      Iterator<Row> rows = sheet.rowIterator();
      int columns = countColumns(sheet.rowIterator());

      String[] columnTitles = createColumnTitles(columns,
              sheet.rowIterator());

      rs = new Recordset(columnTitles);

      if (!addTitleRow) {
        rows.next();
      } else {
        Ivy.log().warn("ReadExcelBean: No title row found.");
      }

      while (rows.hasNext()) {
        HSSFRow row = (HSSFRow) rows.next();

        Iterator<Cell> cells = row.cellIterator();
        Object[] cellValues = new Object[columns];
        while (cells.hasNext()) {
          HSSFCell cell = (HSSFCell) cells.next();
          if (cell.getColumnIndex() < columns) {
            switch (cell.getCellType()) {/*
                                          * case HSSFCell.CELL_TYPE_NUMERIC:
                                          * cellValues[cell.getColumnIndex()] =
                                          * new Double(
                                          * cell.getNumericCellValue()); break;
                                          * case HSSFCell.CELL_TYPE_STRING:
                                          * cellValues[cell.getColumnIndex()] =
                                          * cell .getStringCellValue(); break;
                                          */
              default:
                cellValues[cell.getColumnIndex()] = null;
                break;
            }
          }
        }
        rs.add(new Record(columnTitles, cellValues));
      }

      // wb.close();

    } catch (Exception ex) {
      Ivy.log().error("ReadExcelBean failed! Source file =" + filePath
              + ". Target recordset =" + rsParam + ".");
      ex.printStackTrace();
      throw ex;
    }

    // Store result in process data
    // setProcessDataField(context, rsParam, rs);
    return in;
  }

  private String getFilePath(IIvyScriptContext context, String filepathParam)
          throws IvyScriptException {
    String filePath;
    // Evaluate call parameter
    if (filepathParam.startsWith("in.")) {
      filePath = "";// (String) getProcessDataField(context, filepathParam);
    } else {
      filePath = filepathParam;
    }
    return filePath;
  }

  private String[] createColumnTitles(int numberOfColumns, Iterator<Row> rows) {
    String[] columnTitleSet = new String[numberOfColumns];
    int i = 0;
    HSSFRow row = (HSSFRow) rows.next();
    Iterator<Cell> cells = row.cellIterator();
    addTitleRow = false;
    if (checkCellEntries(row.cellIterator())) {
      while (cells.hasNext()) {
        HSSFCell currentCell = (HSSFCell) cells.next();
        if (currentCell.getColumnIndex() < numberOfColumns) {
          columnTitleSet[currentCell.getColumnIndex()] = currentCell
                  .getStringCellValue();
          i++;
          addTitleRow = false;
        }
      }
    } else {
      while (cells.hasNext() && i < numberOfColumns) {
        cells.next();
        columnTitleSet[i] = "Column" + (i + 1);
        i++;
      }
      addTitleRow = true;
    }
    return columnTitleSet;
  }

  private boolean checkCellEntries(Iterator<Cell> cells) {
    /*
     * boolean type = false; while (cells.hasNext()) { HSSFCell cell =
     * (HSSFCell) cells.next(); switch (cell.getCellType()) { case
     * HSSFCell.CELL_TYPE_NUMERIC: type = false; break;
     *
     * case HSSFCell.CELL_TYPE_STRING: if (cell.getStringCellValue().length() <
     * 100) { type = true; } else { type = false; } break;
     *
     * default: type = false; break; } if (!type) { return false; } }
     */
    return true;
  }

  private int countColumns(Iterator<Row> rows) {
    HSSFRow row = (HSSFRow) rows.next();
    Iterator<Cell> cells = row.cellIterator();
    int i = 0;
    while (cells.hasNext()) {
      i++;
      cells.next();
    }
    return i;
  }
}
