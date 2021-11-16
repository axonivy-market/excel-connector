package com.axonivy.connector.excel;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ch.ivyteam.ivy.scripting.objects.Record;
import ch.ivyteam.ivy.scripting.objects.Recordset;

public class ExcelReader {

  public Recordset getRecordset(String filePath) {
    Recordset recordset = null;
    Iterator<Row> rowIterator = null;

    if (filePath.endsWith(".xls")) {
      HSSFWorkbook workbook = openXls(filePath);
      HSSFSheet sheet = workbook.getSheetAt(0);
      rowIterator = sheet.rowIterator();
    } else {
      XSSFWorkbook workbook = openXlsx(filePath);
      XSSFSheet sheet = workbook.getSheetAt(0);
      rowIterator = sheet.rowIterator();
    }

    ArrayList<String> headerCells = getHeaderCells(rowIterator);
    recordset = new Recordset(headerCells);
    addContentCellsToRecordset(recordset, rowIterator);

    return recordset;
  }

  private static HSSFWorkbook openXls(String filePath) {
    try (InputStream input = new FileInputStream(filePath);
            POIFSFileSystem fs = new POIFSFileSystem(input);
            HSSFWorkbook workbook = new HSSFWorkbook(fs);) {

      return workbook;

    } catch (Exception ex) {
      throw new RuntimeException("Could not read excel file from " + filePath, ex);
    }

  }

  private static XSSFWorkbook openXlsx(String filePath) {
    try (FileInputStream input = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(input);) {

      return workbook;

    } catch (Exception ex) {
      throw new RuntimeException("Could not read excel file from " + filePath, ex);
    }

  }

  private static ArrayList<String> getHeaderCells(Iterator<Row> rowIterator) {
    ArrayList<String> headerCells = new ArrayList<String>();

    if (rowIterator.hasNext()) {
      Row row = rowIterator.next();
      Iterator<Cell> cellIterator = row.cellIterator();
      while (cellIterator.hasNext()) {
        Cell cell = cellIterator.next();
        headerCells.add(cell.getStringCellValue());
      }
    }

    return headerCells;
  }

  private static void addContentCellsToRecordset(Recordset recordset, Iterator<Row> rowIterator) {
    while (rowIterator.hasNext()) {
      ArrayList<Object> contentCells = new ArrayList<Object>();
      Row row = rowIterator.next();
      Iterator<Cell> cellIterator = row.cellIterator();

      while (cellIterator.hasNext()) {
        Cell cell = cellIterator.next();
        switch (cell.getCellType()) {
          case NUMERIC:
            contentCells.add(cell.getNumericCellValue());
            break;
          case STRING:
            contentCells.add(cell.getStringCellValue());
            break;
          default:
            contentCells.add(null);
            break;
        }
      }
      recordset.add(new Record(recordset.getKeys(), contentCells));
    }

  }
}
