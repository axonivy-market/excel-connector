package com.axonivy.util.excel.importer;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelReader {

  public static List<Column> parseColumns(Sheet sheet) {
    Iterator<Row> rows = sheet.rowIterator();
    List<String> headers = getHeaderCellNames(rows);
    return createEntityFields(headers, rows);
  }

  private static List<String> getHeaderCellNames(Iterator<Row> rowIterator) {
    List<String> headerCells = new ArrayList<String>();
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

  private static List<Column> createEntityFields(List<String> names, Iterator<Row> rowIterator) {
    List<Column> columns = new ArrayList<>();
    if (!rowIterator.hasNext()) {
      return List.of();
    }
    Row row = rowIterator.next();
    Iterator<Cell> cellIterator = row.cellIterator();
    int cellNo = 0;
    while (cellIterator.hasNext()) {
      Cell cell = cellIterator.next();
      var name = names.get(cellNo);
      var column = toColumn(name, cell.getCellType());
      columns.add(column);
      cellNo++;
    }
    return columns;
  }

  private static Column toColumn(String name, CellType type) {
    switch (type) {
      case NUMERIC:
        return new Column(name, Float.class);
      case STRING:
        return new Column(name, String.class);
      case BOOLEAN:
        return new Column(name, Boolean.class);
      default:
        return new Column(name, String.class);
    }
  }

  public record Column(String name, Class<?> type) {
}

}
