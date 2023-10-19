package com.axonivy.util.excel.importer;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import ch.ivyteam.ivy.environment.Ivy;
import ch.ivyteam.ivy.scripting.dataclass.IDataClassManager;
import ch.ivyteam.ivy.scripting.dataclass.IEntityClass;
import ch.ivyteam.ivy.scripting.dataclass.IProjectDataClassManager;

public class EntityClassReader {

  private final IProjectDataClassManager manager;

  @SuppressWarnings("restriction")
  public EntityClassReader() {
    this(IDataClassManager.instance().getProjectDataModelFor(Ivy.wfCase().getProcessModelVersion()));
  }

  public EntityClassReader(IProjectDataClassManager manager) {
    this.manager = manager;
  }

  public IEntityClass getEntity(Path filePath) {
    Workbook wb = ExcelLoader.load(filePath);

    List<String> headers = getHeaderCellNames(wb.getSheetAt(0).rowIterator());
    String dataName = StringUtils.substringBeforeLast(filePath.getFileName().toString(), ".");
    String fqName = manager.getDefaultNamespace()+"."+dataName;
    if (manager.findDataClass(fqName) != null) {
      throw new RuntimeException("entity "+fqName+" already exists");
    }
    var entity = manager.createEntityClass(fqName, null);
    var columns =  createEntityFields(headers, wb.getSheetAt(0).rowIterator());
    columns.stream().forEachOrdered(col -> {
      entity.addField(col.name(), col.type().getName());
    });
    return entity;
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
      var name = StringUtils.uncapitalize(names.get(cellNo));
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

  private record Column(String name, Class<?> type) {

  }
}
