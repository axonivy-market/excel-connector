package com.axonivy.util.excel.importer;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

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
    Iterator<Row> rowIterator = loadRowIterator(filePath);
    List<String> headerCells = getHeaderCells(rowIterator);
    String dataName = StringUtils.substringBeforeLast(filePath.getFileName().toString(), ".");
    String fqName = manager.getDefaultNamespace()+"."+dataName;
    if (manager.findDataClass(fqName) != null) {
      throw new RuntimeException("entity "+fqName+" already exists");
    }
    var entity = manager.createEntityClass(fqName, null);
    addContentCellsToRecordset(entity, headerCells, rowIterator);
    return entity;
  }

  private static Iterator<Row> loadRowIterator(Path filePath) {
    return ExcelLoader.load(filePath).getSheetAt(0).rowIterator();
  }

  private static List<String> getHeaderCells(Iterator<Row> rowIterator) {
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

  private static void addContentCellsToRecordset(IEntityClass entity, List<String> names, Iterator<Row> rowIterator) {
    if (!rowIterator.hasNext()) {
      return;
    }
    Row row = rowIterator.next();
    Iterator<Cell> cellIterator = row.cellIterator();
    int cellNo = 0;
    while (cellIterator.hasNext()) {
      Cell cell = cellIterator.next();
      var name = StringUtils.uncapitalize(names.get(cellNo));
      switch (cell.getCellType()) {
        case NUMERIC:
          entity.addField(name, Float.class.getName());
          break;
        case STRING:
          entity.addField(name, String.class.getName());
          break;
        case BOOLEAN:
          entity.addField(name, Boolean.class.getName());
          break;
        default:
          entity.addField(name, String.class.getName());
          break;
      }
      cellNo++;
    }
  }
}
