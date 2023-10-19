package com.axonivy.util.excel.importer;

import java.nio.file.Path;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
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
    String dataName = StringUtils.substringBeforeLast(filePath.getFileName().toString(), ".");
    return toEntity(wb.getSheetAt(0), dataName);
  }

  private IEntityClass toEntity(Sheet sheet, String dataName) {
    String fqName = manager.getDefaultNamespace()+"."+dataName;
    if (manager.findDataClass(fqName) != null) {
      throw new RuntimeException("entity "+fqName+" already exists");
    }
    var entity = manager.createEntityClass(fqName, null);

    ExcelReader.parseColumns(sheet).stream().forEachOrdered(col -> {
      entity.addField(fieldName(col.name()), col.type().getName());
    });
    return entity;
  }

  private String fieldName(String colName) {
    if (StringUtils.isAllUpperCase(colName)) {
      return colName.toLowerCase();
    }
    return StringUtils.uncapitalize(colName);
  }

}
