package com.axonivy.util.excel.importer;

import java.nio.file.Path;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import ch.ivyteam.ivy.environment.Ivy;
import ch.ivyteam.ivy.scripting.dataclass.DataClassFieldModifier;
import ch.ivyteam.ivy.scripting.dataclass.IDataClassManager;
import ch.ivyteam.ivy.scripting.dataclass.IEntityClass;
import ch.ivyteam.ivy.scripting.dataclass.IEntityClassField;
import ch.ivyteam.ivy.scripting.dataclass.IProjectDataClassManager;

public class EntityClassReader {

  public final IProjectDataClassManager manager;

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
    dataName = StringUtils.capitalize(dataName);
    return toEntity(wb.getSheetAt(0), dataName);
  }

  public IEntityClass toEntity(Sheet sheet, String dataName) {
    String fqName = manager.getDefaultNamespace()+"."+dataName;
    if (manager.findDataClass(fqName) != null) {
      throw new RuntimeException("entity "+fqName+" already exists");
    }
    var entity = manager.createEntityClass(fqName, null);

    withIdField(entity);
    ExcelReader.parseColumns(sheet).stream().forEachOrdered(col -> {
      entity.addField(fieldName(col.name()), col.type().getName());
    });
    return entity;
  }

  private void withIdField(IEntityClass entity) {
    IEntityClassField id = entity.addField("id", Integer.class.getSimpleName());
    id.setDatabaseFieldName("id");
    id.addModifier(DataClassFieldModifier.PERSISTENT);
    id.addModifier(DataClassFieldModifier.ID);
    id.addModifier(DataClassFieldModifier.GENERATED);
    id.setComment("Identifier");
  }

  private String fieldName(String colName) {
    if (StringUtils.isAllUpperCase(colName)) {
      return colName.toLowerCase();
    }
    return StringUtils.uncapitalize(colName);
  }

}
