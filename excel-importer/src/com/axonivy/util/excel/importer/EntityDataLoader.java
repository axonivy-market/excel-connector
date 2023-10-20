package com.axonivy.util.excel.importer;

import java.util.Iterator;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

import javax.persistence.EntityManager;
import javax.persistence.EntityTransaction;
import javax.persistence.Query;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import ch.ivyteam.ivy.java.IJavaConfigurationManager;
import ch.ivyteam.ivy.process.data.persistence.IIvyEntityManager;
import ch.ivyteam.ivy.project.IIvyProjectManager;
import ch.ivyteam.ivy.scripting.dataclass.IEntityClass;
import ch.ivyteam.ivy.scripting.dataclass.IEntityClassField;

@SuppressWarnings("restriction")
public class EntityDataLoader {

  private final IIvyEntityManager manager;

  public EntityDataLoader(IIvyEntityManager manager) {
    this.manager = manager;
  }

  public void load(Sheet sheet, IEntityClass entity) {
    createTable(entity); // enforce db-schema to exist!

    Iterator<Row> rows = sheet.rowIterator();
    rows.next(); // skip header

    List<? extends IEntityClassField> fields = entity.getFields();
    var query = buildInsertQuery(entity, fields);

    EntityManager em = manager.createEntityManager();
    EntityTransaction tx = em.getTransaction();
    tx.begin();
    try {
      AtomicInteger rCount = new AtomicInteger();
      rows.forEachRemaining(row -> {
        Query insert = em.createNativeQuery(query);
        insert.setParameter("id", rCount.incrementAndGet());
        Iterator<Cell> cells = row.cellIterator();
        int c = 0;
        c++; // consumed by 'id'
        while(cells.hasNext()) {
          Cell cell = cells.next();
          IEntityClassField field = fields.get(c);
          String column = field.getName();
          insert.setParameter(column, getValue(cell));
          c++;
        }
        insert.executeUpdate();
      });
    } finally {
      tx.commit();
      em.close();
    }
  }

  private String buildInsertQuery(IEntityClass entity, List<? extends IEntityClassField> fields) {
    String colNames = fields.stream().map(IEntityClassField::getName).collect(Collectors.joining(","));
    var query = new StringBuilder("INSERT INTO "+entity.getSimpleName()+" ("+colNames+")\nVALUES (");
    var params = fields.stream().map(IEntityClassField::getName).map(f -> ":"+f+"").collect(Collectors.joining(", "));
    query.append(params);
    query.append(")");
    return query.toString();
  }

  public Class<?> createTable(IEntityClass entity) {
    entity.buildJavaSource(List.of(), null);
    var java = IJavaConfigurationManager.instance().getJavaConfiguration(entity.getResource().getProject());
    var ivy = IIvyProjectManager.instance().getIvyProject(entity.getResource().getProject());
    Class<?> entityClass;
    try {
      ivy.build(null);
      entityClass = java.getClassLoader().loadClass(entity.getName());
    } catch (Exception ex) {
      throw new RuntimeException("Failed to load entity class "+entity, ex);
    }
    manager.findAll(entityClass); // creates the schema through 'hibernate.hbm2ddl.auto=create'
    return entityClass;
  }

  private Object getValue(Cell cell) {
    if (cell.getCellType() == CellType.NUMERIC)  {
      return cell.getNumericCellValue();
    }
    return cell.getStringCellValue();
  }

}
