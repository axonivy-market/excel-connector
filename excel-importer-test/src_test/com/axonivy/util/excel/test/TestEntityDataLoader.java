package com.axonivy.util.excel.test;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.eclipse.core.runtime.CoreException;
import org.eclipse.core.runtime.NullProgressMonitor;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.axonivy.util.excel.importer.EntityClassReader;
import com.axonivy.util.excel.importer.EntityDataLoader;
import com.axonivy.util.excel.importer.ExcelLoader;

import ch.ivyteam.ivy.environment.Ivy;
import ch.ivyteam.ivy.environment.IvyTest;
import ch.ivyteam.ivy.process.data.persistence.IIvyEntityManager;
import ch.ivyteam.ivy.scripting.dataclass.IEntityClass;

@IvyTest
@SuppressWarnings("restriction")
public class TestEntityDataLoader {

  private EntityDataLoader loader;
  private EntityClassReader reader;
  private IIvyEntityManager unit;

  @Test
  void loadDataToEntity(@TempDir Path dir) throws IOException, CoreException {
    Path path = dir.resolve("customers.xlsx");
    TstRes.loadTo(path, "sample.xlsx");

    Workbook wb = ExcelLoader.load(path);
    Sheet customerSheet = wb.getSheetAt(0);

    IEntityClass customer = reader.toEntity(customerSheet, "customer");
    try {
      customer.save(new NullProgressMonitor());
      Class<?> entity = loader.createTable(customer);
      assertThat(unit.findAll(entity)).isEmpty();
      loader.load(customerSheet, customer);
      List<?> records = unit.findAll(entity);
      assertThat(records).hasSize(2);
    } finally {
      customer.getResource().delete(true, new NullProgressMonitor());
    }
  }

  @BeforeEach
  void setup() {
    this.unit = Ivy.persistence().get("testing");
    this.loader = new EntityDataLoader(unit);
    this.reader = new EntityClassReader();
  }

}
