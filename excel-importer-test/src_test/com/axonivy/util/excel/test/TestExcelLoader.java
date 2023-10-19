package com.axonivy.util.excel.test;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.axonivy.util.excel.importer.ExcelLoader;

import ch.ivyteam.ivy.environment.IvyTest;

@IvyTest
public class TestExcelLoader {

  @Test
  void workbook_xlsx(@TempDir Path dir) throws IOException {
    Path path = dir.resolve("customers.xlsx");
    TstRes.loadTo(path, "sample.xlsx");
    Workbook wb = ExcelLoader.load(path);
    assertThat(wb).isNotNull();
  }

  @Test
  void workbook_legacyXls(@TempDir Path dir) throws IOException {
    Path path = dir.resolve("customers.xls");
    TstRes.loadTo(path, "sample_legacy.xls");
    Workbook wb = ExcelLoader.load(path);
    assertThat(wb).isNotNull();
  }

}
