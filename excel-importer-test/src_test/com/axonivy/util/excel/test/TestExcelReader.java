package com.axonivy.util.excel.test;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.axonivy.util.excel.importer.ExcelLoader;
import com.axonivy.util.excel.importer.ExcelReader;
import com.axonivy.util.excel.importer.ExcelReader.Column;

class TestExcelReader {

  @Test
  void parseColumns_xlsx(@TempDir Path dir) throws IOException {
    Path path = dir.resolve("customers.xlsx");
    TstRes.loadTo(path, "sample.xlsx");
    Workbook wb = ExcelLoader.load(path);
    List<Column> columns = ExcelReader.parseColumns(wb.getSheetAt(0));
    assertThat(columns).extracting(Column::name)
      .contains("Firstname", "Lastname");
    assertThat(columns).contains(
      new Column("Firstname", String.class),
      new Column("ZIP", Float.class),
      new Column("Birthdate", Float.class) // should be a date
    );
  }

}
