package com.axonivy.connector.excel.test;

import static org.assertj.core.api.Assertions.assertThat;

import java.nio.file.Path;
import java.util.Date;
import java.util.List;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.axonivy.connector.excel.ExcelReader;
import com.axonivy.connector.excel.ExcelWriter;

import ch.ivyteam.ivy.scripting.objects.DateTime;
import ch.ivyteam.ivy.scripting.objects.Recordset;

class TestExcelConnector {

  private String excelFilePath;

  @BeforeEach
  void setup(@TempDir Path tempDir) throws Exception {
    excelFilePath = tempDir.resolve("myexelfile.xls").toAbsolutePath().toString();
  }

  @Test
  void writeAndRead_string() throws Exception {
    Recordset recordset = new Recordset();
    recordset.addColumn("a1", List.of("ad", "sg"));
    new ExcelWriter().execute(recordset, excelFilePath);
    Recordset excelRec = new ExcelReader().getRecordset(excelFilePath);
    assertThat(excelRec).isEqualTo(recordset);
  }

  @Test
  void writeAndRead_integer() throws Exception {
    Recordset recordset = new Recordset();
    recordset.addColumn("a1", List.of(1, 10));
    Recordset expectedRecordset = new Recordset();
    expectedRecordset.addColumn("a1", List.of(1.0, 10.0));
    new ExcelWriter().execute(recordset, excelFilePath);
    Recordset excelRec = new ExcelReader().getRecordset(excelFilePath);
    assertThat(excelRec).isEqualTo(expectedRecordset);
  }

  @Test
  void writeAndRead_double() throws Exception {
    Recordset recordset = new Recordset();
    recordset.addColumn("a1", List.of(1.5, 10.5932));
    new ExcelWriter().execute(recordset, excelFilePath);
    Recordset excelRec = new ExcelReader().getRecordset(excelFilePath);
    assertThat(excelRec).isEqualTo(recordset);
  }

  @Test
  void writeAndRead_date() throws Exception {
    Date d = new Date();
    DateTime dt = new DateTime();
    Recordset recordset = new Recordset();
    recordset.addColumn("a1", List.of(d, dt));
    Recordset expectedRecordset = new Recordset();
    expectedRecordset.addColumn("a1", List.of(d.toString(), dt.toString()));
    new ExcelWriter().execute(recordset, excelFilePath);
    Recordset excelRec = new ExcelReader().getRecordset(excelFilePath);
    assertThat(excelRec).isEqualTo(expectedRecordset);
  }

  @Test
  void writeAndRead_varyingRowSize() throws Exception {
    Recordset recordset = new Recordset();
    recordset.addColumn("a1", List.of("ad", "sg"));
    recordset.addColumn("a2", List.of("ad2", "sg2", "sg333"));
    recordset.addColumn("a3", List.of());
    new ExcelWriter().execute(recordset, excelFilePath);
    Recordset excelRec = new ExcelReader().getRecordset(excelFilePath);
    assertThat(excelRec).isEqualTo(recordset);
  }

  @Test
  void writeAndRead_nullRecordset() throws Exception {
    Recordset recordset = null;
    new ExcelWriter().execute(recordset, excelFilePath);
    Recordset excelRec = new ExcelReader().getRecordset(excelFilePath);
    assertThat(excelRec).isEqualTo(new Recordset());
  }

}
