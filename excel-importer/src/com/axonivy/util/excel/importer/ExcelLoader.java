package com.axonivy.util.excel.importer;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelLoader {

  public static Workbook load(Path filePath) {
    if (filePath.getFileName().toString().endsWith(".xls")) {
      return openXls(filePath);
    } else {
      return openXlsx(filePath);
    }
  }

  private static Workbook openXls(Path filePath) {
    try (InputStream input = Files.newInputStream(filePath);
            POIFSFileSystem fs = new POIFSFileSystem(input);
            HSSFWorkbook workbook = new HSSFWorkbook(fs);) {
      return workbook;
    } catch (Exception ex) {
      throw new RuntimeException("Could not read excel file from " + filePath, ex);
    }
  }

  private static Workbook openXlsx(Path filePath) {
    try (InputStream input = Files.newInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(input);) {
      return workbook;
    } catch (Exception ex) {
      throw new RuntimeException("Could not read excel file from " + filePath, ex);
    }
  }

}