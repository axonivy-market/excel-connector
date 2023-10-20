package com.axonivy.util.excel.test;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.eclipse.core.runtime.CoreException;
import org.eclipse.core.runtime.NullProgressMonitor;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.axonivy.util.excel.importer.EntityClassReader;
import com.axonivy.util.excel.importer.ExcelLoader;
import com.axonivy.util.excel.importer.ProcessDrawer;

import ch.ivyteam.ivy.environment.IvyTest;
import ch.ivyteam.ivy.process.IProcess;
import ch.ivyteam.ivy.process.model.Process;
import ch.ivyteam.ivy.process.model.element.activity.DialogCall;
import ch.ivyteam.ivy.process.model.element.activity.value.dialog.UserDialogId;
import ch.ivyteam.ivy.scripting.dataclass.IEntityClass;

@IvyTest
@SuppressWarnings("restriction")
public class TestProcessDrawer {

  private EntityClassReader reader;

  @Test
  void drawRunnerProcess(@TempDir Path dir) throws IOException, CoreException {
    Path path = dir.resolve("customers.xlsx");
    TstRes.loadTo(path, "sample.xlsx");

    Workbook wb = ExcelLoader.load(path);
    Sheet customerSheet = wb.getSheetAt(0);

    IEntityClass customer = reader.toEntity(customerSheet, "customer");
    IProcess processRdm = null;
    try {
      customer.save(new NullProgressMonitor());

      var drawer = new ProcessDrawer(reader.manager.getProject());
      processRdm = drawer.drawManager(customer);
      Process process = processRdm.getModel();
      assertThat(process.getProcessElements()).hasSize(3);
      DialogCall callHd = process.search().type(DialogCall.class).findOne();
      assertThat(callHd.getName()).contains("customer");
      assertThat(callHd.getTargetDialog().getId())
        .isNotEqualTo(UserDialogId.EMPTY);

    } finally {
      customer.getResource().delete(true, new NullProgressMonitor());
      if (processRdm != null) {
        processRdm.getResource().delete(true, new NullProgressMonitor());
      }
    }
  }

  @BeforeEach
  void setup() {
    this.reader = new EntityClassReader();
  }

}
