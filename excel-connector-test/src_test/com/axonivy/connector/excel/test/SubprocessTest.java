package com.axonivy.connector.excel.test;

import static org.assertj.core.api.Assertions.assertThat;

import java.util.List;

import org.junit.jupiter.api.Test;

import com.axonivy.connector.excel.ReadExcelData;
import com.axonivy.connector.excel.WriteExcelData;

import ch.ivyteam.ivy.bpm.engine.client.BpmClient;
import ch.ivyteam.ivy.bpm.engine.client.ExecutionResult;
import ch.ivyteam.ivy.bpm.engine.client.element.BpmElement;
import ch.ivyteam.ivy.bpm.engine.client.element.BpmProcess;
import ch.ivyteam.ivy.bpm.exec.client.IvyProcessTest;
import ch.ivyteam.ivy.scripting.objects.Recordset;

@IvyProcessTest
public class SubprocessTest{

  private static final BpmProcess testeeReadExcel = BpmProcess.path("excel_connector/ReadExcel");
  private static final BpmProcess testeeWriteExcel = BpmProcess.path("excel_connector/WriteExcel");

  @Test
  public void callSubprocesses(BpmClient bpmClient){
    Recordset recordset = new Recordset();
    recordset.addColumn("a1", List.of("ad", "sg"));
    BpmElement writeStartable = testeeWriteExcel.elementName("call(WriteExcel(Recordset))");
    ExecutionResult writeResult = bpmClient.start().subProcess(writeStartable).execute(recordset);
    WriteExcelData writeData = writeResult.data().last();
    assertThat(writeData.getExcelFile().exists()).isTrue();

    BpmElement readStartable = testeeReadExcel.elementName("call(ReadExcel(File))");
    ExecutionResult readResult = bpmClient.start().subProcess(readStartable).execute(writeData.getExcelFile());
    ReadExcelData readData = readResult.data().last();
    assertThat(readData.getRecordset()).isEqualTo(recordset);
  }

}
