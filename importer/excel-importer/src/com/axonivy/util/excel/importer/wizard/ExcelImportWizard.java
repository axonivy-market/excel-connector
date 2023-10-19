package com.axonivy.util.excel.importer.wizard;

import org.eclipse.jface.viewers.IStructuredSelection;
import org.eclipse.jface.wizard.Wizard;
import org.eclipse.ui.IImportWizard;
import org.eclipse.ui.IWorkbench;

import ch.ivyteam.ivy.designer.ide.DesignerIDEPlugin;

public class ExcelImportWizard extends Wizard implements IImportWizard {

  private ExcelImportProcessor processor;

  public ExcelImportWizard() {
    setWindowTitle("Import Excel as Dialog");
    var workbenchSettings = DesignerIDEPlugin.getDefault().getDialogSettings();
    var wizardSettings = workbenchSettings.getSection(ExcelImportWizard.class.getName());
    if (wizardSettings == null) {
      wizardSettings = workbenchSettings.addNewSection(ExcelImportWizard.class.getName());
    }
    setDialogSettings(wizardSettings);
  }

  @Override
  public void init(IWorkbench workbench, IStructuredSelection selection) {
    processor = new ExcelImportProcessor(selection);
  }

  @Override
  public void addPages() {
    addPage(new ExcelImportWizardPage(processor));
  }

  @Override
  public boolean performFinish() {
    return ((ExcelImportWizardPage) getPage(ExcelImportWizardPage.PAGE_ID)).finish();
  }

  @Override
  public boolean needsProgressMonitor() {
    return true;
  }
}
