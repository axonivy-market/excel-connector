package com.axonivy.util.excel.importer.wizard;

import java.io.File;
import java.lang.reflect.InvocationTargetException;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;
import java.util.Objects;

import org.apache.commons.lang3.StringUtils;
import org.eclipse.jface.wizard.IWizardPage;
import org.eclipse.jface.wizard.WizardPage;
import org.eclipse.swt.SWT;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Listener;
import org.eclipse.swt.widgets.Widget;

import ch.ivyteam.ivy.designer.ui.wizard.restricted.WizardStatus;
import ch.ivyteam.swt.dialogs.SwtCommonDialogs;

public class ExcelImportWizardPage extends WizardPage implements IWizardPage, Listener {

  static final String PAGE_ID = "ImportExcel";
  private final ExcelImportProcessor processor;
  private ExcelUi ui;

  public ExcelImportWizardPage(ExcelImportProcessor processor) {
    super(PAGE_ID);
    setMessage(processor.getWizardPageOkMessage(PAGE_ID));
    this.processor = processor;
    setPageComplete(false);
  }

  private static class ExcelUi extends Composite {

    public final Combo destinationNameField;
    public final Combo sourceProjectField;
    public final Button destinationBrowseButton;
    public final Combo persistence;

    public ExcelUi(Composite parent) {
      super(parent, SWT.NONE);

      GridLayout layout = new GridLayout();
      layout.numColumns = 3;
      setLayout(layout);
      setLayoutData(new GridData(272));

      Label destinationLabel = new Label(this, 0);
      destinationLabel.setText("From file");
      destinationNameField = new Combo(this, 2052);
      var dataDest = new GridData(768);
      dataDest.widthHint = 250;
      destinationNameField.setLayoutData(dataDest);
      destinationBrowseButton = new Button(this, 8);
      destinationBrowseButton.setText("Browse ...");

      Label sourceLabel = new Label(this, 0);
      sourceLabel.setText("Project");
      this.sourceProjectField = new Combo(this, 2060);
      GridData data = new GridData(768);
      data.widthHint = 250;
      data.horizontalSpan = 2;
      sourceProjectField.setLayoutData(data);

      Label unitLabel = new Label(this, SWT.NONE);
      unitLabel.setText("Persistence");
      this.persistence = new Combo(this, 2060);
      GridData data3 = new GridData(768);
      data3.widthHint = 250;
      data3.horizontalSpan = 2;
      persistence.setLayoutData(data3);
    }
  }

  @Override
  public void createControl(Composite parent) {
    this.ui = new ExcelUi(parent);

    String[] destinations = getDialogSettings().getArray(ExcelImportUtil.DESTINATION_KEY);
    if (destinations != null) {
      ui.destinationNameField.setText(destinations[0]);
      for (String destination : destinations) {
        if (destination.endsWith(ExcelImportUtil.DEFAULT_EXTENSION)) {
          ui.destinationNameField.add(destination);
          handleInputChanged();
        }
      }
    }

    for (String projectName : ExcelImportUtil.getIvyProjectNames()) {
      ui.sourceProjectField.add(projectName);
    }
    ui.sourceProjectField.setText(processor.getSelectedSourceProjectName());
    ui.sourceProjectField.addListener(SWT.Modify, this);
    ui.sourceProjectField.addListener(SWT.Selection, this);

    ui.destinationNameField.addListener(SWT.Modify, this);
    ui.destinationNameField.addListener(SWT.Selection, this);
    ui.destinationBrowseButton.addListener(SWT.Selection, this);

    ui.persistence.addListener(SWT.Modify, this);
    ui.persistence.addListener(SWT.Selection, this);

    setButtonLayoutData(ui.destinationBrowseButton);
    setControl(ui);
  }

  @Override
  public void handleEvent(Event event) {
    Widget source = event.widget;
    if (source.equals(ui.destinationBrowseButton)) {
      handleDestinationBrowseButtonPressed();
    } else {
      handleInputChanged();
    }
  }

  boolean finish() {
    if (processor.wizardFinishInvoked() && executeImport()) {
      saveDialogSettings();
      return true;
    }
    return false;
  }

  protected void handleInputChanged() {
    var status = WizardStatus.createOkStatus();

    status.merge(processor.setImportFile(ui.destinationNameField.getText()));

    String newProject = ui.sourceProjectField.getText();
    var sameProject = Objects.equals(processor.getSelectedSourceProjectName(), newProject);
    status.merge(processor.setSource(newProject));
    if (!sameProject) {
      ui.persistence.setItems(processor.units().toArray(String[]::new)); // update
    }

    status.merge(processor.setPersistence(ui.persistence.getText()));

    setPageComplete(status.isLowerThan(WizardStatus.ERROR));
    if (status.isOk()) {
      setMessage(processor.getWizardPageOkMessage(PAGE_ID), 0);
    } else if (status.isFatal()) {
      SwtCommonDialogs.openBugDialog(getControl(), status.getFatalError());
    } else {
      setMessage(status.getMessage(), status.getSeverity());
    }
  }

  private void saveDialogSettings() {
    List<String> destinations = new LinkedList<String>(Arrays.asList(ui.destinationNameField.getItems()));
    String path = ui.destinationNameField.getText();
    String lowerCasePath = path.toLowerCase();
    if (destinations.contains(path)) {
      destinations.remove(path);
      destinations.add(0, path);
      getDialogSettings().put(ExcelImportUtil.DESTINATION_KEY, destinations.toArray(String[]::new));
    } else if (lowerCasePath.endsWith(ExcelImportUtil.DEFAULT_EXTENSION)) {
      if (destinations.size() == 10) {
        destinations.remove(destinations.size() - 1);
      }
      destinations.add(0, path);
      getDialogSettings().put(ExcelImportUtil.DESTINATION_KEY, destinations.toArray(String[]::new));
    }
  }

  private void handleDestinationBrowseButtonPressed() {
    FileDialog dialog = new FileDialog(getContainer().getShell(), 0);
    dialog.setFilterExtensions(ExcelImportUtil.IMPORT_TYPE);
    dialog.setText("Select import file");
    dialog.setFilterPath(StringUtils.EMPTY);
    String currentSourceString = ui.destinationNameField.getText();
    dialog.setFilterPath(currentSourceString);
    String selectedFileName = dialog.open();
    if (selectedFileName != null) {
      ui.destinationNameField.setText(selectedFileName);
    }
  }

  private boolean executeImport() {
    if (ui.destinationNameField.getText().lastIndexOf(File.separator) == -1) {
      ui.destinationNameField.setText(
              ExcelImportUtil.DEFAULT_FILTER_PATH + File.separator + ui.destinationNameField.getText());
      processor.setImportFile(ui.destinationNameField.getText());
    }
    try {
      getContainer().run(true, true, processor);
    } catch (InterruptedException localInterruptedException) {
      return false;
    } catch (InvocationTargetException e) {
      SwtCommonDialogs.openBugDialog(getControl(), e.getTargetException());
      return false;
    }
    var status = processor.getStatus();
    if (status.isOK()) {
      SwtCommonDialogs.openInformationDialog(getShell(), "Express Import", "Successfully imported");
    } else {
      SwtCommonDialogs.openErrorDialog(getContainer().getShell(),
              "Problems during import of Excel as Dialog", status.getMessage(), status.getException());
      return false;
    }
    return true;
  }
}
