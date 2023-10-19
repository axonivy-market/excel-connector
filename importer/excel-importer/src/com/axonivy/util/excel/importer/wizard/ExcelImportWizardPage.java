package com.axonivy.util.excel.importer.wizard;

import java.io.File;
import java.lang.reflect.InvocationTargetException;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.eclipse.jface.wizard.IWizardPage;
import org.eclipse.jface.wizard.WizardPage;
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
  private Combo destinationNameField;
  private Combo sourceProjectField;
  private Button destinationBrowseButton;
  private final ExcelImportProcessor processor;

  public ExcelImportWizardPage(ExcelImportProcessor processor) {
    super(PAGE_ID);
    setMessage(processor.getWizardPageOkMessage(PAGE_ID));
    this.processor = processor;
    setPageComplete(false);
  }

  @Override
  public void createControl(Composite parent) {
    Composite group = new Composite(parent, 0);
    GridLayout layout = new GridLayout();
    layout.numColumns = 3;
    group.setLayout(layout);
    group.setLayoutData(new GridData(272));
    Label sourceLabel = new Label(group, 0);
    sourceLabel.setText("Project");
    this.sourceProjectField = new Combo(group, 2060);
    GridData data = new GridData(768);
    data.widthHint = 250;
    data.horizontalSpan = 2;
    sourceProjectField.setLayoutData(data);
    for (String projectName : ExcelImportUtil.getIvyProjectNames()) {
      sourceProjectField.add(projectName);
    }
    sourceProjectField.setText(processor.getSelectedSourceProjectName());
    sourceProjectField.addListener(24, this);
    sourceProjectField.addListener(13, this);
    Label destinationLabel = new Label(group, 0);
    destinationLabel.setText("From file");
    destinationNameField = new Combo(group, 2052);
    String[] destinations = getDialogSettings().getArray(ExcelImportUtil.DESTINATION_KEY);
    if (destinations != null) {
      destinationNameField.setText(destinations[0]);
      for (String destination : destinations) {
        if (destination.endsWith(ExcelImportUtil.DEFAULT_EXTENSION)) {
          destinationNameField.add(destination);
          handleInputChanged();
        }
      }
    }
    destinationNameField.addListener(24, this);
    destinationNameField.addListener(13, this);
    data = new GridData(768);
    data.widthHint = 250;
    destinationNameField.setLayoutData(data);
    destinationBrowseButton = new Button(group, 8);
    destinationBrowseButton.setText("Browse ...");
    destinationBrowseButton.addListener(13, this);
    setButtonLayoutData(destinationBrowseButton);
    setControl(group);
  }

  @Override
  public void handleEvent(Event event) {
    Widget source = event.widget;
    if (source.equals(destinationBrowseButton)) {
      handleDestinationBrowseButtonPressed();
    } else if (source.equals(sourceProjectField) || source.equals(destinationNameField)) {
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
    status.merge(processor.setImportFile(destinationNameField.getText()));
    status.merge(processor.setSource(this.sourceProjectField.getText()));
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
    List<String> destinations = new LinkedList<String>(Arrays.asList(destinationNameField.getItems()));
    String path = destinationNameField.getText();
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
    String currentSourceString = destinationNameField.getText();
    dialog.setFilterPath(currentSourceString);
    String selectedFileName = dialog.open();
    if (selectedFileName != null) {
      destinationNameField.setText(selectedFileName);
    }
  }

  private boolean executeImport() {
    if (destinationNameField.getText().lastIndexOf(File.separator) == -1) {
      destinationNameField.setText(
              ExcelImportUtil.DEFAULT_FILTER_PATH + File.separator + destinationNameField.getText());
      processor.setImportFile(destinationNameField.getText());
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
