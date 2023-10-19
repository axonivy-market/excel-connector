package com.axonivy.util.excel.importer.wizard;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import org.apache.commons.lang3.StringUtils;
import org.eclipse.core.resources.IWorkspace;
import org.eclipse.core.resources.ResourcesPlugin;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.core.runtime.IStatus;
import org.eclipse.core.runtime.Status;
import org.eclipse.core.runtime.SubMonitor;
import org.eclipse.jface.operation.IRunnableWithProgress;
import org.eclipse.jface.viewers.IStructuredSelection;

import com.axonivy.util.excel.importer.EntityClassReader;

import ch.ivyteam.eclipse.util.EclipseUtil;
import ch.ivyteam.eclipse.util.MonitorUtil;
import ch.ivyteam.ivy.designer.ui.wizard.restricted.IWizardSupport;
import ch.ivyteam.ivy.designer.ui.wizard.restricted.WizardStatus;
import ch.ivyteam.ivy.eclipse.util.EclipseUiUtil;
import ch.ivyteam.ivy.project.IIvyProject;
import ch.ivyteam.ivy.project.IIvyProjectManager;
import ch.ivyteam.ivy.scripting.dataclass.IDataClassManager;
import ch.ivyteam.ivy.scripting.dataclass.IProjectDataClassManager;
import ch.ivyteam.util.io.resource.FileResource;
import ch.ivyteam.util.io.resource.nio.NioFileSystemProvider;

@SuppressWarnings("restriction")
public class ExcelImportProcessor implements IWizardSupport, IRunnableWithProgress {

  private IIvyProject selectedSourceProject;
  private FileResource importFile;
  private IStatus status = Status.OK_STATUS;

  public ExcelImportProcessor(IStructuredSelection selection) {
    this.selectedSourceProject = ExcelImportUtil.getFirstNonImmutableIvyProject(selection);
  }

  @Override
  public String getWizardPageTitle(String pageId) {
    return "Import Excel as Dialog";
  }

  @Override
  public String getWizardPageOkMessage(String pageId) {
    return "Please specify an Excel file to import as Dialog";
  }

  @Override
  public boolean wizardFinishInvoked() {
    var okStatus = WizardStatus.createOkStatus();
    okStatus.merge(validateImportFileExits());
    okStatus.merge(validateSource());
    return okStatus.isLowerThan(WizardStatus.ERROR);
  }

  @Override
  public boolean wizardCancelInvoked() {
    return true;
  }

  @Override
  public void run(IProgressMonitor monitor) throws InvocationTargetException {
    SubMonitor progress = MonitorUtil.begin(monitor, "Importing", 1);
    try {
      status = Status.OK_STATUS;
      ResourcesPlugin.getWorkspace().run(m -> {
        var manager = IDataClassManager.instance().getProjectDataModelFor(selectedSourceProject.getProject());
        try {
          importExcel(manager, importFile);
        } catch (IOException ex) {
          status = EclipseUtil.createErrorStatus(ex);
        }
      }, null, IWorkspace.AVOID_UPDATE, progress);
    } catch (Exception ex) {
      status = EclipseUtil.createErrorStatus(ex);
    } finally {
      MonitorUtil.ensureDone(monitor);
    }
  }

  private void importExcel(IProjectDataClassManager manager, FileResource excel) throws IOException {
    String name = excel.name();
    var tempFile = Files.createTempFile(name, "."+StringUtils.substringAfterLast(name, "."));
    Files.copy(excel.read().inputStream(), tempFile, StandardCopyOption.REPLACE_EXISTING);

    var newEntity = new EntityClassReader(manager).getEntity(tempFile);
    EclipseUiUtil.openEditor(newEntity);
    newEntity.save();
  }

  String getSelectedSourceProjectName() {
    if (selectedSourceProject == null) {
      return StringUtils.EMPTY;
    }
    return selectedSourceProject.getName();
  }

  public WizardStatus setImportFile(String text) {
    if (text != null) {
      importFile = NioFileSystemProvider.create(Path.of("/")).root().file(text);
    } else {
      importFile = null;
    }
    return validateImportFileExits();
  }

  public WizardStatus setSource(String projectName) {
    selectedSourceProject = null;
    if (projectName != null) {
      selectedSourceProject = IIvyProjectManager.instance().getIvyProject(projectName);
    }
    return validateSource();
  }

  public IStatus getStatus() {
    return status;
  }

  private WizardStatus validateImportFileExits() {
    if (importFile == null || !importFile.exists()) {
      return WizardStatus.createErrorStatus("Import file does not exist");
    }
    return WizardStatus.createOkStatus();
  }

  private WizardStatus validateSource() {
    if (selectedSourceProject == null) {
      return WizardStatus.createErrorStatus("Please specify an Axon Ivy project");
    }
    return WizardStatus.createOkStatus();
  }
}
