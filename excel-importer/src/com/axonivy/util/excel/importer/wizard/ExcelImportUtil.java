package com.axonivy.util.excel.importer.wizard;

import java.util.List;
import java.util.stream.Collectors;

import org.eclipse.core.resources.IResource;
import org.eclipse.jface.viewers.IStructuredSelection;

import ch.ivyteam.ivy.project.IIvyProject;
import ch.ivyteam.ivy.project.IIvyProjectManager;
import ch.ivyteam.ivy.project.IvyProjectNavigationUtil;

@SuppressWarnings("restriction")
final class ExcelImportUtil {

  static final String PAGE_ICON = "/resources/icons/cms_wizard_header.png";
  static final String DEFAULT_EXTENSION = ".xlsx";
  static final String[] IMPORT_TYPE = new String[] {"*" + DEFAULT_EXTENSION};
  static final String DESTINATION_KEY = "Destinations";
  static final String DEFAULT_FILTER_PATH = javax.swing.filechooser.FileSystemView.getFileSystemView()
          .getDefaultDirectory().getAbsolutePath();

  static List<String> getIvyProjectNames() {
    return IIvyProjectManager.instance().getIvyProjects().stream()
            .filter(p -> !p.isImmutable())
            .map(IIvyProject::getName)
            .collect(Collectors.toList());
  }

  static IIvyProject getFirstNonImmutableIvyProject(IStructuredSelection selection) {
    if (selection == null) {
      return null;
    }
    for (var localIterator = selection.toList().iterator(); localIterator.hasNext();) {
      var selectedObject = localIterator.next();
      if (selectedObject == null) {
        continue;
      }
      IIvyProject project = null;
      if (selectedObject instanceof IResource) {
        project = IvyProjectNavigationUtil.getIvyProject((IResource) selectedObject);
      }
      if (project != null && !project.isImmutable()) {
        return project;
      }
    }
    return null;
  }
}
