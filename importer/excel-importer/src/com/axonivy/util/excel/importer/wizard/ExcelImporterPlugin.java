package com.axonivy.util.excel.importer.wizard;

import org.eclipse.ui.plugin.AbstractUIPlugin;
import org.osgi.framework.BundleContext;


public class ExcelImporterPlugin extends AbstractUIPlugin {

  private static ExcelImporterPlugin plugin;

  public ExcelImporterPlugin() {
    super();
    plugin = this;
  }

  @Override
  public void start(BundleContext arg0) throws Exception {
  }

  @Override
  public void stop(BundleContext arg0) throws Exception {
  }

  public static ExcelImporterPlugin getPlugin() {
    return plugin;
  }

}
