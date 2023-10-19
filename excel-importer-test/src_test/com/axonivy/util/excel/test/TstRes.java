package com.axonivy.util.excel.test;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

public class TstRes {

  public static void loadTo(Path path, String resource) throws IOException {
    try(InputStream is = TstRes.class.getResourceAsStream(resource)) {
      Files.copy(is, path, StandardCopyOption.REPLACE_EXISTING);
    }
  }

}