package com.axonivy.util.excel.test;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.axonivy.util.excel.importer.EntityClassReader;

import ch.ivyteam.ivy.environment.IvyTest;

@IvyTest
public class TestEntityClassCreator {

  @Test
  void readToEntity(@TempDir Path dir) throws IOException {
    Path path = dir.resolve("customers.xlsx");
    try(InputStream is = TestEntityClassCreator.class.getResourceAsStream("sample.xlsx")) {
      Files.copy(is, path, StandardCopyOption.REPLACE_EXISTING);
    }

    var reader = new EntityClassReader();
    var entity = reader.getEntity(path.toFile().toString());
    assertThat(entity).isNotNull();
  }

}
