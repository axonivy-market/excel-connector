package com.axonivy.util.excel.test;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.axonivy.util.excel.importer.EntityClassReader;

import ch.ivyteam.ivy.environment.IvyTest;

@IvyTest
public class TestEntityClassCreator {

  private EntityClassReader reader;

  @Test
  void readToEntity(@TempDir Path dir) throws IOException {
    Path path = dir.resolve("customers.xlsx");
    loadTo(path, "sample.xlsx");

    var entity = reader.getEntity(path.toFile().toString());
    assertThat(entity).isNotNull();
  }

  @BeforeEach
  void setup() {
    this.reader = new EntityClassReader();
  }

  private static void loadTo(Path path, String resource) throws IOException {
    try(InputStream is = TestEntityClassCreator.class.getResourceAsStream(resource)) {
      Files.copy(is, path, StandardCopyOption.REPLACE_EXISTING);
    }
  }

}
