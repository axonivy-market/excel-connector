package com.axonivy.util.excel.test;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.nio.file.Path;

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
    TstRes.loadTo(path, "sample.xlsx");

    var entity = reader.getEntity(path);
    assertThat(entity).isNotNull();
  }

  @BeforeEach
  void setup() {
    this.reader = new EntityClassReader();
  }

}
