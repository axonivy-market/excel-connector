package com.axonivy.connector.excel.test;

import static org.assertj.core.api.Assertions.assertThat;

import java.nio.file.Path;
import java.util.Arrays;
import java.util.Date;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.axonivy.connector.excel.ExcelWriter;

import ch.ivyteam.ivy.environment.Ivy;
import ch.ivyteam.ivy.scripting.objects.DateTime;
import ch.ivyteam.ivy.scripting.objects.Recordset;

/**
 * This sample UnitTest runs java code in an environment as it would exists when
 * being executed in Ivy Process. Popular projects API facades, such as
 * {@link Ivy#persistence()} are setup and ready to be used.
 *
 * <p>
 * The test can either be run
 * <ul>
 * <li>in the Designer IDE ( <code>right click > run as > JUnit Test </code>
 * )</li>
 * <li>or in a Maven continuous integration build pipeline (
 * <code>mvn clean verify</code> )</li>
 * </ul>
 * </p>
 *
 * <p>
 * Detailed guidance on writing these kind of tests can be found in our <a href=
 * "https://developer.axonivy.com/doc/dev/concepts/testing/unit-testing.html">Unit
 * Testing docs</a>
 * </p>
 */

class SampleIvyTest {

  @Test
  void sample(@TempDir Path tempDir) throws Exception {
    assertThat(true).as("I can use Ivy API facade in tests").isEqualTo(true);

    Path path = tempDir.resolve("myexelfile.xls");
    Recordset recordset = new Recordset();

    recordset.addColumn("a1", Arrays.asList("ad", "sg"));
    recordset.addColumn("a2", Arrays.asList(45, 1234));
    recordset.addColumn("a3", Arrays.asList(new Date(), new DateTime()));

    new ExcelWriter().writeExcel(recordset, path.toAbsolutePath().toString());

    // read my generated file
  }

}
