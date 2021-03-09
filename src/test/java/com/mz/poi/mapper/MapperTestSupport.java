package com.mz.poi.mapper;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import org.junit.Before;

public class MapperTestSupport {

  public static String testDir = "test_excel";

  @Before
  public void init() throws IOException {
    Files.createDirectories(Paths.get(testDir));
  }
}
