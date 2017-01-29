package com.monitorjbl.xlsx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Locale;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

public class StreamingWorkbookTest {
  @BeforeClass
  public static void init() {
    Locale.setDefault(Locale.ENGLISH);
  }

  @Test
  public void testIterateSheets() throws Exception {
      InputStream is = new FileInputStream(new File("src/test/resources/sheets.xlsx"));
      Workbook workbook = StreamingReader.builder().open(is);

      assertEquals(2, workbook.getNumberOfSheets());

      Sheet alpha = workbook.getSheetAt(0);
      Sheet zulu = workbook.getSheetAt(1);
      assertEquals("SheetAlpha", alpha.getSheetName());
      assertEquals("SheetZulu", zulu.getSheetName());

      Row rowA = alpha.rowIterator().next();
      Row rowZ = zulu.rowIterator().next();

      assertEquals("stuff", rowA.getCell(0).getStringCellValue());
      assertEquals("yeah", rowZ.getCell(0).getStringCellValue());
      
      workbook.close();
      is.close();
  }

  @Test
  public void testHiddenCells() throws Exception {
      InputStream is = new FileInputStream(new File("src/test/resources/hidden.xlsx"));
      Workbook workbook = StreamingReader.builder().open(is);
      
      assertEquals(3, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);

      assertFalse("Column 0 should not be hidden", sheet.isColumnHidden(0));
      assertTrue("Column 1 should be hidden", sheet.isColumnHidden(1));
      assertFalse("Column 2 should not be hidden", sheet.isColumnHidden(2));

      assertFalse("Row 0 should not be hidden", sheet.rowIterator().next().getZeroHeight());
      assertTrue("Row 1 should be hidden", sheet.rowIterator().next().getZeroHeight());
      assertFalse("Row 2 should not be hidden", sheet.rowIterator().next().getZeroHeight());
      
      workbook.close();
      is.close();
  }

  @Test
  public void testHiddenSheets() throws Exception {
      InputStream is = new FileInputStream(new File("src/test/resources/hidden.xlsx"));
      Workbook workbook = StreamingReader.builder().open(is);
      
      assertEquals(3, workbook.getNumberOfSheets());
      assertFalse(workbook.isSheetHidden(0));

      assertTrue(workbook.isSheetHidden(1));
      assertFalse(workbook.isSheetVeryHidden(1));

      assertFalse(workbook.isSheetHidden(2));
      assertTrue(workbook.isSheetVeryHidden(2));
      
      workbook.close();
      is.close();
  }

  @Test
  public void testFormulaCells() throws Exception {
      InputStream is = new FileInputStream(new File("src/test/resources/formula_cell.xlsx"));
      Workbook workbook = StreamingReader.builder().open(is);
      
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);

      Iterator<Row> rowIterator = sheet.rowIterator();
      rowIterator.next();
      rowIterator.next();
      Row row3 = rowIterator.next();
      Cell A3 = row3.getCell(0);

      assertEquals("Cell A3 should be of type formula", CELL_TYPE_FORMULA, A3.getCellType());
      assertEquals("Cell A3's value should be of type numeric", CELL_TYPE_NUMERIC, A3.getCachedFormulaResultType());
      assertEquals("Wrong formula", "SUM(A1:A2)", A3.getCellFormula());

      workbook.close();
      is.close();
  }
}
