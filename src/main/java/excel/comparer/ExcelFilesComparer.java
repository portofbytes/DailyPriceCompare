package excel.comparer;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.FileSystems;

import org.apache.commons.math3.util.Precision;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFilesComparer {
	private static String fileSeparator = System.getProperty("file.separator");
	private static final String oldFile = fileSeparator	+ "old.xlsx";
	private static final String newFile = fileSeparator	+ "new.xlsx";
	private static final String comparisonFile = fileSeparator	+ "comparisonFile.xlsx";
	private static final int SHEETINDEX = 0;
	
	public static void main(String[] args)
			throws IOException, InvalidFormatException {

		String oldFilePath = FileSystems.getDefault().getPath("")
				.toAbsolutePath().toString() + fileSeparator + "Excel Files" + oldFile;

		String newFilePath = FileSystems.getDefault().getPath("")
				.toAbsolutePath().toString() + fileSeparator + "Excel Files" + newFile;
		
		String comparisonFilePath = FileSystems.getDefault().getPath("")
				.toAbsolutePath().toString() + fileSeparator + "Excel Files" + comparisonFile;

		// Read old file and convert to workbook
		
		InputStream newInp = new FileInputStream(newFilePath);
			Workbook newWorkbook = WorkbookFactory.create(newInp);
		
		InputStream oldInp = new FileInputStream(oldFilePath);
			Workbook oldWorkbook = WorkbookFactory.create(oldInp);
		
		// File manipulation
		ExcelFilesComparer excelComparer = new ExcelFilesComparer();
		newWorkbook = excelComparer.compareFiles(oldWorkbook, newWorkbook);
		
	    // Write the output to a file
	    try (OutputStream fileOut = new FileOutputStream(comparisonFilePath)) {
	    	newWorkbook.write(fileOut);
	    }
		oldWorkbook.close();
		newWorkbook.close();
	}

	/*
	 * Compare two excel files with one sheet only.
	 * */
	private Workbook compareFiles(Workbook oldWorkbook,	Workbook newWorkbook) throws IOException {
		
		

		
		CreationHelper factory = newWorkbook.getCreationHelper();
		
		// get working worksheet for both workbooks
		Sheet newFileSheet = newWorkbook.getSheetAt(SHEETINDEX);
		Sheet oldFileSheet = oldWorkbook.getSheetAt(SHEETINDEX);
		
		// loop through all the rows of the new file and get the ID in each row
		// Decide which rows to process
		int rowStart = Math.min(0, newFileSheet.getFirstRowNum());
		int rowEnd = getNumberOfRows(newFileSheet);
		
		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row newSheetRow = newFileSheet.getRow(rowNum);
			if (newSheetRow == null) {
				// This whole row is empty. Handle it as needed
				continue;
			} else {
				// get SKU number of row
				Cell newSheetCell = newSheetRow.getCell(0);
				CellStyle cellStyle = newWorkbook.createCellStyle();
				//to enable newlines you need set a cell styles with wrap=true
				cellStyle.setWrapText(true);
				cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
				cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);	
								
				String SKU = getCellValue(newSheetCell);
				
				// get Working Row of the old sheet
				Row oldSheetRow = getOldSheetRow(oldFileSheet, SKU);
				
				if (oldSheetRow != null) {
					int lastColumn = getNrColumns(oldFileSheet);
					for (int col = 0; col < lastColumn; col++) {
						String newSheetCellValue = "";
						String oldSheetCellValue = "";
						newSheetCell = newSheetRow.getCell(col,	Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

						// Do something useful with the cell's contents
						if (newSheetCell != null) {
							// Get the corresponding cell values from both the old sheet and the new sheet.
							Cell oldSheetCell = oldSheetRow.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
							newSheetCellValue = getCellValue(newSheetCell);
							oldSheetCellValue = getCellValue(oldSheetCell);
						
							//compare cell values of both old sheet and new sheet and do changes as required.
							if (newSheetCellValue.compareToIgnoreCase(oldSheetCellValue) != 0) {
								System.out.println(SKU + ", " + newSheetCellValue + ", " + oldSheetCellValue);				
								Cell newCell = newSheetRow.createCell(newSheetCell.getColumnIndex() + 1, CellType.STRING);
						        if(newSheetRow.getRowNum() == 0)
						        	newCell.setCellValue("NEW-COLUMN");
								newCell.setCellValue(newSheetCellValue + "\r\n" + oldSheetCellValue);
								newSheetCell.setCellStyle(cellStyle);
								newFileSheet.autoSizeColumn(0);
								
								
								// When the comment box is visible, have it show in a 1x3 space
								Drawing<?> drawing = newFileSheet.createDrawingPatriarch();
								ClientAnchor anchor = factory.createClientAnchor();
								anchor.setCol1(newSheetCell.getColumnIndex());
								anchor.setCol2(newSheetCell.getColumnIndex()+1);
								anchor.setRow1(newSheetRow.getRowNum());
								anchor.setRow2(newSheetRow.getRowNum()+3);

								// Create the comment and set the text+author
								Comment comment = drawing.createCellComment(anchor);
								RichTextString str = factory.createRichTextString(oldSheetCellValue);
								comment.setString(str);
								comment.setAuthor("Author");
								
								newSheetCell.setCellComment(comment);
							} // end if.
						} // end else.
					} // end for loop.

				} // end if.
			} // end else
		} // end for loop

		System.out.println("end of file reached");
		return newWorkbook;
	}// end method
	

	/*
	 * Get the row from the old sheet based on the SKU number from the new sheet
	 * */
	private Row getOldSheetRow(Sheet oldFileSheet, String newSheetSKU) {
		int rowStart = Math.min(0, oldFileSheet.getFirstRowNum());
		int rowEnd = Math.max(20000, oldFileSheet.getLastRowNum());
		String oldSheetSKU = "";
		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row oldSheetRow = oldFileSheet.getRow(rowNum);
			if (oldSheetRow == null) {
				// This whole row is empty. Handle it as needed
				continue;
			} else {
				// get SKU number of the row
				Cell cell = oldSheetRow.getCell(0);
				oldSheetSKU = getCellValue(cell);

				if (newSheetSKU.equalsIgnoreCase(oldSheetSKU)) {
					return oldSheetRow;
				} //end if

			}// end else
		}

		return null;
	} // end getOldSheetRow
	
	/*
	 Get the last row number
	 */
	public int getNumberOfRows(Sheet sheet ) {

		int rowNum = Math.max(20000, sheet.getLastRowNum());

		System.out.println("Comparing " + rowNum + " rows from " + newFile + " with " + oldFile);

		return rowNum;
	}	
	/*
	 Get the last cell number
	 */
	public int getNrColumns(Sheet sheet) {

		// get header row
		Row headerRow = sheet.getRow(0);
		int nrCol = headerRow.getLastCellNum();

	    System.out.println("Found " + nrCol + " columns.");
		return nrCol;

	}	

	/*
	 * Takes an existing Cell and merges all the styles and formula into the new
	 * one
	 */
	private void cloneCell(Cell cNew, Cell cOld) {
		cNew.setCellComment(cOld.getCellComment());
		cNew.setCellStyle(cOld.getCellStyle());
		cNew.setCellValue(getCellValue(cOld));
	}// end method
	/*
	 * Get the value of a given cell
	 */
	private String getCellValue(Cell cell) {
		if (cell != null) {
			switch (cell.getCellType()) {
				case STRING :
					return cell.getStringCellValue().strip();
				case NUMERIC :
					return String.valueOf( Precision.round(	cell.getNumericCellValue(), 2)).strip();
				case BLANK :
					return "";
				case BOOLEAN :
					return String.valueOf(
							cell.getBooleanCellValue()).strip();
				case ERROR :
					return String.valueOf(
							cell.getErrorCellValue()).strip();
				case FORMULA :
					return String.valueOf(
							cell.getCellFormula()).strip();
				case _NONE :
					return "";
				default :
					return "";
				}// end switch.
		}else {
			return "";
		}// end else if

	}// end method		

}// end class
