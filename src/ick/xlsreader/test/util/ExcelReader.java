package ick.xlsreader.test.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class ExcelReader {

	protected Workbook workbook;
	private FormulaEvaluator evaluator;

	private final String filePath;
		
	/** 
	 * Constructor
	 * 
	 * @param filePath
	 */
	public ExcelReader(final String filePath) {
		this.filePath = filePath;
	}

	public String getFilePath() {
		return filePath;
	}
		
	/**
	 * Init start
	 * 
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public void start() throws InvalidFormatException, IOException {
		workbook = WorkbookFactory.create(new FileInputStream(filePath));
		evaluator = workbook.getCreationHelper().createFormulaEvaluator();
	}
	

	/** All value of particular rowIdx
	 * 
	 * @param sheetIdx
	 * @param rowIdx
	 * @return
	 */
	public List<String> getRowCellsAt(final int sheetIdx, final int rowIdx) {

		List<String> result = new ArrayList<String>();
		Sheet sheet = workbook.getSheetAt(sheetIdx);
		Row row = sheet.getRow(rowIdx);

		if (row != null) {
			int cells = row.getLastCellNum();
			for (int c = 0; c < cells; c++) {
				Cell cell = row.getCell(c);
				if (cell != null) {
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_FORMULA:
						CellValue cellValue = evaluator.evaluate(cell);
						String formulVal = Double.toString(cellValue.getNumberValue());
						result.add(formulVal == null ? "" : formulVal);
						break;
					case Cell.CELL_TYPE_NUMERIC:
						String numVal = Double.toString(cell.getNumericCellValue());
						result.add(numVal == null ? "" : numVal);
						break;
					case Cell.CELL_TYPE_STRING:
						String strVal = cell.getRichStringCellValue().getString();
						result.add(strVal == null ? "" : strVal);
						break;
					case Cell.CELL_TYPE_BLANK:
						result.add("");
						break;
					}
				}
			}
		} else {
			result.add("");
		}
		return result;
	}

	/** All value of particular colIdx
	 * 
	 * @param sheetIdx
	 * @param rowIdx
	 * @param colIdx
	 */
	public List<String> getColCellsAt(final int sheetIdx, final int colIdx) {
		List<String> result = new ArrayList<String>();
		Sheet sheet = workbook.getSheetAt(sheetIdx);
		
		int rows = sheet.getLastRowNum();
		
		for (int r = 0; r < rows; r++) {
			Row row = sheet.getRow(r);

			if (row != null) {

				Cell cell = row.getCell(colIdx);
				if (cell != null) {
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_FORMULA:
						CellValue cellValue = evaluator.evaluate(cell);
						String formulVal = Double.toString(cellValue.getNumberValue());
						result.add(formulVal == null ? "" : formulVal);
						break;
					case Cell.CELL_TYPE_NUMERIC:
						String numVal = Double.toString(cell.getNumericCellValue());
						result.add(numVal == null ? "" : numVal);
						break;
					case Cell.CELL_TYPE_STRING:
						String strVal = cell.getRichStringCellValue().getString();
						result.add(strVal == null ? "" : strVal);
						break;
					case Cell.CELL_TYPE_BLANK:
						result.add("");
						break;
					}
				}
			}
		}
		return result;
	}

	/** value of particular rowIdx, colIdx
	 * 
	 * @param sheetIdx
	 * @param rowIdx
	 * @param colIdx
	 */
	public String getRowColCellsAt(final int sheetIdx, final int rowIdx, final int colIdx) {

		String result = "";
		Sheet sheet = workbook.getSheetAt(sheetIdx);
		Row row = sheet.getRow(rowIdx);

		if (row != null) {
			Cell cell = row.getCell(colIdx);
			if (cell != null) {
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_FORMULA:
					CellValue cellValue = evaluator.evaluate(cell);
					String formulVal = Double.toString(cellValue.getNumberValue());
					result = formulVal;
					break;
				case Cell.CELL_TYPE_NUMERIC:
					String numVal = Double.toString(cell.getNumericCellValue());
					result = numVal;
					break;
				case Cell.CELL_TYPE_STRING:
					String strVal = cell.getRichStringCellValue().getString();
					result = strVal;
					break;
				case Cell.CELL_TYPE_BLANK:
					result = "";
					break;
				}
			}

		} else {
			result = "";
		}
		return result;
	}


	protected Workbook getWorkbook() {
		return workbook;
	}

	protected void setWorkbook(final Workbook workbook) {
		this.workbook = workbook;
	}

	protected FormulaEvaluator getEvaluator() {
		return evaluator;
	}

	protected void setEvaluator(final FormulaEvaluator evaluator) {
		this.evaluator = evaluator;
	}


}
