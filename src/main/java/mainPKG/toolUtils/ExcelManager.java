package mainPKG.toolUtils;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map.Entry;

public class ExcelManager {
	
	//================================================>Static Part<================================================
	private static HashMap<String, ExcelManager> map4AllResult = new HashMap<String, ExcelManager>();
	public static ExcelManager createAExcelMgr(String excelPathIn, String excelPathOut, String indexName) {
		ExcelManager excelMgr;
		if (map4AllResult.containsKey(indexName)) {
			excelMgr = map4AllResult.get(indexName);
		} else {
			excelMgr = new ExcelManager(excelPathIn, excelPathOut);
			map4AllResult.put(indexName, excelMgr);
		}
		return excelMgr;
	}
	public static void outputAllResult(boolean wantWrite) {
		if (!wantWrite) {
			return;
		}
		for (Entry<String, ExcelManager> entry : map4AllResult.entrySet()) {
			entry.getValue().outputResult();
		}
	}
	public static void outputWithClear(boolean wantWrite) {
		if (wantWrite) {
			outputAllResult(wantWrite);
			clear();
		}
	}
	public static void clear() {
		map4AllResult.clear();
	}
	
	//================================================>Active Part<================================================
	private Workbook rootWorkbook = null;
	public Workbook getRootWorkbook() {
		return rootWorkbook;
	}
	public CellStyle getCellStyleAt(int sheetIndex, int row, int col) {
		return rootWorkbook.getSheetAt(sheetIndex).getRow(row).getCell(col).getCellStyle();
	}
	public CellStyle getCellStyleAt(String sheetName, int row, int col) {
		return rootWorkbook.getSheetAt(rootWorkbook.getSheetIndex(sheetName)).getRow(row).getCell(col).getCellStyle();
	}
	private HashMap<String, ExcelSheet> map4Check_ID_Sheet = new HashMap();
	private String excelPathOut;
	
	public ExcelManager(String excelPathIn, String excelPathOut) {
		this(excelPathIn);
		
		this.excelPathOut = excelPathOut;
	}
	public ExcelManager(String excelPathIn) {
		rootWorkbook = ExcelUtils.getExcel4Xlsx(excelPathIn);
	}
	public ExcelManager() {
		this(null);
	}
	
	public ExcelSheet getSheetAt(int sheetIndex, int startRow, int startCol) {
		String sheetName = rootWorkbook.getSheetAt(sheetIndex).getSheetName();
		ExcelSheet excelSheet = map4Check_ID_Sheet.get(sheetName);
		if (excelSheet == null) {
			Sheet newSheet = rootWorkbook.getSheet(sheetName);
			if (newSheet == null) {
				newSheet = rootWorkbook.createSheet(sheetName);
			}
			excelSheet = new ExcelSheet(newSheet, startRow, startCol);
			map4Check_ID_Sheet.put(sheetName, excelSheet);
		}
		return excelSheet;
	}
	public ExcelSheet getSheet(String sheetName, int startRow, int startCol) {
		ExcelSheet excelSheet = map4Check_ID_Sheet.get(sheetName);
		if (excelSheet == null) {
			Sheet newSheet = rootWorkbook.getSheet(sheetName);
			if (newSheet == null) {
				newSheet = rootWorkbook.createSheet(sheetName);
			}
			excelSheet = new ExcelSheet(newSheet, startRow, startCol);
			map4Check_ID_Sheet.put(sheetName, excelSheet);
		}
		return excelSheet;
	}
	public void outputResult() {
		ExcelUtils.outputExcel4Xlsx(rootWorkbook, excelPathOut);
	}
	public void outputResult(String excelPathOut) {
		ExcelUtils.outputExcel4Xlsx(rootWorkbook, excelPathOut);
	}

}
