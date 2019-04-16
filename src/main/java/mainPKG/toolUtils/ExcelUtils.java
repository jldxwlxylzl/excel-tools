package mainPKG.toolUtils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelUtils {

	// ================================================>将Excel的列名转换为索引<================================================
	public static String transColIndex2Name(int colIndex) {
		if ( colIndex >= 0 ) {
			int devMain = colIndex / 26;
			int devTail = colIndex % 26;
			if ( devMain == 0 ) {
				return String.valueOf( (char) (devTail + 'A') );
			} else {
				return transColIndex2NameNotTail( devMain ) + String.valueOf( (char) (devTail + 'A') );
			}
		} else {
			return null;
		}
	}

	private static String transColIndex2NameNotTail(int colIndex) {
		if ( colIndex >= 0 ) {
			int devMain = colIndex / 26;
			int devTail = colIndex % 26 - 1;
			if (devTail == -1) {
				devMain--;
				devTail = 25;
			}
			if (devMain > 0) {
				return transColIndex2NameNotTail( devMain ) + String.valueOf( (char) (devTail + 'A') );
			} else {
				return String.valueOf( (char) (devTail + 'A') );
			}
		} else {
			return null;
		}
	}

	public static int transColName2Index(String colName) {
		colName = colName.toUpperCase();
		if ( colName.matches( "[A-Z]+" ) ) {
			if ( colName.length() == 1 ) {
				return colName.charAt( 0 ) - 'A';
			} else if ( colName.length() == 2 ) {
				int outNum = (colName.charAt( 0 ) - 'A' + 1) * 26;
				outNum += colName.charAt( 1 ) - 'A';
				return outNum;
			} else {
				return (colName.charAt( 0 ) - 'A' + 1) * (int) Math.pow( 26, colName.length() - 1 ) + transColName2Index( colName.substring( 1 ) );
			}
		} else {
			return -1;
		}
	}

	public static int transColName2Index(char colName) {
		if ( colName >= 'A' && colName <= 'Z' ) {
			return colName - 'A';
		} else if ( colName >= 'a' && colName <= 'z' ) {
			return colName - 'a';
		} else {
			return -1;
		}
	}

	// ================================================>文件路径修改<================================================
	/**
	 * 在文件名后添加信息
	 * @param filePath	文件路径
	 * @param info2Add	要添加的信息
	 * @return
	 */
	public static String addPostfix2FileName(String filePath, String info2Add) {
		if ( filePath.contains( "." ) ) {
			return filePath.replace( ".", info2Add + "." );
		} else {
			return filePath + info2Add;
		}
	}
	/**
	 * 在文件名前添加信息
	 * @param filePath	文件路径
	 * @param info2Add	要添加的信息
	 * @return
	 */
	public static String addPrefix2FileName(String filePath, String info2Add) {
		filePath = filePath.replace('\\', '/');
		if ( filePath.contains( "/" ) ) {
			int posSep = filePath.lastIndexOf('/')+1;
			return filePath.substring(0, posSep) + info2Add + filePath.substring(posSep);
		} else {
			return null;
		}
	}

	// ================================================>打开指定路径的Excel或其他文件<================================================
	/**
	 * 打开指定路径的文件
	 * @param filePath	文件路径
	 */
	public static void openFile(String filePath) {
		try {
			Runtime.getRuntime().exec( "cmd /c start " + filePath );
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// ================================================>单行单列操作<================================================
	/**
	 * 从sheetOpr的startRowNum行startColNum列开始,将datas4write(数组)依次写入
	 * @param startRowNum	要写入的行的序号
	 * @param startColNum	列序号的起始位置
	 * @param datas4write	要写入的数据数组
	 * @param sheetOpr	要写入的sheet页
	 * @param isRow		是否为行操作
	 */
	public static void writeARowCols(int startRowNum, int startColNum, String[] datas4write, Sheet sheetOpr, boolean isRow) {
		Row rowTmp;
		Cell cellTmp;
		if ( isRow ) {
			rowTmp = sheetOpr.getRow( startRowNum );
			if ( rowTmp == null ) {
				rowTmp = sheetOpr.createRow( startRowNum );
			}
			for (int i = 0 ; i < datas4write.length ; i++) {
				cellTmp = rowTmp.getCell( startColNum + i );
				if ( cellTmp == null ) {
					cellTmp = rowTmp.createCell( startColNum + i );
				}
				cellTmp.setCellValue( datas4write[i] );
			}
		} else {
			for (int i = 0 ; i < datas4write.length ; i++) {
				rowTmp = sheetOpr.getRow( startRowNum + i );
				if ( rowTmp == null ) {
					rowTmp = sheetOpr.createRow( startRowNum + i );
				}
				cellTmp = rowTmp.getCell( startColNum );
				if ( cellTmp == null ) {
					cellTmp = rowTmp.createCell( startColNum );
				}
				cellTmp.setCellValue( datas4write[i] );
			}
		}

	}
	public static void writeARowCols(int startRowNum, int startColNum, int[] datas4write, Sheet sheetOpr, boolean isRow) {
		Row rowTmp;
		Cell cellTmp;
		if ( isRow ) {
			rowTmp = sheetOpr.getRow( startRowNum );
			if ( rowTmp == null ) {
				rowTmp = sheetOpr.createRow( startRowNum );
			}
			for (int i = 0 ; i < datas4write.length ; i++) {
				cellTmp = rowTmp.getCell( startColNum + i );
				if ( cellTmp == null ) {
					cellTmp = rowTmp.createCell( startColNum + i );
				}
				cellTmp.setCellValue( datas4write[i] );
			}
		} else {
			for (int i = 0 ; i < datas4write.length ; i++) {
				rowTmp = sheetOpr.getRow( startRowNum + i );
				if ( rowTmp == null ) {
					rowTmp = sheetOpr.createRow( startRowNum + i );
				}
				cellTmp = rowTmp.getCell( startColNum );
				if ( cellTmp == null ) {
					cellTmp = rowTmp.createCell( startColNum );
				}
				cellTmp.setCellValue( datas4write[i] );
			}
		}

	}
	public static void writeARowCols(int startRowNum, int startColNum, boolean[] datas4write, Sheet sheetOpr, boolean isRow) {
		Row rowTmp;
		Cell cellTmp;
		if ( isRow ) {
			rowTmp = sheetOpr.getRow( startRowNum );
			if ( rowTmp == null ) {
				rowTmp = sheetOpr.createRow( startRowNum );
			}
			for (int i = 0 ; i < datas4write.length ; i++) {
				cellTmp = rowTmp.getCell( startColNum + i );
				if ( cellTmp == null ) {
					cellTmp = rowTmp.createCell( startColNum + i );
				}
				cellTmp.setCellValue( datas4write[i] );
			}
		} else {
			for (int i = 0 ; i < datas4write.length ; i++) {
				rowTmp = sheetOpr.getRow( startRowNum + i );
				if ( rowTmp == null ) {
					rowTmp = sheetOpr.createRow( startRowNum + i );
				}
				cellTmp = rowTmp.getCell( startColNum );
				if ( cellTmp == null ) {
					cellTmp = rowTmp.createCell( startColNum );
				}
				cellTmp.setCellValue( datas4write[i] );
			}
		}

	}

	/**
	 * 从sheetOpr的startRowNum行startColNum列开始,将datas4write(List)依次写入
	 * @param startRowNum	要写入的行的序号
	 * @param startColNum	列序号的起始位置
	 * @param datas4write	要写入的数据List
	 * @param sheetOpr	要写入的sheet页
	 * @param isRow		是否为行操作
	 */
	public static void writeARowCols(int startRowNum, int startColNum, List<Object> datas4write, Sheet sheetOpr, boolean isRow) {
		Row rowTmp;
		Cell cellTmp;
		Object dataTmp;
		if ( isRow ) {
			rowTmp = sheetOpr.getRow( startRowNum );
			if ( rowTmp == null ) {
				rowTmp = sheetOpr.createRow( startRowNum );
			}
			for (int i = 0 ; i < datas4write.size() ; i++) {
				cellTmp = rowTmp.getCell( startColNum + i );
				if ( cellTmp == null ) {
					cellTmp = rowTmp.createCell( startColNum + i );
				}
				dataTmp = datas4write.get( i );
				if (dataTmp instanceof String) {
					cellTmp.setCellValue( (String)dataTmp );
				} else if (dataTmp instanceof Integer) {
					cellTmp.setCellValue( (Integer)dataTmp );
				} else if (dataTmp instanceof Boolean) {
					cellTmp.setCellValue( (Boolean) dataTmp );
				}
			}
		} else {
			for (int i = 0 ; i < datas4write.size() ; i++) {
				rowTmp = sheetOpr.getRow( startRowNum + i );
				if ( rowTmp == null ) {
					rowTmp = sheetOpr.createRow( startRowNum + i );
				}
				cellTmp = rowTmp.getCell( startColNum );
				if ( cellTmp == null ) {
					cellTmp = rowTmp.createCell( startColNum );
				}
				dataTmp = datas4write.get( i );
				if (dataTmp instanceof String) {
					cellTmp.setCellValue( (String)dataTmp );
				} else if (dataTmp instanceof Integer) {
					cellTmp.setCellValue( (Integer)dataTmp );
				} else if (dataTmp instanceof Boolean) {
					cellTmp.setCellValue( (Boolean)dataTmp );
				}
			}
		}
	}

	// ================================================>多行多列操作<================================================
	/**
		 * 复制fromSheet中{copyLenth 复制的行或列的长度}的数据,复制起点 ---> 单元格[fromRowNum : fromColNum]|
		 * 粘贴到toSheet中,粘贴起点 ---> 单元格[toRowNum : toColNum]
		 * @param fromRowNum	复制的起始行序号
		 * @param fromColNum	复制的起始列序号
		 * @param fromSheet	源数据Sheet页
		 * @param toRowNum	粘贴的起始行序号
		 * @param toColNum	粘贴的起始列序号
		 * @param toSheet	目标Sheet页
		 * @param copyLenth	复制的行或列数量
		 * @param isRow		是否为行操作
		 */
	public static void copyRowCols(int fromRowNum, int fromColNum, Sheet fromSheet, int toRowNum, int toColNum, Sheet toSheet, int copyLenth, boolean isRow) {
		// 备用变量声明
		Row rowFrom, rowTo;
		Cell cellFrom, cellTo;
		int loopLenth = 0;
		if ( isRow ) {
			// 遍历行
			for (int i = 0 ; i < copyLenth ; i++) {
				rowFrom = fromSheet.getRow( fromRowNum + i );
				if ( rowFrom == null ) {
					continue;
				}
				rowTo = toSheet.getRow( toRowNum + i );
				if ( rowTo == null ) {
					rowTo = toSheet.createRow( toRowNum + i );
				}
				// 遍历列
				loopLenth = rowFrom.getLastCellNum();
				for (int j = 0 ; j <= loopLenth - fromColNum ; j++) {
					cellFrom = rowFrom.getCell( fromColNum + j );
					if ( cellFrom == null ) {
						continue;
					}
					cellTo = rowTo.getCell( toColNum + j );
					if ( cellTo == null ) {
						cellTo = rowTo.createCell( toColNum + j );
					}
					switch (cellFrom.getCellType()) {
						case Cell.CELL_TYPE_STRING:cellTo.setCellValue( cellFrom.getStringCellValue() );break;
						case Cell.CELL_TYPE_NUMERIC:cellTo.setCellValue( cellFrom.getNumericCellValue() );break;
						case Cell.CELL_TYPE_BOOLEAN:cellTo.setCellValue( cellFrom.getBooleanCellValue() );break;
						case Cell.CELL_TYPE_FORMULA:cellTo.setCellValue( cellFrom.getCellFormula() );break;
						default:cellTo.setCellValue( cellFrom.getStringCellValue() );break;
					}

//					// 单元格样式
//					CellStyle cellStyle = cellFrom.getCellStyle();
//					if ( cellStyle != null ) {
//						cellTo.setCellStyle( cellStyle );
//					}
				}
//				// 复制行高
//				rowTo.setHeight( rowFrom.getHeight() );
//				// 复制行样式
//				CellStyle rowStyle = rowFrom.getRowStyle();
//				if ( rowStyle != null ) {
//					rowTo.setRowStyle( rowStyle );
//				}
			}
		} else {
			// 遍历行
			loopLenth = fromSheet.getLastRowNum();
			for (int i = 0 ; i <= loopLenth - fromRowNum ; i++) {
				rowFrom = fromSheet.getRow( fromRowNum + i );
				if ( rowFrom == null ) {
					continue;
				}
				rowTo = toSheet.getRow( toRowNum + i );
				if ( rowTo == null ) {
					rowTo = toSheet.createRow( toRowNum + i );
				}
				// 遍历列
				for (int j = 0 ; j < copyLenth ; j++) {
					cellFrom = rowFrom.getCell( fromColNum + j );
					if ( cellFrom == null ) {
						continue;
					}
					cellTo = rowTo.getCell( toColNum + j );
					if ( cellTo == null ) {
						cellTo = rowTo.createCell( toColNum + j );
					}
					switch (cellFrom.getCellType()) {
						case Cell.CELL_TYPE_STRING:cellTo.setCellValue( cellFrom.getStringCellValue() );break;
						case Cell.CELL_TYPE_NUMERIC:cellTo.setCellValue( cellFrom.getNumericCellValue() );break;
						case Cell.CELL_TYPE_BOOLEAN:cellTo.setCellValue( cellFrom.getBooleanCellValue() );break;
						case Cell.CELL_TYPE_FORMULA:cellTo.setCellValue( cellFrom.getCellFormula() );break;
						default:cellTo.setCellValue( cellFrom.getStringCellValue() );break;
					}
//					// 单元格样式
//					CellStyle cellStyle = cellFrom.getCellStyle();
//					if ( cellStyle != null ) {
//						cellTo.setCellStyle( cellStyle );
//					}
//					cellTo.setCellValue( cellFrom.getStringCellValue() );
				}
//				// 复制行高
//				rowTo.setHeight( rowFrom.getHeight() );
//				// 复制行样式
//				CellStyle rowStyle = rowFrom.getRowStyle();
//				if ( rowStyle != null ) {
//					rowTo.setRowStyle( rowStyle );
//				}
			}
		}

	}

	// ================================================>单元格操作<================================================
	/**
	 * 在currSheet中rowNum行colNum列添加data和样式cellStyle
	 * @param rowNum	行号
	 * @param colNum	列号
	 * @param data		要写入的数据
	 * @param sheetOpr	目标Sheet页
	 * @param cellStyle	单元格样式
	 */
	public static void addCellData(int rowNum, int colNum, String data, Sheet sheetOpr, CellStyle cellStyle) {
		Row rowTmp;
		Cell cellTmp;
		rowTmp = sheetOpr.getRow( rowNum );
		if ( rowTmp == null ) {
			rowTmp = sheetOpr.createRow( rowNum );
		}
		cellTmp = rowTmp.getCell( colNum );
		if ( cellTmp == null ) {
			cellTmp = rowTmp.createCell( colNum );
		}
		if ( cellStyle != null ) {
			cellTmp.setCellStyle( cellStyle );
		}
		cellTmp.setCellValue( data );
	}
	/**
	 * 在currSheet中rowNum行colNum列添加data和样式cellStyle
	 * @param rowNum	行号
	 * @param colNum	列号
	 * @param data		要写入的数据
	 * @param sheetOpr	目标Sheet页
	 * @param cellStyle	单元格样式
	 */
	public static void addCellData(int rowNum, int colNum, Boolean data, Sheet sheetOpr, CellStyle cellStyle) {
		Row rowTmp;
		Cell cellTmp;
		rowTmp = sheetOpr.getRow( rowNum );
		if ( rowTmp == null ) {
			rowTmp = sheetOpr.createRow( rowNum );
		}
		cellTmp = rowTmp.getCell( colNum );
		if ( cellTmp == null ) {
			cellTmp = rowTmp.createCell( colNum );
		}
		if ( cellStyle != null ) {
			cellTmp.setCellStyle( cellStyle );
		}
		cellTmp.setCellValue( data );
	}
	/**
	 * 在currSheet中rowNum行colNum列添加data和样式cellStyle
	 * @param rowNum	行号
	 * @param colNum	列号
	 * @param data		要写入的数据
	 * @param sheetOpr	目标Sheet页
	 * @param cellStyle	单元格样式
	 */
	public static void addCellData(int rowNum, int colNum, int data, Sheet sheetOpr, CellStyle cellStyle) {
		Row rowTmp;
		Cell cellTmp;
		rowTmp = sheetOpr.getRow( rowNum );
		if ( rowTmp == null ) {
			rowTmp = sheetOpr.createRow( rowNum );
		}
		cellTmp = rowTmp.getCell( colNum );
		if ( cellTmp == null ) {
			cellTmp = rowTmp.createCell( colNum );
		}
		if ( cellStyle != null ) {
			cellTmp.setCellStyle( cellStyle );
		}
		cellTmp.setCellValue( data );
	}

	// ================================================>Excel输出<================================================
	/**
	 * 输出 workbook 到 storePath
	 * @param xlsxWorkbook
	 * @param storePath
	 */
	public static void outputExcel4Xlsx(Workbook xlsxWorkbook, String storePath) {
		FileOutputStream fileOutputStream = null;
		try {
			File file = new File(storePath);
			fileOutputStream = new FileOutputStream( file );
			xlsxWorkbook.write( fileOutputStream );
			// 输出提示
			System.out.println( "\t OUTPUT : OK ---> " + file.getAbsolutePath() );
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if ( fileOutputStream != null ) {
				try {
					fileOutputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	// ================================================>Excel读入<================================================
	/**
	 * 从xlsPath路径下读取Excel文件
	 * @param xlsxPath
	 * @return
	 */
	public static Workbook getExcel4Xlsx(String xlsxPath) {
		if (xlsxPath == null) {
			Workbook outWorkbook = new XSSFWorkbook();
			return outWorkbook;
		}
		FileInputStream fileInputStream = null;
		try {
			fileInputStream = new FileInputStream( xlsxPath );
			Workbook outWorkbook = new XSSFWorkbook( fileInputStream );
			return outWorkbook;
		} catch (Exception e) {
			try {
				fileInputStream = new FileInputStream( xlsxPath );
				Workbook outWorkbook = new HSSFWorkbook( fileInputStream );
				return outWorkbook;
			} catch (Exception e1) {
				e1.printStackTrace();
				e.printStackTrace();
			}
		} finally {
			if ( fileInputStream != null ) {
				try {
					fileInputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return null;
	}

}
