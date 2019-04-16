package mainPKG.toolUtils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelSheet {
    private CellStyle cellStyle = null;

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    private Sheet sheet;
    private int rowNum = -1;
    private int colNum = -1;
    private int colNumTmp = -1;

    private void updateColNumTmp() {
        colNumTmp = colNum;
    }

    public int getRowNum() {
        return rowNum;
    }

    public int getColNum() {
        return colNum;
    }

    public void reSetRow(int newRow) {
        rowNum = newRow;
        updateColNumTmp();
    }

    public void reSetCol(int newCol) {
        colNum = newCol;
        updateColNumTmp();
    }

    public void reSetRowCol(int newRow, int newCol) {
        rowNum = newRow;
        colNum = newCol;
        updateColNumTmp();
    }

    public Sheet getSheet() {
        return sheet;
    }

    public ExcelSheet(Sheet newSheet, int startRow, int startCol) {
        sheet = newSheet;
        rowNum = startRow;
        colNum = startCol;
        updateColNumTmp();
    }

    public void writeOneLine(Object... items) {
        addOneLine(true, items);
        moveToNextLine();
    }

    public void moveToNextLine() {
        rowNum++;
        updateColNumTmp();
    }

    public void addOneLine(boolean updateColNum, Object... items) {
        for (Object item : items) {
            if (item instanceof String) {
                ExcelUtils.addCellData(rowNum, colNumTmp++, (String) item, sheet, cellStyle);
            } else if (item instanceof Integer) {
                ExcelUtils.addCellData(rowNum, colNumTmp++, (Integer) item, sheet, cellStyle);
            } else if (item instanceof Character) {
                ExcelUtils.addCellData(rowNum, colNumTmp++, "" + (Character) item, sheet, cellStyle);
            } else if (item instanceof Boolean) {
                ExcelUtils.addCellData(rowNum, colNumTmp++, (Boolean) item, sheet, cellStyle);
            } else if (item instanceof List) {
                colNumTmp = writeList(colNumTmp, item);
            } else if (item == null) {
                ExcelUtils.addCellData(rowNum, colNumTmp++, "", sheet, cellStyle);
            } else if (item.getClass().isArray()) {
                colNumTmp = writeArray(colNumTmp, (Object[]) item);
            } else {
                ExcelUtils.addCellData(rowNum, colNumTmp++, "???@" + item.toString(), sheet, cellStyle);
            }
        }
        if (updateColNum) {
            updateColNumTmp();
        }
    }

    private int writeList(int col, Object list) {
        for (Object item : (List) list) {
            if (item instanceof String) {
                ExcelUtils.addCellData(rowNum, col++, (String) item, sheet, cellStyle);
            } else if (item instanceof Integer) {
                ExcelUtils.addCellData(rowNum, col++, (Integer) item, sheet, cellStyle);
            } else if (item instanceof Boolean) {
                ExcelUtils.addCellData(rowNum, col++, (Boolean) item, sheet, cellStyle);
            } else if (item == null) {
                ExcelUtils.addCellData(rowNum, col++, "", sheet, cellStyle);
            } else if (item instanceof List) {
                col = writeList(col, item);
            } else if (item.getClass().isArray()) {
                col = writeArray(col, (Object[]) item);
            } else {
                ExcelUtils.addCellData(rowNum, col++, "[BadType@]toString=" + item.toString(), sheet, cellStyle);
            }
        }
        return col;
    }

    private int writeArray(int col, Object[] list) {
        for (Object item : list) {
            if (item instanceof String) {
                ExcelUtils.addCellData(rowNum, col++, (String) item, sheet, cellStyle);
            } else if (item instanceof Integer) {
                ExcelUtils.addCellData(rowNum, col++, (Integer) item, sheet, cellStyle);
            } else if (item instanceof Boolean) {
                ExcelUtils.addCellData(rowNum, col++, (Boolean) item, sheet, cellStyle);
            } else if (item == null) {
                ExcelUtils.addCellData(rowNum, col++, "", sheet, cellStyle);
            } else if (item instanceof List) {
                col = writeList(col, item);
            } else if (item.getClass().isArray()) {
                col = writeArray(col, (Object[]) item);
            } else {
                ExcelUtils.addCellData(rowNum, col++, "[BadType@]toString=" + item.toString(), sheet, cellStyle);
            }
        }
        return col;
    }

    public Object getValueAt(int rowNum, int colNum, Object defaultValue) {
        try {
            Cell cell = sheet.getRow(rowNum).getCell(colNum);
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    return cell.getHyperlink();
                case Cell.CELL_TYPE_BOOLEAN:
                    return cell.getBooleanCellValue();
                case Cell.CELL_TYPE_ERROR:
                    return cell.getErrorCellValue();
                case Cell.CELL_TYPE_NUMERIC:
                    return cell.getNumericCellValue();
                case Cell.CELL_TYPE_FORMULA:
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue();
            }
            return cell.getStringCellValue();
        } catch (Exception e) {
            return defaultValue;
        }
    }

}
