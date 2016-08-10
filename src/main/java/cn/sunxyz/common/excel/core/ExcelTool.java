package cn.sunxyz.common.excel.core;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 
* excel 常用功能
* @author 神盾局
* @date 2016年8月9日 下午2:19:05
*
 */
public class ExcelTool {
	
	//当前行数
	private static int num = 0;
	
	private static HSSFWorkbook wb;
	
	private static HSSFCell cell;
	
	private static HSSFRow row;
	
	
	public static HSSFWorkbook createWorkbook(){
		wb = new HSSFWorkbook();
		return wb;
	}
	
	/**
	 * 
	* 创建一个sheet
	* @param sheetname
	* @return  HSSFSheet 返回类型  
	* @throws
	 */
	public static HSSFSheet createSheet(int ii, String sheetname){
		// 创建excel工作簿
		HSSFSheet sheet = wb.createSheet();
		wb.setSheetName(ii, sheetname);
		return sheet;
	}
	
	/**
	 * 
	* 创建一行
	* @param sheet
	* @return  HSSFRow 返回类型  
	* @throws
	 */
	public static HSSFRow createRow(HSSFSheet sheet){
		row = sheet.createRow(num);
		num++;
		return row;
	}
	
	/**
	 * 
	* 创建一格
	* @param row
	* @param col
	* @return  HSSFCell 返回类型  
	* @throws
	 */
	public static HSSFCell createCell(HSSFRow row, int col){
		cell = row.createCell(col);// 创建cell  
		return cell;
	}
	
	/**
	 * 获取合并单元格的值
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	public static String getMergedRegionValue(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();

		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();

			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					Row fRow = sheet.getRow(firstRow);
					Cell fCell = fRow.getCell(firstColumn);

					return getCellValue(fCell);
				}
			}
		}

		return null;
	}

	/**
	 * 判断指定的单元格是否是合并单元格
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	public static boolean isMergedRegion(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();

		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();

			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {

					return true;
				}
			}
		}

		return false;
	}

	/**
	 * 获取单元格的值
	 * 
	 * @param cell
	 * @return
	 */
	public static String getCellValue(Cell cell) {
		if (cell == null)
			return "";
		if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
			return cell.getStringCellValue();

		} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
			return String.valueOf(cell.getBooleanCellValue());

		} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
			return cell.getCellFormula();

		} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			return String.valueOf(cell.getNumericCellValue());
		}

		return "";
	}

	public static int getNum() {
		return num;
	}

	public static void setNum(int num) {
		ExcelTool.num = num;
	}

	public static HSSFWorkbook getWb() {
		return wb;
	}

	public static void setWb(HSSFWorkbook wb) {
		ExcelTool.wb = wb;
	}

	public static HSSFCell getCell() {
		return cell;
	}

	public static void setCell(HSSFCell cell) {
		ExcelTool.cell = cell;
	}

	public static HSSFRow getRow() {
		return row;
	}

	public static void setRow(HSSFRow row) {
		ExcelTool.row = row;
	}

	
	
}
