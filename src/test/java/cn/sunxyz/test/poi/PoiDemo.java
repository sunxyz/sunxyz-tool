package cn.sunxyz.test.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;


public class PoiDemo {
	
	private Workbook workbook;  
    
    
    
    
    
    /*  
     * 由于Excel当中的单元格Cell存在类型,若获取类型错误 就会产生错误,   
     * 所以通过此方法将Cell内容全部转换为String类型  
     */ 
    private String getCellValue(Cell cell) {  
        String str = null;  
        switch (cell.getCellType()) {  
        case Cell.CELL_TYPE_BLANK:  
            str = "";  
            break;  
        case Cell.CELL_TYPE_BOOLEAN:  
            str = String.valueOf(cell.getBooleanCellValue());  
            break;  
        case Cell.CELL_TYPE_FORMULA:  
            str = String.valueOf(cell.getCellFormula());  
            break;  
        case Cell.CELL_TYPE_NUMERIC:  
            str = String.valueOf(cell.getNumericCellValue());  
            break;  
        case Cell.CELL_TYPE_STRING:  
            str = String.valueOf(cell.getStringCellValue());  
            break;  
        default:  
            str = null;  
            break;  
        }  
        return str;  
    }  
      
 
    @Before 
    public void setUp() throws InvalidFormatException, IOException {  
        // 加载excel文件,自动判断是HSSF还是XSSF  
        workbook = WorkbookFactory.create(new File("E:/aaa.xls"));  
    }  
 
    /*  
     * 读取一个已存在的Excel  
     */ 
    @Test 
    public void testReadExcel() throws InvalidFormatException, IOException {  
 
        // 获取第一个工作目录,下标从0开始  
        Sheet sheet = workbook.getSheetAt(0);  
 
        // 获取该工作目录最后一行的行数  
        int lastRowNum = sheet.getLastRowNum();  
 
        for (int i = 0; i < lastRowNum; i++) {  
 
            // 获取下标为i的行  
            Row row = sheet.getRow(i);  
 
            // 获取该行单元格个数  
            int lastCellNum = row.getLastCellNum();  
 
            for (int j = 0; j < lastCellNum; j++) {  
 
                // 获取下标为j的单元格  
                Cell cell = row.getCell(j);  
 
                // 调用获取方法  
                String cellValue = this.getCellValue(cell);  
            }  
        }  
    }  
 
    /*  
     * 使用Foreach方式读取Excel  
     */ 
    @Test 
    public void testForeachReadExcel() {  
        // 根据sheet的名字获取  
        Sheet sheet = workbook.getSheet("test");  
 
        // 处了上面testReadExcel的方式读取以外,还支持foreach的方式读取  
        for (Row row : sheet) {  
            for (Cell cell : row) {  
                String cellValue = this.getCellValue(cell);  
                System.out.println(cellValue);  
            }  
        }  
    }  
 
    /*  
     * 创建简单的Excel  
     */ 
    @Test 
    public void testWriteExcel() throws IOException {  
        // 创建一个XSSF的Excel文件  
        workbook = new XSSFWorkbook();  
        FileOutputStream fos = new FileOutputStream("E:/test.xlsx");  
 
        // 创建名称为test的工作目录  
        Sheet sheet = workbook.createSheet("test");  
 
        /*  
         * 创建1个10行x10列的工作目录  
         */ 
        for (int i = 0; i < 10; i++) {  
            // 创建一行  
            Row row = sheet.createRow(i);  
            for (int j = 0; j < 10; j++) {  
                // 创建一个单元格  
                Cell cell = row.createCell(j);  
                // 设置单元格value  
                cell.setCellValue("test");  
 
                // 此处为设置Excel的样式,设置单元格内容居中,  
                // 但这样设置方式并不常用,请留意下面的方法  
                CellStyle cs = workbook.createCellStyle();  
                cs.setAlignment(CellStyle.ALIGN_CENTER);  
                cell.setCellStyle(cs);  
 
            }  
        }  
 
        // 将Excel写出到文件流  
        workbook.write(fos);  
    }  
 
    /*  
     * 通过使用模板生成Excel文件,模板当中包含样式,   
     * 这样我们只为模板填充数据就可以有相应的样式  
     */ 
    @Test 
    public void testWriteExcelByTemplate() throws InvalidFormatException,  
            IOException {  
        String fileName = "test.xlsx";  
 
        // 通过类加载器获取模板  
        workbook = WorkbookFactory.create(this.getClass().getClassLoader()  
                .getResourceAsStream(fileName));  
        FileOutputStream fos = new FileOutputStream("E:/test.xlsx");  
 
        Sheet sheet = workbook.getSheetAt(0);  
        Row row = sheet.getRow(0);  
        Cell cell = row.getCell(0);  
        /*  
         * 此时可以通过getCellStyle()来获取到该单元格对象的样式,  
         * 获取到样式只要将此样式放入新创建Excel单元格中,  
         * 就可以完成样式的替换 获取可以直接填充此模板再进行输出,  
         * 注意插入新一行时,要使用sheet.shiftRows(0, 7, 1, true, true);   
         * 这里代表从第0行到第7向下移动1行,保持宽度和高度  
         */ 
        CellStyle cellStyle = cell.getCellStyle();  
 
        workbook.write(fos);  
    }  
}  

