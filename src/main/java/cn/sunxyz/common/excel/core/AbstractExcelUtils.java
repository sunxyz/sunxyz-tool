package cn.sunxyz.common.excel.core;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 
 * 常用 抽象方法实现
 * 
 * @author 神盾局
 * @date 2016年8月9日 下午2:33:17
 * 
 * @param <T>
 */
public abstract class AbstractExcelUtils<T> extends AbstractExcelUtil<T> {



	@Override
	public List<T> importExcel(String sheetName, InputStream input) {
		List<T> list = null;
        try {
			HSSFWorkbook workbook = new HSSFWorkbook(input);  
			HSSFSheet sheet = workbook.getSheet(sheetName);  
			if (!sheetName.trim().equals("")) {  
			    sheet = workbook.getSheet(sheetName);// 如果指定sheet名,则取指定sheet中的内容.  
			}  
			if (sheet == null) {  
			    sheet = workbook.getSheetAt(0); // 如果传入的sheet名不存在则默认指向第1个sheet.  
			}		
			list = dispatch(sheet);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

        return list;  
	}

	@SuppressWarnings("unchecked")
	@Override
	public boolean exportExcel(List<T> list, String sheetName, OutputStream output) {
		//此处 对类型进行转换
        List<T> ilist = new ArrayList<>();
        for (T t : list) {
            ilist.add(t);
        }
        List<T>[] lists = new ArrayList[1];  
        lists[0] = ilist;  

        String[] sheetNames = new String[1];  
        sheetNames[0] = sheetName;  
        return exportExcel(lists, sheetNames, output);  
	}

	@Override
	public boolean exportExcel(List<T>[] lists, String[] sheetNames, OutputStream output) {
		if (lists.length != sheetNames.length) {
			System.out.println("数组长度不一致");
			return false;
		}

		// 创建excel工作簿
		HSSFWorkbook wb = ExcelTool.createWorkbook();
		// 创建第一个sheet（页），命名为 new sheet
		for (int ii = 0; ii < lists.length; ii++) {
			List<T> list = lists[ii];
			// 产生工作表对象			
			HSSFSheet sheet = ExcelTool.createSheet(ii, sheetNames[ii]);
			// 创建表头
			createHeader(wb, sheet);
			// 写入数据
			int rowStart = 1;
			for (T t : list) {
				createRow(t, sheet);
				rowStart = mergedRegio(t, sheet, rowStart);
			}
			
		}
		try {
			output.flush();
			wb.write(output);
			output.close();
			return true;
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("Output is closed ");
			return false;
		}

	}
	
}
