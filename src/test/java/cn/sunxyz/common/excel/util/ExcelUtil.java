package cn.sunxyz.common.excel.util;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.GenericArrayType;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.lang.reflect.TypeVariable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

import cn.sunxyz.common.excel.annotation.ExcelAttribute;
import cn.sunxyz.common.excel.annotation.ExcelElement;
import cn.sunxyz.common.excel.annotation.ExcelID;
import cn.sunxyz.common.excel.config.ElementTypePath;

/**
 * 
* excel 导入出工具 (8月9日添加导入功能)
* @author 神盾局
* @date 2016年8月6日 下午6:12:53
* @version 1.2
* @param <T>
 */
@Deprecated
public class ExcelUtil<T> {
	
	private static  Logger logger = Logger.getLogger(ExcelUtil.class);
	
	private Class<T> clazz;
	
	//行数
	private int rowNum = 0; 
	
	private boolean flag = true;
	
	// 行需要 判断确定（因此会共享）
	private HSSFRow row;
	
	private ExceMergedRegion exceMergedRegion = new ExceMergedRegion();
	
	public ExcelUtil(Class<T> clazz){
		this.clazz = clazz;
	}
	
	/**
	 * 
	* @Title: importExcel 
	* @Description: TODO(暂时不需要此功能) 
	* @param sheetName
	* @param input
	* @return  List<T> 返回类型  
	* @throws
	 */
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
			
//			HSSFRow title = sheet.getRow(0);
			//获取数据
			
			
			list = dispatch(sheet);
//			for (Map.Entry<String, String> entry : tuples.entrySet()) {
//				logger.debug(entry.getKey()+"=====>"+entry.getValue());
//			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

        return list;  
    }
	
	 /** 
     * 对list数据源将其里面的数据导入到excel表单 
     * @param lists
     * @param sheetName  工作表的名称 
     * @param output java输出流 
     */  
	public boolean exportExcel(List<T> lists[], String sheetNames[], OutputStream output) {
		if (lists.length != sheetNames.length) {
			System.out.println("数组长度不一致");
			return false;
		}

		// 创建excel工作簿
		HSSFWorkbook wb = new HSSFWorkbook();
		// 创建第一个sheet（页），命名为 new sheet
		for (int ii = 0; ii < lists.length; ii++) {
			List<T> list = lists[ii];
			// 产生工作表对象
			HSSFSheet sheet = wb.createSheet();
			wb.setSheetName(ii, sheetNames[ii]);
			// 创建表头
			createHeader(wb, sheet);
			// 写入数据
			int rowStart = 1;
			for (T t : list) {
				logger.debug(statisticsNode(t, 0));
				createRow(t,sheet);
				rowStart = mergedRegio(t,sheet,rowStart);
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
	
	 /** 
     * 对list数据源将其里面的数据导入到excel表单 
     *  
     * @param sheetName 工作表的名称 
     * @param sheetSize 每个sheet中数据的行数,此数值必须小于65536 
     * @param output java输出流 
     */  
    @SuppressWarnings("unchecked")
    public boolean exportExcel(List<T> list, String sheetName,  
            OutputStream output) {
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
	
	/**
	 * 
	* 设置表头
	* @param listField  
	* @throws
	 */
	private void createHeader(HSSFWorkbook wb, HSSFSheet sheet){
		// 产生一行
		HSSFRow row = sheet.createRow(0);
		HSSFCellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
		style.setFillBackgroundColor(HSSFColor.GREY_40_PERCENT.index);
		//得到所有的字段
		List<Field> fields = getAllField(clazz, null);
		HSSFCell cell;// 产生单元格
		for (Field field : fields) {
			ExcelAttribute attr = field  
                    .getAnnotation(ExcelAttribute.class);  
            int col = getExcelCol(attr.column());// 获得列号  
            cell = row.createCell(col);// 创建列  
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);// 设置列中写入内容为String类型  
            cell.setCellValue(attr.name());// 写入列名  

            // 如果设置了提示信息则鼠标放上去提示.  
            if (!attr.prompt().trim().equals("")) {  
                setHSSFPrompt(sheet, "", attr.prompt(), 1, 100, col, col);// 这里默认设了2-101列提示.  
            }  
            // 如果设置了combo属性则本列只能选择不能输入  
            if (attr.combo().length > 0) {  
                setHSSFValidation(sheet, attr.combo(), 1, 100, col, col);// 这里默认设了2-101列只能选择不能输入.  
            }  
            cell.setCellStyle(style);  
		}
	}
	
	
	/**
	 * 
	* 获取所有字段的值
	* TODO （Map暂时只支持普通数据类型 未对普通类型处理--一对一暂不支持）
	* @param t  void 返回类型  
	* @throws
	 */
	private void createRow(Object t,HSSFSheet sheet){
		
		HSSFCell cell;// 产生单元格
		if(flag){
			rowNum++;
			row = sheet.createRow(rowNum);
		}

		Field[] fields = t.getClass().getDeclaredFields();
		for (Field field : fields) {
			try {
				if(!field.isAccessible()){
					// 设置私有属性为可访问
					field.setAccessible(true);
				}
				if(field.isAnnotationPresent(ExcelAttribute.class)&&!field.isAnnotationPresent(ExcelElement.class)){
					ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
					logger.debug(ea.column()+"====>"+field.get(t));
					flag = true;
					try {  
                        // 根据ExcelVOAttribute中设置情况决定是否导出,有些情况需要保持为空,希望用户填写这一列.  
                        if (ea.isExport()) {  
                            cell = row.createCell(getExcelCol(ea.column()));// 创建cell  
                            cell.setCellType(HSSFCell.CELL_TYPE_STRING);  
                            cell.setCellValue(field.get(t) == null ? ""  
                                    : String.valueOf(field.get(t)));// 如果数据存在就填入,不存在填入空格.  
                        }  
                    } catch (IllegalArgumentException e) {  
                        e.printStackTrace();  
                    } catch (IllegalAccessException e) {  
                        e.printStackTrace();  
                    }  
				}else if(field.isAnnotationPresent(ExcelElement.class)){
					flag = false;
					logger.debug("泛型：=====>"+getClass(field.getGenericType(),0));
					switch (ElementTypePath.getElementTypePath(field.getType().getTypeName())) {
					case SET:
						Set<?> set = (Set<?>)field.get(t);
						if(set!=null){
							for (Object object : set) {
								createRow(object,sheet);
							}
						}
						break;
					case LIST:
						List<?> list = (List<?>)field.get(t);
						if(list!=null){
							for (Object object : list) {
								createRow(object,sheet);
							}
						}
						break;
					case MAP:
						
						ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
						Map<?,?> map = (Map<?,?>)field.get(t);
						if(map!=null){
							StringBuffer strB = new StringBuffer();
							for (Map.Entry<?, ?> entry : map.entrySet()) {
								strB.append(entry.getKey()+" : "+entry.getValue()+" , ");
							}
							if(strB.length()>0){
								strB.deleteCharAt(strB.length()-1);
								strB.deleteCharAt(strB.length()-1);
							}
							
							try {  
		                        // 根据ExcelVOAttribute中设置情况决定是否导出,有些情况需要保持为空,希望用户填写这一列.  
		                        if (ea.isExport()) {  
		                            cell = row.createCell(getExcelCol(ea.column()));// 创建cell  
		                            cell.setCellType(HSSFCell.CELL_TYPE_STRING);  
		                            cell.setCellValue(strB== null ? ""  
		                                    : strB.toString());// 如果数据存在就填入,不存在填入空格.  
		                        }  
		                    } catch (IllegalArgumentException e) {  
		                        e.printStackTrace();  
		                    }  
						}
						break;
					default:
						createRow(field.get(t),sheet);
						break;
					}
				}
			} catch (IllegalArgumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		flag = true;
	}
	
	
	
	/**
	 * 
	* 合并数据
	* @param t
	* @param sheet  void 返回类型  
	* @throws
	 */
	private int mergedRegio(Object t,HSSFSheet sheet,int rowStart){
		//获取子节点的数目
		Field[] fields = t.getClass().getDeclaredFields();
		int rowEnd = rowStart + statisticsNode(t, 0) - 1;
		
		for (Field field : fields) {
			if (field.isAnnotationPresent(ExcelAttribute.class)){
				
//				logger.debug(rowStart+"-----------------"+rowEnd);
				ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
				//合并县区名称单元格
		        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowStart, rowEnd, getExcelCol(ea.column()),getExcelCol(ea.column()));
		        sheet.addMergedRegion(cellRangeAddress);				
			}else if(field.isAnnotationPresent(ExcelElement.class)){
				if(!field.isAccessible()){
					field.setAccessible(true);
				}
				try {
					switch (ElementTypePath.getElementTypePath(field.getType().getTypeName())) {
					case SET:
						Set<?> set = (Set<?>)field.get(t);
						if(set!=null){
							for (Object object : set) {
								rowStart = mergedRegio(object, sheet, rowStart);
							}
						}
						break;
					case LIST:
						List<?> list = (List<?>)field.get(t);
						if(list!=null){
							for (Object object : list) {
								rowStart = mergedRegio(object, sheet, rowStart);
							}
						}
						break;
					case MAP:
						break;
					default:
						rowStart = mergedRegio(field.get(t), sheet, rowStart);
						break;
					}
				} catch (IllegalArgumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

			}
		}
		return rowEnd+1;
		
	}
	
	
	
	
	
	/**
	 * 
	* 创建实例对象 并装配关系
	* @param clazz 当前实例对象的类
	* @param row 当前实例对象的行
	* @param idCol 当前实例对象的id列
	* @param parentIdCol 父实例对象的id列
	* @param tuples 存放所有的值
	* @param instanceMap  装配后的对象 
	* @throws
	 */
	@SuppressWarnings("unchecked")
	private Map<String,Map<String, Object>>  createInstance(Class<?> clazz, int row, int idCol, int parentIdCol, int childIdCol,Map<String, String> tuples, Map<String,Map<String, Object>> instanceMap) {
		if(!tuples.containsKey(row + "," + idCol)){
			return instanceMap;
		}
		// TODO判断 是否存在该实例
		if(instanceMap==null){
			instanceMap = new HashMap<>();
		}
		Object entity = null;
		try {
			entity = clazz.newInstance();
		} catch (InstantiationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//同一类对象的map 以 id_parentId作为键
		Map<String,Object> colMap;
		if(instanceMap.containsKey(idCol+"")){
			colMap = instanceMap.get(idCol+"");
		}else{
			colMap = new HashMap<>();
			instanceMap.put(idCol+"", colMap);
		}
		//判断是否有父级对象
		String id = tuples.get(row + "," + idCol);;
		if(parentIdCol>-1){
			String parentId = tuples.get(row + "," + parentIdCol);
			colMap.put(parentId+"_"+id, entity);
			
		}else{
			colMap.put(id, entity);
		}
		logger.debug("colMap===========>"+colMap.size());
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			if (!field.isAccessible()) {
				field.setAccessible(true);
			}
			//对普通属性进行设值
			if(field.isAnnotationPresent(ExcelAttribute.class)&&!field.isAnnotationPresent(ExcelElement.class)){
				ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
				String value = tuples.get(row + "," + getExcelCol(ea.column()));
				Class<?> fieldType = field.getType();
				try {
					if (String.class == fieldType) {
						field.set(entity, String.valueOf(value));
					} else if ((Integer.TYPE == fieldType) || (Integer.class == fieldType)) {
						field.set(entity, Integer.parseInt(value));
					} else if ((Long.TYPE == fieldType) || (Long.class == fieldType)) {
						field.set(entity, Long.valueOf(value));
					} else if ((Float.TYPE == fieldType) || (Float.class == fieldType)) {
						field.set(entity, Float.valueOf(value));
					} else if ((Short.TYPE == fieldType) || (Short.class == fieldType)) {
						field.set(entity, Short.valueOf(value));
					} else if ((Double.TYPE == fieldType) || (Double.class == fieldType)) {
						field.set(entity, Double.valueOf(value));
					} else if (Character.TYPE == fieldType) {
						if ((value != null) && (value.length() > 0)) {
							field.set(entity, Character.valueOf(value.charAt(0)));
						}
					}
				} catch (NumberFormatException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}else if(field.isAnnotationPresent(ExcelElement.class)){
				Map<String,Object> map = instanceMap.get(childIdCol+"");
				List<Object> entitys = new ArrayList<>();
				for (Map.Entry<String,Object> entry : map.entrySet()) {
					String key = entry.getKey();
					if(key.indexOf("_")>-1){
						String[] str = key.split("_");
						String childParentId = str[0];
						if(childParentId.equals(id)){
							entitys.add(entry.getValue());
						}
					}
				}
				//TODO 进行类注入
				String typeName = field.getType().getName();
				try {

					logger.debug("entitys:"+entitys.size());
					switch (ElementTypePath.getElementTypePath(typeName)) {
					case SET:
						Set<Object> set = (Set<Object>)field.get(entity);
						if(set==null){
							set = new HashSet<>();
						}
						for (Object object : entitys) {
							set.add(object);
						}
						break;
					case LIST:
						List<Object> list = (List<Object>)field.get(entity);
						if(list==null){
							list = entitys;
						}else{
							for (Object object : entitys) {
								list.add(object);
							}
						}
						break;
					case MAP:
						if(field.isAnnotationPresent(ExcelAttribute.class)){
							if(getClass(field.getGenericType(),0).getName().equals("java.lang.String")){
								Map<String,String> imap = (Map<String,String>)field.get(entitys);
								ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
								String value = tuples.get(row + "," + getExcelCol(ea.column()));
								if(value.indexOf(",")>-1){
									String[] str = value.split(",");
									for (String string : str) {
										if(string.indexOf(":")>-1){
											String[] keyAndVlaue = string.split(",");
											if(keyAndVlaue.length==2){
												imap.put(keyAndVlaue[0], keyAndVlaue[1]);
											}
										}
									}
								}
							}
							
							
						}
						break;
					default:
						if(entitys.size()==1){
							field.set(entity, entitys.get(0));
						}
						break;
						
					}
				} catch (IllegalArgumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}	
		}
		return instanceMap;
	}
	
	//调度者
	@SuppressWarnings("unchecked")
	private List<T> dispatch(HSSFSheet sheet){
		int rows = sheet.getPhysicalNumberOfRows();
		List<Integer> idCols = getIdCols(clazz,null);
		List<Class<?>> clazzs = getAllClass(clazz,null);
		Map<String,String> tuples = getTuple(sheet,rows);
		if(idCols.size()!=clazzs.size()){
			logger.error("class 数目不一致");
			return null;
		}
		Map<String,Map<String, Object>> instanceMap = null;
		int size = idCols.size();
		for(int i=size-1;i>-1;i--){
			//默认起始行为1
			for(int j=1;j<rows;j++){
				int row = j;
				int childIdCol = -1;
				int idCol = idCols.get(i);
				int parentIdCol = -1; 
				if(i>0){
					parentIdCol = idCols.get(i-1);	
				}
				if(i<size-1){
					childIdCol = idCols.get(i+1);	
				}
				logger.debug(clazzs.get(i)+"==="+row+"===="+idCol+"======"+parentIdCol);
				instanceMap = createInstance(clazzs.get(i), row, idCol, parentIdCol, childIdCol, tuples, instanceMap);
			}
			
		}
		List<T> list = new ArrayList<T>();  
		Map<String,Object> map = instanceMap.get(idCols.get(0)+"");
		for (Map.Entry<String,Object>  entry : map.entrySet()) {
			list.add((T)entry.getValue());
		}
		return list;
	}
	
	/**
	 * 
	* 获取所有标识列（用来获取父级id）
	* @param clazz
	* @param idCols
	* @return  List<Integer> 返回类型  
	* @throws
	 */
	private List<Integer> getIdCols(Class<?> clazz,List<Integer> idCols){
		if(idCols==null){
			idCols = new ArrayList<>();
		}
		Field[] fields= clazz.getDeclaredFields();
		for (Field field : fields) {
			if(field.isAnnotationPresent(ExcelID.class)&&field.isAnnotationPresent(ExcelAttribute.class)){
				ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
				idCols.add(getExcelCol(ea.column()));
			}
		}
		for (Field field : fields) {
			//TODO 此处集合需要做判断
			if(field.isAnnotationPresent(ExcelElement.class)){
				clazz = getClass(field.getGenericType(),0);
//				logger.debug(clazz);
				getIdCols(clazz,idCols);
			}
		}
		return idCols;
	}
	
	private List<Class<?>> getAllClass(Class<?> clazz,List<Class<?>> clazzs){
		if(clazzs==null){
			clazzs = new ArrayList<>();
		}
		clazzs.add(clazz);
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			//TODO 此处集合需要做判断
			if(field.isAnnotationPresent(ExcelElement.class)){
				clazz = getClass(field.getGenericType(),0);
				getAllClass(clazz, clazzs);
			}
		}
		return clazzs;
	}
	
	/**
	 * 
	* 获取 excel中的数据 元组
	* @param sheet
	* @return  List<List<String>> excel元组 
	* @throws
	 */
	private Map<String,String> getTuple(HSSFSheet sheet,int rows){
		
		Map<String,String> tuples = new HashMap<>();
		//获取列
		List<Field> fields = getAllField(clazz, null);
		
		// 从第2行开始取数据,默认第一行是表头.
		for (int i = 1; i < rows; i++) {

			for (Field field : fields) {
				if(field.isAnnotationPresent(ExcelAttribute.class)){
					ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
					int col = getExcelCol(ea.column());
					if(exceMergedRegion.isMergedRegion(sheet, i, col)){
						// 以(行,列）作为一个元组
						String key = i+","+col;
						String value = exceMergedRegion.getMergedRegionValue(sheet, i, col);
//						logger.debug(key+" ===>"+value);  
						tuples.put(key, value);
					}
				}
			}
			
		}
		return tuples;
	}
	
	/**
	 * 
	* 统计子节点的叶子数目
	* @Description: TODO(一个类只支持包含一个集合) 
	* @param t
	* @param childNodeNum
	* @return  int 子节点的数目 
	* @throws
	 */
	private int statisticsNode(Object t, int childNodeNum){
		Field[] fields = t.getClass().getDeclaredFields();
		boolean childNodeFlag = true;
		for (Field field : fields) {
			if(field.isAnnotationPresent(ExcelElement.class)){
				childNodeFlag = false;
				if(!field.isAccessible()){
					field.setAccessible(true);
				}
				try {
					switch (ElementTypePath.getElementTypePath(field.getType().getTypeName())) {
					case SET:
						Set<?> set = (Set<?>)field.get(t);
						if(set!=null){
							for (Object object : set) {
								childNodeNum = statisticsNode(object,childNodeNum);
							}
						}else{
							childNodeNum++;
						}
						break;
					case LIST:
						List<?> list = (List<?>)field.get(t);
						if(list!=null){
							for (Object object : list) {
								childNodeNum = statisticsNode(object,childNodeNum);
							}
						}else{
							childNodeNum++;
						}
						break;
					case MAP:
						childNodeNum++;
						break;
					default:
						childNodeNum = statisticsNode(t,childNodeNum);
						break;
					}
				} catch (IllegalArgumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		if(childNodeFlag){
			childNodeNum++;
		}
		return childNodeNum;
	}

	
	/**
	 * 
	* 获取所有的字段
	* @param clazz
	* @param listField
	* @return  List<Field> 所有字段的集合 
	* @throws
	 */
	private List<Field> getAllField(Class<?> clazz, List<Field> listField) {
		if (listField == null) {
			listField = new ArrayList<>();
		}
		// 获取所有属性
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			Type fieldType = field.getType();

			if (field.isAnnotationPresent(ExcelAttribute.class)) {
				listField.add(field);
				// 类名,属性名
			} else if (field.isAnnotationPresent(ExcelElement.class)) {
				/**
				 * TODO 类型判断
				 */
				switch (ElementTypePath.getElementTypePath(fieldType.getTypeName())) {
				case SET:
				case LIST:
					Type genericFieldType = field.getGenericType();
					getAllField(getClass(genericFieldType, 0), listField);
					break;
				case MAP:
					listField.add(field);
					break;
				default:
					getAllField(field.getClass(),null);
					break;
				}
			}
		}

		return listField;
	}
	
 
	/**
	 * 
	* 得到泛型类对象
	* @param type
	* @param i
	* @return  Class 返回类型  
	* @throws
	 */
	@SuppressWarnings("rawtypes")
	private static Class getClass(Type type, int i) {     
        if (type instanceof ParameterizedType) { // 处理泛型类型     
            return getGenericClass((ParameterizedType) type, i);     
        } else if (type instanceof TypeVariable) {     
            return (Class) getClass(((TypeVariable) type).getBounds()[0], 0); // 处理泛型擦拭对象     
        } else {// class本身也是type，强制转型     
            return (Class) type;     
        }     
    }     
    
    @SuppressWarnings("rawtypes")
	private static Class getGenericClass(ParameterizedType parameterizedType, int i) {     
        Object genericClass = parameterizedType.getActualTypeArguments()[i];     
        if (genericClass instanceof ParameterizedType) { // 处理多级泛型     
            return (Class) ((ParameterizedType) genericClass).getRawType();     
        } else if (genericClass instanceof GenericArrayType) { // 处理数组泛型     
            return (Class) ((GenericArrayType) genericClass).getGenericComponentType();     
        } else if (genericClass instanceof TypeVariable) { // 处理泛型擦拭对象     
            return (Class) getClass(((TypeVariable) genericClass).getBounds()[0], 0);     
        } else {     
            return (Class) genericClass;     
        }     
    } 
    
    
    
    
    /** 
     * 将EXCEL中A,B,C,D,E列映射成0,1,2,3 
     *  
     * @param col 
     */  
    public static int getExcelCol(String col) {  
        col = col.toUpperCase();  
        // 从-1开始计算,字母重1开始运算。这种总数下来算数正好相同。  
        int count = -1;  
        char[] cs = col.toCharArray();  
        for (int i = 0; i < cs.length; i++) {  
            count += (cs[i] - 64) * Math.pow(26, cs.length - 1 - i);  
        }  
        return count;  
    }  

	/**
	 * 设置单元格上提示
	 * 
	 * @param sheet  要设置的sheet.
	 * @param promptTitle  标题
	 * @param promptContent 内容
	 * @param firstRow 开始行
	 * @param endRow 结束行
	 * @param firstCol 开始列
	 * @param endCol  结束列
	 * @return 设置好的sheet.
	 */
	public static HSSFSheet setHSSFPrompt(HSSFSheet sheet, String promptTitle, String promptContent, int firstRow,
			int endRow, int firstCol, int endCol) {
		// 构造constraint对象
		DVConstraint constraint = DVConstraint.createCustomFormulaConstraint("DD1");
		// 四个参数分别是：起始行、终止行、起始列、终止列
		CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
		// 数据有效性对象
		HSSFDataValidation data_validation_view = new HSSFDataValidation(regions, constraint);
		data_validation_view.createPromptBox(promptTitle, promptContent);
		sheet.addValidationData(data_validation_view);
		return sheet;
	}

    /** 
     * 设置某些列的值只能输入预制的数据,显示下拉框. 
     *  
     * @param sheet 要设置的sheet. 
     * @param textlist 下拉框显示的内容 
     * @param firstRow  开始行 
     * @param endRow 结束行 
     * @param firstCol 开始列 
     * @param endCol  结束列 
     * @return 设置好的sheet. 
     */  
    public static HSSFSheet setHSSFValidation(HSSFSheet sheet,  
            String[] textlist, int firstRow, int endRow, int firstCol,  
            int endCol) {  
        // 加载下拉列表内容  
        DVConstraint constraint = DVConstraint  
                .createExplicitListConstraint(textlist);  
        // 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列  
        CellRangeAddressList regions = new CellRangeAddressList(firstRow,  
                endRow, firstCol, endCol);  
        // 数据有效性对象  
        HSSFDataValidation data_validation_list = new HSSFDataValidation(  
                regions, constraint);  
        sheet.addValidationData(data_validation_list);  
        return sheet;  
    } 


}
