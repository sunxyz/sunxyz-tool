package cn.sunxyz.common.excel.core;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSON;

import cn.sunxyz.common.excel.annotation.ExcelAttribute;
import cn.sunxyz.common.excel.annotation.ExcelElement;
import cn.sunxyz.common.excel.annotation.ExcelID;
import cn.sunxyz.common.excel.config.ElementTypePath;

public class ExcelUtils<T> extends AbstractExcelUtilss<T>{
	
	private Logger logger = LoggerFactory.getLogger(ExcelUtils.class);
	
	private Class<T> clazz;
	
	private HSSFRow row;
	
	private boolean flag;
	
	{
		flag = true;
	}
	
	@Override
	public IExcelUtil<T> build(Class<T> clazz) {
		this.clazz = clazz;
		return this;
	}
	
	@Override
	public void createHeader(HSSFWorkbook wb, HSSFSheet sheet) {
		// 产生一行
		HSSFRow row = ExcelTool.createRow(sheet);
		HSSFCellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
		style.setFillBackgroundColor(HSSFColor.GREY_40_PERCENT.index);
		// 得到所有的字段
		List<Field> fields = getAllField(clazz, null);
		HSSFCell cell;// 产生单元格
		for (Field field : fields) {
			ExcelAttribute attr = field.getAnnotation(ExcelAttribute.class);
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

	@Override
	public void createRow(Object t, HSSFSheet sheet) {
		HSSFCell cell;// 产生单元格
		if (flag) {
			row = ExcelTool.createRow(sheet);
		}

		Field[] fields = t.getClass().getDeclaredFields();
		try {
			for (Field field : fields) {

				if (!field.isAccessible()) {
					// 设置私有属性为可访问
					field.setAccessible(true);
				}
				if (field.isAnnotationPresent(ExcelAttribute.class) && field.isAnnotationPresent(ExcelElement.class)) {
					flag = false;
//					logger.debug("泛型：=====>" + getClass(field.getGenericType(), 0));
					switch (ElementTypePath.getElementTypePath(field.getType().getTypeName())) {
					case MAP:
						ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
						Map<?, ?> map = (Map<?, ?>) field.get(t);
						if (map != null) {
							StringBuffer strB = new StringBuffer();
							for (Map.Entry<?, ?> entry : map.entrySet()) {
								strB.append(entry.getKey() + " : " + entry.getValue() + " , ");
							}
							if (strB.length() > 0) {
								strB.deleteCharAt(strB.length() - 1);
								strB.deleteCharAt(strB.length() - 1);
							}

							try {
								// 根据ExcelVOAttribute中设置情况决定是否导出,有些情况需要保持为空,希望用户填写这一列.
								if (ea.isExport()) {
									cell = row.createCell(getExcelCol(ea.column()));// 创建cell
									cell.setCellType(HSSFCell.CELL_TYPE_STRING);
									cell.setCellValue(strB == null ? "" : strB.toString());// 如果数据存在就填入,不存在填入空格.
								}
							} catch (IllegalArgumentException e) {
								e.printStackTrace();
							}
						}
						break;
					default:
						break;
					}
				}
			}
			for (Field field : fields) {
				if (!field.isAccessible()) {
					// 设置私有属性为可访问
					field.setAccessible(true);
				}
				if (field.isAnnotationPresent(ExcelAttribute.class) && !field.isAnnotationPresent(ExcelElement.class)) {
					ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
					logger.debug("当前行:"+ea.column() + "====>value:" + field.get(t));
					flag = true;
					try {
						// 根据ExcelVOAttribute中设置情况决定是否导出,有些情况需要保持为空,希望用户填写这一列.
						if (ea.isExport()) {
							cell = row.createCell(getExcelCol(ea.column()));// 创建cell
							cell.setCellType(HSSFCell.CELL_TYPE_STRING);
							cell.setCellValue(field.get(t) == null ? "" : String.valueOf(field.get(t)));// 如果数据存在就填入,不存在填入空格.
						}
					} catch (IllegalArgumentException e) {
						e.printStackTrace();
					} catch (IllegalAccessException e) {
						e.printStackTrace();
					}
				} else if (field.isAnnotationPresent(ExcelElement.class)) {
					flag = false;
//					logger.debug("泛型：=====>" + getClass(field.getGenericType(), 0));
					switch (ElementTypePath.getElementTypePath(field.getType().getTypeName())) {
					case SET:
						Set<?> set = (Set<?>) field.get(t);
						if (set != null) {
							for (Object object : set) {
								createRow(object, sheet);
							}
						}
						break;
					case LIST:
						List<?> list = (List<?>) field.get(t);
						if (list != null) {
							for (Object object : list) {
								createRow(object, sheet);
							}
						}
						break;
					case MAP:
						break;
					default:
						createRow(field.get(t), sheet);
						break;
					}
				}

			}
		} catch (SecurityException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		flag = true;
	}

	@SuppressWarnings("unchecked")
	@Override
	public List<T> dispatch(HSSFSheet sheet) {
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
//				logger.debug(clazzs.get(i)+"==="+row+"===="+idCol+"======"+parentIdCol);
				instanceMap = createInstance(clazzs.get(i), row, idCol, parentIdCol, childIdCol, tuples, instanceMap);
			}
			
		}
		List<T> list = new ArrayList<T>();  
		logger.debug(JSON.toJSONString(instanceMap));
		Map<String,Object> map = instanceMap.get(idCols.get(0)+"");
		for (Map.Entry<String,Object>  entry : map.entrySet()) {
			list.add((T)entry.getValue());
		}
		return list;
	}

	/**
	 * 合并数据
	 */
	public int mergedRegio(Object t,HSSFSheet sheet,int rowStart){
		//获取子节点的数目
		Field[] fields = t.getClass().getDeclaredFields();
		int rowEnd = rowStart + childNodes(t, 0) - 1;
//		logger.debug(rowStart+"====>"+rowEnd);
		for (Field field : fields) {
			if(field.isAnnotationPresent(ExcelAttribute.class)){
				ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
				CellRangeAddress cellRangeAddress = new CellRangeAddress(rowStart, rowEnd, getExcelCol(ea.column()),getExcelCol(ea.column()));
				sheet.addMergedRegion(cellRangeAddress);	
			}else if(field.isAnnotationPresent(ExcelElement.class)&&!field.isAnnotationPresent(ExcelAttribute.class)){
				if(!field.isAccessible()){
					field.setAccessible(true);
				}
				int childRowStart = rowStart;
				try {
					switch (ElementTypePath.getElementTypePath(field.getType().getName())) {
					case SET:
						Set<?> set = (Set<?>)field.get(t);
						if (set != null) {
							for (Object object : set) {
								childRowStart = mergedRegio(object,sheet,childRowStart);
							}
						}
						break;
					case LIST:
						List<?> list = (List<?>)field.get(t);
						if (list != null) {
							for (Object object : list) {
								childRowStart = mergedRegio(object,sheet,childRowStart);
							}
						}
						break;
					case MAP:
						break;
					default:
						childRowStart = mergedRegio(field.get(t),sheet,childRowStart);
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
	* 统计子节点的叶子数目
	* @Description: TODO(一个类只支持包含一个集合) 
	* @param t
	* @param childNodeNum
	* @return  int 子节点的数目 
	* @throws
	 */
	private int childNodes(Object t, int childNodeNum){
		Field[] fields =t.getClass().getDeclaredFields();
		boolean childNodeFlag = true;
		for (Field field : fields) {
			if(field.isAnnotationPresent(ExcelElement.class)&&!field.isAnnotationPresent(ExcelAttribute.class)){
				if(!field.isAccessible()){
					field.setAccessible(true);
				}
				childNodeFlag = false;
				try {
					switch (ElementTypePath.getElementTypePath(field.getType().getName())) {
					case SET:
						Set<?> set = (Set<?>)field.get(t);
						if (set != null) {
							if(set.size()==0){
								childNodeFlag = true;
							}else{
								for (Object object : set) {
									childNodeNum = childNodes(object, childNodeNum);
								}
							}
							
						}else{
							childNodeFlag = true;
						}
						break;
					case LIST:
						List<?> list = (List<?>)field.get(t);
						if (list != null) {
							if(list.size()==0){
								childNodeFlag = true;
							}else{
								for (Object object : list) {
									childNodeNum = childNodes(object, childNodeNum);
								}
							}
						}else{
							childNodeFlag = true;
						}
						break;
					case MAP:
						break;
					default:
						childNodeNum = childNodes(field.get(t), childNodeNum);
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
//		logger.debug("colMap===========>"+colMap.size());
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
				if(map!=null){
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
				}

				//TODO 进行类注入
				String typeName = field.getType().getName();
				try {

//					logger.debug("entitys:"+entitys.size());
					switch (ElementTypePath.getElementTypePath(typeName)) {
					case SET:
						Set<Object> set = (Set<Object>)field.get(entity);
						if(set==null){
							set = new HashSet<>();
							field.set(entity, set);
						}
						for (Object object : entitys) {
							set.add(object);
						}
						break;
					case LIST:
						List<Object> list = (List<Object>)field.get(entity);
						if(list==null){
							list = entitys;
							field.set(entity, list);
						}else{
							for (Object object : entitys) {
								list.add(object);
							}
						}
						break;
					case MAP:
						if(field.isAnnotationPresent(ExcelAttribute.class)){
							if(getClass(field.getGenericType(),0).getName().equals("java.lang.String")){
								Map<String,String> imap = (Map<String,String>)field.get(entity);
								if(imap==null){
									imap = new HashMap<>();
									field.set(entity, imap);
								}
								ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
								String value = tuples.get(row + "," + getExcelCol(ea.column()));
								if(value.indexOf(",")>-1){
									String[] str = value.split(",");
									for (String string : str) {
										if(string.indexOf(":")>-1){
											String[] keyAndVlaue = string.split(":");
											if(keyAndVlaue.length==2){
												imap.put(keyAndVlaue[0].trim(), keyAndVlaue[1].trim());
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
			if(field.isAnnotationPresent(ExcelElement.class)&&!field.isAnnotationPresent(ExcelAttribute.class)){
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
					if(ExcelTool.isMergedRegion(sheet, i, col)){
						// 以(行,列）作为一个元组
						String key = i+","+col;
						logger.debug(key);
						String value = ExcelTool.getMergedRegionValue(sheet, i, col);
//						logger.debug(key+" ===>"+value);  
						tuples.put(key, value);
					}
				}
			}
			
		}
		return tuples;
	}
}


