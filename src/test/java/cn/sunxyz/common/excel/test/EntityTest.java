package cn.sunxyz.common.excel.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.log4j.Logger;
import org.junit.Test;

import com.alibaba.fastjson.JSON;

import cn.sunxyz.common.excel.entity.Clazz;
import cn.sunxyz.common.excel.entity.School;
import cn.sunxyz.common.excel.entity.Student;
import cn.sunxyz.common.excel.util.ExcelUtil;


public class EntityTest {

	private Logger logger = Logger.getLogger(EntityTest.class);
	
	@SuppressWarnings("deprecation")
	@Test
	public void test(){
		

		Set<Student> students = new HashSet<>();
		Student student = new Student();;
		student.setId("121");
		student.setAge(8);
		student.setName("小明");
		students.add(student);
		
		Student student2 = new Student();;
		student2.setId("122");
		student2.setAge(9);
		student2.setName("小李");
		students.add(student2);
		
		
		Set<Clazz> clazzs = new HashSet<>();
		Clazz clazz = new Clazz();
		clazz.setId("11");
		clazz.setName("一年级");
		clazz.setStudents(students);
		clazzs.add(clazz);
		
		Clazz clazz2 = new Clazz();
		clazz2.setId("12");
		clazz2.setName("二年级");
		clazz2.setStudents(students);
		clazzs.add(clazz2);
		
		Clazz clazz3 = new Clazz();
		clazz3.setId("13");
		clazz3.setName("三年级");
		clazzs.add(clazz3);
		
		Clazz clazz4 = new Clazz();
		clazz4.setId("14");
		clazz4.setName("四年级");
		clazz4.setStudents(students);
		clazzs.add(clazz4);
		
		List<School> list = new ArrayList<>();
		
		School school = new School();
		school.setId("1");
		school.setName("中山");
		school.setClazzs(clazzs);
		list.add(school);
		
		Map<String,String> map = new HashMap<>();
		map.put("1", "红星小学");
		map.put("2", "TOP");
		School school1 = new School();
		school1.setId("2");
		school1.setName("红星");
		school1.setClazzs(clazzs);
		school1.setMap(map);
		list.add(school1);
		 
        FileOutputStream output = null;  
        try {  
        	output = new FileOutputStream("d:\\success3.xls");  
        } catch (FileNotFoundException e) {  
            e.printStackTrace();  
        }  
		new ExcelUtil<School>(School.class).exportExcel(list, "学校信息", output);
	}
	
	@SuppressWarnings("deprecation")
	@Test
	public void importExcel(){
		FileInputStream fis = null;  
        try {  
            fis = new FileInputStream("d:\\success3.xls");  
            ExcelUtil<School> util = new ExcelUtil<>(School.class);// 创建excel工具类  
            List<School> list = util.importExcel("学校信息", fis);// 导入  
            logger.info(JSON.toJSONString(list));  
        } catch (FileNotFoundException e) {  
            e.printStackTrace();  
        }
	}
}
