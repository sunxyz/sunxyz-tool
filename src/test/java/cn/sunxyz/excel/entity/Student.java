package cn.sunxyz.excel.entity;

import java.util.Map;

import cn.sunxyz.common.excel.annotation.ExcelAttribute;
import cn.sunxyz.common.excel.annotation.ExcelElement;
import cn.sunxyz.common.excel.annotation.ExcelID;

public class Student {
	
	@ExcelID
	@ExcelAttribute(name="学生编号",column="F")
	private String id;
	
	@ExcelAttribute(name="学生姓名",column="G")
	private String name;
	
	@ExcelAttribute(name="学生年龄",column="H")
	private Integer age;
	
	@ExcelElement
	@ExcelAttribute(name="学生详细信息",column="I")
	private Map<String,String> map;

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Integer getAge() {
		return age;
	}

	public void setAge(Integer age) {
		this.age = age;
	}

	public Map<String, String> getMap() {
		return map;
	}

	public void setMap(Map<String, String> map) {
		this.map = map;
	}
	
	
	
	

}
