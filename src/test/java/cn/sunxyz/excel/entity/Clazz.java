package cn.sunxyz.excel.entity;

import java.util.HashSet;
import java.util.Set;

import cn.sunxyz.common.excel.annotation.ExcelAttribute;
import cn.sunxyz.common.excel.annotation.ExcelElement;
import cn.sunxyz.common.excel.annotation.ExcelID;



public class Clazz{
   
	@ExcelID
	@ExcelAttribute(name="教室编号",column="D")
	private String id;
	
	@ExcelAttribute(name="教室名称",column="E")
	private String name;
	
	@ExcelElement
	private Set<Student> students = new HashSet<>();

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

	public Set<Student> getStudents() {
		return students;
	}

	public void setStudents(Set<Student> students) {
		this.students = students;
	}
	
	
}
