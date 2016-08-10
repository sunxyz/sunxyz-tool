package cn.sunxyz.common.excel.entity;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import cn.sunxyz.common.excel.annotation.ExcelAttribute;
import cn.sunxyz.common.excel.annotation.ExcelElement;
import cn.sunxyz.common.excel.annotation.ExcelID;


public class School {
	
	@ExcelID
	@ExcelAttribute(name="学校编号",column="A")
	private String id;
	
	@ExcelAttribute(name="学校名称",column="B")
	private String name;
	
	@ExcelElement
	private Set<Clazz> clazzs =  new HashSet<>();
	
	@ExcelElement
	@ExcelAttribute(name="学校描述",column="C")
	private Map<String,String> map = new HashMap<>();

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

	public Set<Clazz> getClazzs() {
		return clazzs;
	}

	public void setClazzs(Set<Clazz> clazzs) {
		this.clazzs = clazzs;
	}

	public Map<String, String> getMap() {
		return map;
	}

	public void setMap(Map<String, String> map) {
		this.map = map;
	}
	

}
