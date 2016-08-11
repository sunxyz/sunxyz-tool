# sunxyz-tool
**一些常用的工具类库**

**poi excel-tool**

这是一个简单的poi excel 导入导出工具库 他可以帮你完成导入导出级联关系的操作

首先需要导入pom.xml 依赖

```
<!-- 为POI支持Office Open XML -->
	<dependency>
		<groupId>org.apache.poi</groupId>
		<artifactId>poi-ooxml</artifactId>
		<version>3.9</version>
	</dependency>
	<dependency>
		<groupId>org.apache.poi</groupId>
		<artifactId>poi-ooxml-schemas</artifactId>
		<version>3.9</version>
	</dependency> 
```
然后引入 cn.sunxyz.common.excel下的包

在实体类上标注相关注解 test中已经给出一个简单的例子

[使用介绍：](http://blog.csdn.net/zhugeyangyang1994/article/details/52184742)

导出示例 

```
List<School> list = new ArrayList<>();
FileOutputStream output = null;  
try {  
	output = new FileOutputStream("d:\\success3.xls");  
} catch (FileNotFoundException e) {  
    e.printStackTrace();  
}  
IExcelUtil<School> eu = new ExcelUtils<>();
eu.build(School.class).exportExcel(list, "学校信息", output);
```

导入示例

```
FileInputStream fis = null;  
try {  
    fis = new FileInputStream("d:\\success3.xls");  
    IExcelUtil<School> util = new ExcelUtils<>();//创建excel工具类  
    List<School> list = util.build(School.class).importExcel("学校信息", fis);// 导入  
    logger.info(JSON.toJSONString(list));  
} catch (FileNotFoundException e) {  
    e.printStackTrace();  
}
```
