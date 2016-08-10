package cn.sunxyz.common.excel.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 
 * 用于普通类型字段
 * @author 神盾局
 * @date 2016年8月5日 上午9:46:04
 *
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelAttribute {

	/**
	 * 导出到Excel中的名字.
	 */
	String name();

	/**
	 * 配置列的名称,对应A,B,C,D....
	 */
	String column();

	/**
	 * 提示信息
	 */
	String prompt() default "";

	/**
	 * 设置只能选择不能输入的列内容.
	 */
	String[] combo() default {};

	/**
	 * 是否导出数据,应对需求:有时我们需要导出一份模板,这是标题需要但内容需要用户手工填写.
	 */
	boolean isExport() default true;

}