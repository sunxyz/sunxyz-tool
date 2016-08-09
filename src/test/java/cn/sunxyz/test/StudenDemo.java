package cn.sunxyz.test;

import org.junit.Test;

public class StudenDemo {
	
	/**
	 * 
	* easy 二分法
	 */
	public int search(int key, int a[]){
		int left = 0;
		int right = a.length-1;
		while(left<=right){
			int mid = (left+right)/2;
			if(key>a[mid]){
				left = mid + 1;
			}else if(key<a[mid]){
				right = mid - 1;
			}else{
				return mid;
			}
		}
		return -1;
	}
	
	@Test
	public void test(){
		int[] a = {1,2,3,5,7,8,9,10,11,15,19};
		System.out.println(search(10, a));
	}
	
	@Test
	public void test2(){
		System.out.println(Math.abs(-2147483648));
	}

}
