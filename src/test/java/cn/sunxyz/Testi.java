package cn.sunxyz;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.junit.Test;

public class Testi {
	
	@Test
	public void teest(){
		List<Integer> list = new ArrayList<>();
		list.add(5);
		list.add(7);
		list.add(3);
		int i = 1;
		Collections.sort(list, (o1,o2)->o1.compareTo(o2));
		for (Integer integer : list) {
			System.out.println(integer);
		}
	}

}
