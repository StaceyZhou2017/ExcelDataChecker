package nonused;

import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map.Entry;
import java.util.Set;

public class Case51 {
	
	public void readData(LinkedHashMap<String, LinkedHashMap<String,String>> tData)  {
		
		Set set = tData.keySet();
		Iterator it = set.iterator();
		while(it.hasNext())  {
			Object obj = it.next();
			System.out.println("Employee id:"+obj);
			LinkedHashMap<String, String> td = tData.get(obj);
			Set<Entry<String,String>> colData = td.entrySet();
			Iterator itr = colData.iterator();
			while(itr.hasNext())  {
				Entry<String,String> ob = (Entry<String, String>) itr.next();
				System.out.println(ob.getKey()+"\t"+ob.getValue()+"\n");
			}
			
		}
		
	}

}
