package nonused;

import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map.Entry;
import java.util.Set;

public class Case2 {

	public void readMyExcelData(LinkedHashMap<String,LinkedHashMap<String,String>> tabData)  {
		
		Set<String> idSet = tabData.keySet();
		Iterator<String> it = idSet.iterator();
		while(it.hasNext()) {
			String emplyee = it.next();
			//System.out.println(it.next());
			LinkedHashMap<String, String> tData = tabData.get(emplyee);
			Set<Entry<String,String>> colData = tData.entrySet();
			Iterator<Entry<String,String>> colItr = colData.iterator();
			System.out.println("Details for empid : "+emplyee);
			while(colItr.hasNext())  {
				Entry<String, String> colMap = colItr.next();
				System.out.print(colMap.getKey()+"\t"+colMap.getValue()+"\n");
			}
		}
		
	}

}
