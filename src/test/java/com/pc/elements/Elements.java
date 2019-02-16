/**
 * @ClassPurpose Object repository
 * @Scriptor All
 * @ReviewedBy
 * @ModifiedBy All
 * @LastDateModified 3/17/2017
 */
package com.pc.elements;

import java.util.HashMap;
import org.openqa.selenium.By;

public class Elements
{
		private  HashMap<String,By> hm = new HashMap<String,By>();  
		
		public Elements()
		{	

		}
		
		public By getObject(String ff)
		{
			By retuValue = null;
			if(hm.containsKey(ff))
			{
			  retuValue = hm.get(ff);
			}
			return retuValue;
		}
}