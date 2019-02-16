/**
 * @ClassPurpose This Class is to create instance for Common and this instance will called across all other classes
 * @Scriptor Banu Pradeep
 * @ReviewedBy
 * @ModifiedBy 
 * @LastDateModified 3/17/2017
 */
package com.pc.utilities;

public class CommonManager 
{	
	   private static CommonManager instance = new CommonManager();
	   ThreadLocal<Common> common = new ThreadLocal<Common>();
	   
	   public static CommonManager getInstance()
	   {
	      return instance;
	   }
				   
	   public Common getCommon() 
	   {
	      return common.get();
	      
	   }
	   
	   public void setCommon(Common com){
		   common.set(com);
	   }

}
