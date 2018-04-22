package com.qa.test;

public class Product {

    private String productName;
	private String productScreenshot;
	    

	    public Product(String productName, String productScreenshot) {
	        this.productName = productName;
	        this.productScreenshot = productScreenshot;
	        
	    }
		public void setproductName(String productName) {
			this.productName = productName;
		}
		
		public String getproductName() {
			return productName;
		}
		
		public void setproductScreenshot(String productScreenshot) {
			this.productScreenshot = productScreenshot;
		}
		public String getproductScreenshot() {
			return productScreenshot;
		}
}
