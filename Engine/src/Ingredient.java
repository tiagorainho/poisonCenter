
public class Ingredient {
	private String casNumber;
	private String productCode;
	private String description;
	private String percent;
	
	public Ingredient(String casNumber, String productCode, String description, String percent) {
		this.casNumber = casNumber;
		this.productCode = productCode;
		this.description = description;
		this.percent = percent;
	}
	
	public Ingredient() {}
	
	public String getCasNumber() {return this.casNumber;}
	public String getProductCode() {return this.productCode;}
	public String getName() {return this.description;}
	public String getPercent() {return this.percent;}
	
	public void setCasNumber(String casNumber) {this.casNumber = casNumber;}
	public void setProductCode(String productCode) {this.productCode = productCode;}
	public void setDescription(String description) {this.description = description;}
	public void setPercent(String percent) {this.percent = percent;}
}
