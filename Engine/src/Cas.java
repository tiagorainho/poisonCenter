
public class Cas {
	private String casNumber;
	private String productCode;
	private String descriptionEN;
	private String descriptionPT;
	
	public Cas(String casNumber, String productCode, String descriptionEN, String descriptionPT) {
		this.casNumber = casNumber;
		this.productCode = productCode;
		this.descriptionEN = descriptionEN;
		this.descriptionPT = descriptionPT;
	}
	
	public String getCasNumber() { return this.casNumber;}
	public String getProductCode() { return this.productCode;}
	public String getDescriptionEN() { return this.descriptionEN;}
	public String getDescriptionPT() { return this.descriptionPT;}
	
	public void setCasNumber(String casNumber) {this.casNumber = casNumber;}
	public void setProductCode(String productCode) {this.productCode = productCode;}
	public void setDescriptionEN(String descriptionEN) {this.descriptionEN = descriptionEN.toLowerCase();}
	public void setDescriptionPT(String descriptionPT) {this.descriptionPT = descriptionPT.toLowerCase();}
}
