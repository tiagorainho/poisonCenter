
public class Company {
	private String name;
	private String person;
	private String position;
	private String street;
	private String postalCode;
	private String locality;
	private String country;
	private String phoneNumber;
	private String email;
	
	public Company(String name, String person, String position, String street, String postalCode, String locality, String country, String phoneNumber, String email) {
		this.name = name;
		this.person = person;
		this.position = position;
		this.street = street;
		this.postalCode = postalCode;
		this.locality = locality;
		this.country = country;
		this.phoneNumber = phoneNumber;
		this.email = email;
	}
	
	public String getName() { return this.name; }
	public String getPerson() { return this.person; }
	public String getPosition() { return this.position; }
	public String getStreet() { return this.street; }
	public String getPostalCode() { return this.postalCode; }
	public String getLocality() { return this.locality; }
	public String getCountry() { return this.country; }
	public String getPhoneNumber() { return this.phoneNumber; }
	public String getEmail() { return this.email; }

}
