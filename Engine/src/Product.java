import java.util.LinkedList;
import java.util.List;

public class Product {
	private String code;
	private String preparation;
	private String description;
	private String UFIcodeCalculated;
	private String UFIcodeNotificated;
	private String EuPCS;
	private List<Ingredient> ingredients;
	
	public Product(String code, String preparation, String description, String UFIcodeCalculated, String UFIcodeNotificated, String EuPCS) {
		this.code = code;
		this.preparation = preparation;
		this.description = description;
		this.UFIcodeCalculated = UFIcodeCalculated;
		this.UFIcodeNotificated = UFIcodeNotificated;
		this.EuPCS = EuPCS;
		this.ingredients = new LinkedList<Ingredient>();
	}
	
	public void addIngredient(Ingredient ingredient) { this.ingredients.add(ingredient); }
	public String getName() { return this.description; }
	public String getCode() { return this.code; }
	public String getReference() { return this.preparation; }
	public String getUFIcodeCalculated() { return this.UFIcodeCalculated; }
	public String getUFIcodeNotificated() { return this.UFIcodeNotificated; }
	public String getEuPCS() { return this.EuPCS; }
	public List<Ingredient> getIngredients(){ return this.ingredients; }

}
