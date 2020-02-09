import java.util.LinkedList;
import java.util.List;

public class Product {
	private String code;
	private String preparation;
	private String description;
	private List<Ingredient> ingredients;
	
	public Product(String code, String preparation, String description) {
		this.code = code;
		this.preparation = preparation;
		this.description = description;
		this.ingredients = new LinkedList<Ingredient>();
	}
	
	public void addIngredient(Ingredient ingredient) {
		this.ingredients.add(ingredient);
	}
	
	public String getName() {
		return this.description;
	}
	
	public String getCode() {
		return this.code;
	}
	
	public String getReference() {
		return this.preparation;
	}
	
	public List<Ingredient> getIngredients(){
		return this.ingredients;
	}

}
