package entity;

/**
 *
 * @author Dinhisme
 */
public class Product {

    String idProduct, product, type, brand;

    public Product() {
    }

    public Product(String idProduct, String product, String type, String brand) {
        this.idProduct = idProduct;
        this.product = product;
        this.type = type;
        this.brand = brand;
    }

    public String getIdProduct() {
        return idProduct;
    }

    public void setIdProduct(String idProduct) {
        this.idProduct = idProduct;
    }

    public String getProduct() {
        return product;
    }

    public void setProduct(String product) {
        this.product = product;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getBrand() {
        return brand;
    }

    public void setBrand(String brand) {
        this.brand = brand;
    }

}
