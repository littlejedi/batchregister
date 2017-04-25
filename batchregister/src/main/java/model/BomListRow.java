package model;

public class BomListRow {
	
	private String productSerialId;
	
	private String productName;
	
	private String materialSerialId;
	
	private String materialName;
	
	private String materialQuantity;
	
	public BomListRow(String productSerialId, String productName, String materialSerialId, String materialName, String materialQuantity) {
		this.productSerialId = productSerialId;
		this.productName = productName;
		this.materialSerialId = materialSerialId;
		this.materialName = materialName;
		this.materialQuantity = materialQuantity;
	}

	public String getProductSerialId() {
		return productSerialId;
	}

	public void setProductSerialId(String productSerialId) {
		this.productSerialId = productSerialId;
	}

	public String getProductName() {
		return productName;
	}

	public void setProductName(String productName) {
		this.productName = productName;
	}

	public String getMaterialSerialId() {
		return materialSerialId;
	}

	public void setMaterialSerialId(String materialSerialId) {
		this.materialSerialId = materialSerialId;
	}

	public String getMaterialName() {
		return materialName;
	}

	public void setMaterialName(String materialName) {
		this.materialName = materialName;
	}

	public String getMaterialQuantity() {
		return materialQuantity;
	}

	public void setMaterialQuantity(String materialQuantity) {
		this.materialQuantity = materialQuantity;
	}
	
}