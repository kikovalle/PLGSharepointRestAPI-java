package com.panxoloto.sharepoint.rest.helper;

/**
 * Enum with permission ids of a sharepoint site.
 * 
 * @author kikovalle
 *
 */
public enum Permission {

	 Full_Control("1073741829"),
	 Design("1073741828"),
	 Edit("1073741830"),
	 Contribute("1073741827"),	
	 Read("1073741826"),
	 View_Only("1073741924");

	private final String codigo;

	Permission(String string) {
		this.codigo = string;
	}

	public String toString() {
		return codigo;
	}
}
