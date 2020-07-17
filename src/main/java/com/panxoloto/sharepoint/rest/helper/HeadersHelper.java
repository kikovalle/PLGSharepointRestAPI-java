package com.panxoloto.sharepoint.rest.helper;

import java.util.stream.Collectors;

import org.springframework.util.LinkedMultiValueMap;

public class HeadersHelper {

	private AuthTokenHelper tokenHelper;

	public HeadersHelper(AuthTokenHelper tokenHelper) {
		this.tokenHelper = tokenHelper;
	}
	
	/**
	 * @return
	 */
	public LinkedMultiValueMap<String, String> getGetHeaders(boolean includeAuthHeader) {
		LinkedMultiValueMap<String, String> headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    if (includeAuthHeader) {
	    	headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    } else {
	    	headers.add("X-RequestDigest", this.tokenHelper.getFormDigestValue());
	    }
	    return headers;
	}
	
	/**
	 * @param payloadStr
	 * @return
	 */
	public LinkedMultiValueMap<String, String> getPostHeaders(String payloadStr) {
		LinkedMultiValueMap<String, String> headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("Content-length", "" + payloadStr.getBytes().length);
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    return headers;
	}
	
	/**
	 * @param payloadStr
	 * @return
	 */
	public LinkedMultiValueMap<String, String> getUpdateHeaders(String payloadStr) {
		LinkedMultiValueMap<String, String> headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("Content-length", "" + payloadStr.getBytes().length);
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
		headers.add("X-HTTP-Method", "MERGE");
		headers.add("IF-Match", "*");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    return headers;
	}
	
	/**
	 * @return
	 */
	public LinkedMultiValueMap<String, String> getDeleteHeaders() {
		LinkedMultiValueMap<String, String> headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    headers.add("X-HTTP-Method", "DELETE");
	    headers.add("IF-Match", "*");
	    return headers;
	}
	
}
