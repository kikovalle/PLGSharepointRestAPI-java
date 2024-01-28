package com.panxoloto.sharepoint.rest.helper;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.http.Header;
import org.apache.http.message.BasicHeader;

public class HeadersHelper {

	private AuthTokenHelperOnline tokenHelper;

	public HeadersHelper(AuthTokenHelperOnline tokenHelper) {
		this.tokenHelper = tokenHelper;
	}
	
	/**
	 * @return
	 */
	public List<Header> getGetHeaders(boolean includeAuthHeader) {
		List<Header> headers = new ArrayList<>();
		headers.add(new BasicHeader("Cookie", this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) ));
		headers.add(new BasicHeader("Accept", "application/json;odata=verbose"));
		headers.add(new BasicHeader("X-ClientService-ClientTag", "SDK-JAVA"));
	    if (includeAuthHeader || tokenHelper.isUseClientId()) {
	    	headers.add(new BasicHeader("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue()));
	    } else {
	    	headers.add(new BasicHeader("X-RequestDigest", this.tokenHelper.getFormDigestValue()));
	    }
	    return headers;
	}
	
	/**
	 * @param payloadStr
	 * @return
	 */
	public List<Header> getPostHeaders(String payloadStr) {
		List<Header> headers = new ArrayList<>();
		headers.add(new BasicHeader("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";"))) );
		headers.add(new BasicHeader("Accept", "application/json;odata=verbose"));
		headers.add(new BasicHeader("Content-Type", "application/json;odata=verbose"));
		headers.add(new BasicHeader("Content-length", "" + payloadStr.getBytes().length));
		headers.add(new BasicHeader("X-ClientService-ClientTag", "SDK-JAVA"));
	    headers.add(new BasicHeader("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue()));
	    return headers;
	}
	
	/**
	 * @param payloadStr
	 * @return
	 */
	public List<Header> getUpdateHeaders(String payloadStr) {
		List<Header> headers = new ArrayList<>();
		headers.add(new BasicHeader("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";"))) );
		headers.add(new BasicHeader("Accept", "application/json;odata=verbose"));
		headers.add(new BasicHeader("Content-Type", "application/json;odata=verbose"));
		headers.add(new BasicHeader("Content-length", "" + payloadStr.getBytes().length));
		headers.add(new BasicHeader("X-ClientService-ClientTag", "SDK-JAVA"));
		headers.add(new BasicHeader("X-HTTP-Method", "MERGE"));
		headers.add(new BasicHeader("IF-Match", "*"));
	    headers.add(new BasicHeader("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue()));
	    return headers;
	}
	
	/**
	 * @return
	 */
	public List<Header> getDeleteHeaders() {
		List<Header> headers = new ArrayList<>();
		headers.add(new BasicHeader("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";"))) );
		headers.add(new BasicHeader("Accept", "application/json;odata=verbose"));
		headers.add(new BasicHeader("X-ClientService-ClientTag", "SDK-JAVA"));
		headers.add(new BasicHeader("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue()));
		headers.add(new BasicHeader("X-HTTP-Method", "DELETE"));
		headers.add(new BasicHeader("IF-Match", "*"));
	    return headers;
	}
	
}
