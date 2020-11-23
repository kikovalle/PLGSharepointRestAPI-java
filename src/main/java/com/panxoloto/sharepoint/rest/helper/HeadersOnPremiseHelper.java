package com.panxoloto.sharepoint.rest.helper;

import org.apache.http.HttpHeaders;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;

import com.panxoloto.sharepoint.rest.PLGSharepointOnPremisesClient;

public class HeadersOnPremiseHelper {

	PLGSharepointOnPremisesClient client;

	public HeadersOnPremiseHelper(PLGSharepointOnPremisesClient client) {
		this.client = client;
	}

	private void addAcceptJson(LinkedMultiValueMap<String, String> headers) {
		headers.add("Accept", "application/json;odata=verbose");
	}

	private void addClientHeader(LinkedMultiValueMap<String, String> headers) {
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	}

	private void addAgentHeader(LinkedMultiValueMap<String, String> headers) {
		headers.add(HttpHeaders.USER_AGENT, "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)");
	}

	private void addXFormsAuth(LinkedMultiValueMap<String, String> headers) {
		headers.add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
	}

	private void addContentType(LinkedMultiValueMap<String, String> headers) {
		headers.add("Content-Type", "application/json;odata=verbose");
	}

	private void addDigestKeyHeader(MultiValueMap<String, String> headers) throws Exception {
		headers.add("X-RequestDigest", client.getDigestKey());
	}


	/**
	 * @return
	 */
	public LinkedMultiValueMap<String, String> getGetHeaders(boolean includeAuthHeader) throws Exception {
		LinkedMultiValueMap<String, String> headers = new LinkedMultiValueMap<>();
		addAcceptJson(headers);
		addClientHeader(headers);
		addXFormsAuth(headers);
		addAgentHeader(headers);
		addDigestKeyHeader(headers);
		return headers;
	}


	/**
	 * @param payloadStr
	 * @return
	 */
	public LinkedMultiValueMap<String, String> getPostHeaders(String payloadStr) throws Exception {
		LinkedMultiValueMap<String, String> headers	= new LinkedMultiValueMap<>();
		addAcceptJson(headers);
		addClientHeader(headers);
		addXFormsAuth(headers);
		addAgentHeader(headers);

		addContentType(headers);
		headers.add("Content-length", "" + payloadStr.getBytes().length);
		addDigestKeyHeader(headers);
		return headers;
	}


	/**
	 * @param payloadStr
	 * @return
	 */
	public LinkedMultiValueMap<String, String> getUpdateHeaders(String payloadStr) throws Exception {
		LinkedMultiValueMap<String, String> headers = new LinkedMultiValueMap<>();
		addAcceptJson(headers);
		addContentType(headers);
		headers.add("Content-length", "" + payloadStr.getBytes().length);
		addClientHeader(headers);
		headers.add("X-HTTP-Method", "MERGE");
		headers.add("IF-Match", "*");
		addDigestKeyHeader(headers);
	    return headers;
	}
	
	/**
	 * @return
	 */
	public LinkedMultiValueMap<String, String> getDeleteHeaders() throws Exception {
		LinkedMultiValueMap<String, String> headers = new LinkedMultiValueMap<>();
		addAcceptJson(headers);
		addClientHeader(headers);
		headers.add("X-HTTP-Method", "DELETE");
	    headers.add("IF-Match", "*");
		addDigestKeyHeader(headers);
	    return headers;
	}

	public MultiValueMap<String, String> getCommonHeaders() {
		LinkedMultiValueMap<String, String> headers = new LinkedMultiValueMap<>();
		addAcceptJson(headers);
		addClientHeader(headers);
		addXFormsAuth(headers);
		addAgentHeader(headers);
		return headers;
	}
}
