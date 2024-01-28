package com.panxoloto.sharepoint.rest.helper;

import java.util.ArrayList;
import java.util.List;

import org.apache.http.Header;
import org.apache.http.HttpHeaders;
import org.apache.http.message.BasicHeader;

import com.panxoloto.sharepoint.rest.PLGSharepointOnPremisesClient;

public class HeadersOnPremiseHelper {

	PLGSharepointOnPremisesClient client;

	public HeadersOnPremiseHelper(PLGSharepointOnPremisesClient client) {
		this.client = client;
	}

	private void addAcceptJson(List<Header> headers) {
		headers.add(new BasicHeader("Accept", "application/json;odata=verbose"));
	}

	private void addClientHeader(List<Header>  headers) {
		headers.add(new BasicHeader("X-ClientService-ClientTag", "SDK-JAVA"));
	}

	private void addAgentHeader(List<Header>  headers) {
		headers.add(new BasicHeader(HttpHeaders.USER_AGENT, "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"));
	}

	private void addXFormsAuth(List<Header>  headers) {
		headers.add(new BasicHeader("X-FORMS_BASED_AUTH_ACCEPTED", "f"));
	}

	private void addContentType(List<Header>  headers) {
		headers.add(new BasicHeader("Content-Type", "application/json;odata=verbose"));
	}

	private void addDigestKeyHeader(List<Header>  headers) throws Exception {
		headers.add(new BasicHeader("X-RequestDigest", client.getDigestKey()));
	}


	/**
	 * @return
	 */
	public List<Header>  getGetHeaders(boolean includeAuthHeader) throws Exception {
		List<Header>  headers = getCommonHeaders();
		addDigestKeyHeader(headers);
		return headers;
	}

	/**
	 * @param payloadStr
	 * @return
	 */
	public List<Header> getPostHeaders(String payloadStr) throws Exception {
		List<Header>  headers = getCommonHeaders();

		addContentType(headers);
		headers.add(new BasicHeader("Content-length", "" + payloadStr.getBytes().length));
		addDigestKeyHeader(headers);
		return headers;
	}

	/**
	 * @param payloadStr
	 * @return
	 */
	public List<Header> getUpdateHeaders(String payloadStr) throws Exception {
		List<Header>  headers = new ArrayList<>();
		addAcceptJson(headers);
		addContentType(headers);
		addClientHeader(headers);
		addDigestKeyHeader(headers);
		headers.add(new BasicHeader("Content-length", "" + payloadStr.getBytes().length));
		headers.add(new BasicHeader("X-HTTP-Method", "MERGE"));
		headers.add(new BasicHeader("IF-Match", "*"));
	    return headers;
	}
	
	/**
	 * @return
	 */
	public List<Header> getDeleteHeaders() throws Exception {
		List<Header>  headers = new ArrayList<>();
		addAcceptJson(headers);
		addClientHeader(headers);
		addDigestKeyHeader(headers);
		headers.add(new BasicHeader("X-HTTP-Method", "DELETE"));
	    headers.add(new BasicHeader("IF-Match", "*"));
	    return headers;
	}

	public List<Header> getCommonHeaders() {
		List<Header>  headers = new ArrayList<>();
		addAcceptJson(headers);
		addClientHeader(headers);
		addXFormsAuth(headers);
		addAgentHeader(headers);
		return headers;
	}
}
