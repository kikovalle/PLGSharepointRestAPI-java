package com.panxoloto.sharepoint.rest.helper;

import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.List;
import java.util.function.Supplier;
import java.util.stream.Collectors;

import javax.xml.transform.TransformerException;

import org.apache.http.HttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.util.EntityUtils;
import org.json.JSONException;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class AuthTokenHelperOnPremises {

	private static final Logger LOG = LoggerFactory.getLogger(AuthTokenHelperOnPremises.class);
	private String spDomain;
	private String spSitePrefix;
	private HttpProtocols protocol = HttpProtocols.HTTPS;

	/**
	 * Helper class to manage login against SharepointOnline and retrieve auth token and cookies to
	 * perform calls to rest API.
	 * Retrieves all info needed to set auth headers to call Sharepoint Rest API v1.
	 * 
	 * @param spSitePrefix
	 * @param spDomain
	 */
	public AuthTokenHelperOnPremises(String spSitePrefix, String spDomain) {
		super();
		this.spDomain = spDomain;
		this.spSitePrefix = spSitePrefix;
	}

	public void setProtocol(HttpProtocols protocol) {
		this.protocol = protocol;
	}

	public HttpProtocols getProtocol() {
		return protocol;
	}

	private String getProtocolString() {
		return protocol == HttpProtocols.HTTPS ? "https" : "http";
	}

	public String getFormDigestValue(List<String> cookies, Supplier<HttpClientBuilder> httpClientBuilderSupplier)
			throws IOException, URISyntaxException, TransformerException, JSONException {
		try (CloseableHttpClient client = httpClientBuilderSupplier.get().build()) {
			HttpPost post = new HttpPost(String.format("https://%s/_api/contextinfo", this.spDomain));
			post.addHeader("Cookie",  cookies.stream().collect(Collectors.joining(";")));
			post.addHeader("Accept", "application/json;odata=verbose");
			post.addHeader("X-ClientService-ClientTag", "SDK-JAVA");
			post.setEntity(new StringEntity(""));
			HttpResponse response = client.execute(post);
			JSONObject json = new JSONObject(EntityUtils.toString(response.getEntity()));
			
			return json.getJSONObject("d").getJSONObject("GetContextWebInformation").getString("FormDigestValue");
		}
	}
	
	/**
	 * Mounts the sharepoint online site url, composed by the protocol, domain and spSiteUri.
	 * 
	 * @return
	 * @throws URISyntaxException 
	 */
	public URI getSharepointSiteUrl(String apiPath) throws URISyntaxException {
		LOG.debug("getSharepointSiteUrl {}", apiPath);
		return new URI(getProtocolString(),
				this.spDomain,
				this.spSitePrefix +  apiPath,
				null
				);
	}
	
	/**
	 * @param apiPath
	 * @param query
	 * @return
	 * @throws URISyntaxException
	 */
	public URI getSharepointSiteUrl(String apiPath, String query) throws URISyntaxException {
		LOG.debug("getSharepointSiteUrl {} {query}", apiPath);
		return new URI(getProtocolString(),
				this.spDomain,
				this.spSitePrefix + apiPath,
				query,
				null
				);
	}
}
