package com.panxoloto.sharepoint.rest.helper;

import java.io.InputStream;
import java.io.StringReader;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.Date;
import java.util.function.Supplier;

import org.apache.commons.io.IOUtils;
import org.apache.hc.client5.http.classic.methods.HttpGet;
import org.apache.hc.client5.http.classic.methods.HttpPost;
import org.apache.hc.client5.http.entity.mime.HttpMultipartMode;
import org.apache.hc.client5.http.entity.mime.MultipartEntityBuilder;
import org.apache.hc.client5.http.entity.mime.StringBody;
import org.apache.hc.client5.http.impl.classic.CloseableHttpClient;
import org.apache.hc.client5.http.impl.classic.CloseableHttpResponse;
import org.apache.hc.client5.http.impl.classic.HttpClientBuilder;
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.core5.http.ContentType;
import org.apache.hc.core5.http.Header;
import org.apache.hc.core5.http.HttpEntity;
import org.apache.hc.core5.http.HttpHeaders;
import org.apache.hc.core5.http.protocol.HttpContext;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.google.gson.stream.JsonReader;

public class CloudTokenForClientIdGetter {

	private final Logger LOG = LoggerFactory.getLogger(this.getClass());

	private String clientId;
	private String clientSecret;
	private String siteURL;
	private Supplier<HttpClientBuilder> httpClientBuilderSupplier;

	public CloudTokenForClientIdGetter(String clientId, String clientSecret, String siteURL) {
		this(clientId, clientSecret, siteURL, HttpClients::custom);
	}

	public CloudTokenForClientIdGetter(String clientId, String clientSecret, String siteURL, Supplier<HttpClientBuilder> httpClientBuilderSupplier) {
		this.clientId = clientId;
		this.clientSecret = clientSecret;
		this.siteURL = siteURL;
		this.httpClientBuilderSupplier = httpClientBuilderSupplier;
	}

	private String spOnlineRealm = null;
	private String spOnlineClientId;
	private String spOnlineToken = null;
	private String spOnlineTokenType;
	private Date spOnlineTokenExpiration = new Date(0l);

	public String getToken() {
		try {
			if (spOnlineRealm == null) {
				getTenantId();
			}
			if (spOnlineToken == null || new Date().after(spOnlineTokenExpiration)) {
				getBearerToken();
			}
		} catch (Exception e) {
			throw new RuntimeException("can't authenticate to Sharepoint online", e);
		}
		return spOnlineToken;
	}


	private void getTenantId() throws Exception {
		String url = siteURL+"/_vti_bin/client.svc/";
		HttpGet get = new HttpGet(url);
		get.setHeader(HttpHeaders.AUTHORIZATION, "Bearer");
		CloseableHttpClient httpClient = httpClientBuilderSupplier.get().build();

		try (CloseableHttpResponse response = httpClient.execute(get, (HttpContext) null)) {
			Header[] headers = response.getHeaders("WWW-Authenticate");
			for (Header h : headers) {
				if ("Bearer realm".equals(h.getName())) {
					spOnlineRealm = h.getValue();
				}
				if ("client_id".equals(h.getName())) {
					spOnlineClientId = h.getValue();
				}
			}
		}
	}

	private void getBearerToken() throws Exception {
		String url = "https://accounts.accesscontrol.windows.net/" + spOnlineRealm + "/tokens/OAuth/2";
		HttpPost post = new HttpPost(url);
		HttpEntity multipart = fillInSPOnlineTokenRequestData();
		post.setEntity(multipart);

		CloseableHttpClient httpClient = httpClientBuilderSupplier.get().build();
		Date reqDate = new Date();
		try (CloseableHttpResponse response = httpClient.execute(post, (HttpContext) null)) {
			HttpEntity entity = response.getEntity();
			InputStream instream = entity.getContent();
			String json = IOUtils.toString(instream, StandardCharsets.UTF_8.name());

			JsonParser jp = new JsonParser();
			JsonReader jr = new JsonReader(new StringReader(json));
			jr.setLenient(true);
			JsonObject reply = jp.parse(jr).getAsJsonObject();
			spOnlineToken = reply.get("access_token").getAsString();
			spOnlineTokenType = reply.get("token_type").getAsString();
			spOnlineTokenExpiration = new Date(reqDate.getTime() + reply.get("expires_in").getAsLong() * 1000);
		}

		LOG.debug("got SPonline token", "token: " + spOnlineToken.substring(0, 15) + "..., expiration: " + spOnlineTokenExpiration);
	}

	private HttpEntity fillInSPOnlineTokenRequestData() throws MalformedURLException {
		String clientId = this.clientId + "@" + spOnlineRealm;
		URL url = new URL(siteURL);
		String resourceStr = spOnlineClientId + "/" + url.getHost() + "@" + spOnlineRealm;

		MultipartEntityBuilder builder = MultipartEntityBuilder.create();
		builder.setMode(HttpMultipartMode.LEGACY);
		StringBody stringBody = new StringBody("client_credentials", ContentType.MULTIPART_FORM_DATA);
		builder.addPart("grant_type", stringBody);
		stringBody = new StringBody(clientId, ContentType.MULTIPART_FORM_DATA);
		builder.addPart("client_id", stringBody);
		stringBody = new StringBody(clientSecret, ContentType.MULTIPART_FORM_DATA);
		builder.addPart("client_secret", stringBody);
		stringBody = new StringBody(resourceStr, ContentType.MULTIPART_FORM_DATA);
		builder.addPart("resource", stringBody);

		HttpEntity multipart = builder.build();
		return multipart;
	}

}
