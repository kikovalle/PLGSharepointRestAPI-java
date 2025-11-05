package com.panxoloto.sharepoint.rest.helper;

import java.io.InputStream;
import java.io.StringReader;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;
import java.util.concurrent.ExecutionException;
import java.util.function.Supplier;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.http.Header;
import org.apache.http.HeaderElement;
import org.apache.http.HttpEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.entity.mime.content.StringBody;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.protocol.HttpContext;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.azure.core.credential.AccessToken;
import com.azure.core.credential.TokenRequestContext;
import com.azure.identity.ClientCertificateCredential;
import com.azure.identity.ClientCertificateCredentialBuilder;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.google.gson.stream.JsonReader;

public class CloudTokenForClientIdGetter {

	private final Logger LOG = LoggerFactory.getLogger(this.getClass());

	private String clientId;
	private String clientSecret;
	private String siteURL;
	private Supplier<HttpClientBuilder> httpClientBuilderSupplier;

	private String certificatePath;
	private String certificatePassword;
	private String clientTenant;
	private String sharepointScope;

	public CloudTokenForClientIdGetter(String clientId, String clientSecret, String siteURL,
									   Supplier<HttpClientBuilder> httpClientBuilderSupplier,
									   String certificatePath,
									   String certificatePassword,
									   String clientTenant,
									   String sharepointScope) {
		this(clientId, clientSecret, siteURL, HttpClients::custom);
		this.certificatePath = certificatePath;
		this.certificatePassword = certificatePassword;
		this.clientTenant = clientTenant;
		this.sharepointScope = sharepointScope;
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
	private LocalDateTime spOnlineTokenExpiration = LocalDateTime.now().minusMinutes(1);

	public String getToken() {
		try {
			if (spOnlineRealm == null) {
				getTenantId();
			}
			if (spOnlineToken == null || LocalDateTime.now().isAfter(spOnlineTokenExpiration)) {
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
		get.setHeader(org.apache.http.HttpHeaders.AUTHORIZATION, "Bearer");
		CloseableHttpClient httpClient = httpClientBuilderSupplier.get().build();

		try (CloseableHttpResponse response = httpClient.execute(get, (HttpContext) null)) {
			Header[] headers = response.getHeaders("WWW-Authenticate");
			for (Header h : headers) {
				HeaderElement[] elements = h.getElements();
				for (HeaderElement e : elements) {
					if ("Bearer realm".equals(e.getName())) {
						spOnlineRealm = e.getValue();
					}
					if ("client_id".equals(e.getName())) {
						spOnlineClientId = e.getValue();
					}
				}
			}
		}
	}

	private void getBearerToken() throws Exception {

		AccessToken spToken = getSharepointAccessToken();

		if (spToken != null) {
			spOnlineToken = spToken.getToken();
			spOnlineTokenType = spToken.getTokenType();
			spOnlineTokenExpiration = spToken.getExpiresAt().atZoneSameInstant(ZoneId.systemDefault()).toLocalDateTime().minusMinutes(1);
		} else {
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
				spOnlineTokenExpiration = LocalDateTime.now().plusSeconds(reply.get("expires_in").getAsLong()).minusMinutes(1);
			}
		}

		LOG.debug("got SPonline token", "token: " + spOnlineToken.substring(0, 15) + "..., expiration: " + spOnlineTokenExpiration);
	}

	private HttpEntity fillInSPOnlineTokenRequestData() throws MalformedURLException {
		String clientId = this.clientId + "@" + spOnlineRealm;
		URL url = new URL(siteURL);
		String resourceStr = spOnlineClientId + "/" + url.getHost() + "@" + spOnlineRealm;

		MultipartEntityBuilder builder = MultipartEntityBuilder.create();
		builder.setMode(HttpMultipartMode.BROWSER_COMPATIBLE);
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


	private AccessToken getAccessToken(ClientCertificateCredential clientCertificateCredential, String scope) throws ExecutionException, InterruptedException {
		TokenRequestContext request = new TokenRequestContext().addScopes(scope);
		AccessToken token = clientCertificateCredential.getToken(request).toFuture().get();
		return token;
	}

	private ClientCertificateCredential getClientCertCredential(String spClientId, String spTenantId, String spCertPrivateKey, String certPrivateKeyPath) {
		return new ClientCertificateCredentialBuilder()
				.clientId(spClientId)
				.tenantId(spTenantId)
				.pfxCertificate(certPrivateKeyPath, spCertPrivateKey).build();
	}


	private AccessToken getSharepointAccessToken() throws Exception {
		if (StringUtils.isBlank(certificatePath)) {
			return null;
		}
		URL certPrivateKeyPathUrl = new URL("file://" + certificatePath);

		if(certPrivateKeyPathUrl != null) {
				ClientCertificateCredential certificateCredential
						= getClientCertCredential(clientId, clientTenant, certificatePassword, certPrivateKeyPathUrl.getPath());
				return getAccessToken(certificateCredential, sharepointScope);
		}
		return null;
	}

}
