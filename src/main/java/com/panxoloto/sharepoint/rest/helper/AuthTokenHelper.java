package com.panxoloto.sharepoint.rest.helper;

import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Supplier;
import java.util.stream.Collectors;

import javax.xml.transform.TransformerException;

import org.apache.http.Header;
import org.apache.http.HttpResponse;
import org.apache.http.client.ClientProtocolException;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.util.EntityUtils;
import org.json.JSONException;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class AuthTokenHelper {

	private static final Logger LOG = LoggerFactory.getLogger(AuthTokenHelper.class);
	private String spSiteUri;
	private String formDigestValue ;
	private String domain;
	private Supplier<HttpClientBuilder> httpClientBuilderSupplier;
	private List<String> cookies;
	private final String TOKEN_LOGIN_URL = "https://login.microsoftonline.com/extSTS.srf";
	private String payload = "<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"\n"
			+ "      xmlns:a=\"http://www.w3.org/2005/08/addressing\"\n"
			+ "      xmlns:u=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd\">\n"
			+ "  <s:Header>\n"
			+ "    <a:Action s:mustUnderstand=\"1\">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>\n"
			+ "    <a:ReplyTo>\n" 
			+ "      <a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address>\n"
			+ "    </a:ReplyTo>\n"
			+ "    <a:To s:mustUnderstand=\"1\">https://login.microsoftonline.com/extSTS.srf</a:To>\n"
			+ "    <o:Security s:mustUnderstand=\"1\"\n"
			+ "       xmlns:o=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd\">\n"
			+ "      <o:UsernameToken>\n" 
			+ "        <o:Username>%s</o:Username>\n"
			+ "        <o:Password>%s</o:Password>\n" 
			+ "      </o:UsernameToken>\n" 
			+ "    </o:Security>\n"
			+ "  </s:Header>\n" 
			+ "  <s:Body>\n"
			+ "    <t:RequestSecurityToken xmlns:t=\"http://schemas.xmlsoap.org/ws/2005/02/trust\">\n"
			+ "      <wsp:AppliesTo xmlns:wsp=\"http://schemas.xmlsoap.org/ws/2004/09/policy\">\n"
			+ "        <a:EndpointReference>\n" 
			+ "          <a:Address>%s</a:Address>\n"
			+ "        </a:EndpointReference>\n" 
			+ "      </wsp:AppliesTo>\n"
			+ "      <t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>\n"
			+ "      <t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>\n"
			+ "      <t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType>\n"
			+ "    </t:RequestSecurityToken>\n" 
			+ "  </s:Body>\n" 
			+ "</s:Envelope>";

	/**
	 * Helper class to manage login against SharepointOnline and retrieve auth token and cookies to
	 * perform calls to rest API.
	 * Retrieves all info needed to set auth headers to call Sharepoint Rest API v1.
	 * 
	 * @param httpClientBuilderSupplier
	 * @param user
	 * @param passwd
	 * @param domain
	 * @param spSiteUri
	 */
	public AuthTokenHelper(Supplier<HttpClientBuilder> httpClientBuilderSupplier, String user, String passwd, String domain, String spSiteUri) {
		super();
		this.domain = domain;
		this.spSiteUri = spSiteUri;
		this.payload = String.format(this.payload, user, passwd, domain);
	}
	
	
	protected String receiveSecurityToken() throws URISyntaxException, ClientProtocolException, IOException {
		CloseableHttpClient client = httpClientBuilderSupplier.get().build();
		HttpPost post = new HttpPost(TOKEN_LOGIN_URL);
		post.setEntity(new StringEntity(payload));
		HttpResponse response = client.execute(post);

		String securityToken = EntityUtils.toString(response.getEntity());
		String clave1 = "<wsse:BinarySecurityToken";
		String clave2 = "</wsse:BinarySecurityToken>";
		securityToken = securityToken.substring(securityToken.indexOf(clave1));
		securityToken = securityToken.substring(securityToken.indexOf(">") + 1);
		securityToken = securityToken.substring(0, securityToken.indexOf(clave2));
		return securityToken;
	}

	protected List<String> getSignInCookies(String securityToken)
			throws TransformerException, URISyntaxException, Exception {
		CloseableHttpClient client = httpClientBuilderSupplier.get().build();
		HttpPost post = new HttpPost(String.format("https://%s/_forms/default.aspx?wa=wsignin1.0", this.domain));
		post.setEntity(new StringEntity(securityToken));
		HttpResponse response = client.execute(post);
		Header[] headers = response.getHeaders("Set-Cookie");
		
		List<String> cookies = new ArrayList<>();
		for (Header header : headers) {
			cookies.add(header.getValue());
		}
		if (cookies.size() < 1) {
			throw new Exception("Unable to sign in: no cookies returned in response");
		}
		return cookies;
	}

	protected String getFormDigestValue(List<String> cookies)
			throws IOException, URISyntaxException, TransformerException, JSONException {
		CloseableHttpClient client = httpClientBuilderSupplier.get().build();
		HttpPost post = new HttpPost(String.format("https://%s/_api/contextinfo", this.domain));
		post.addHeader("Cookie",  cookies.stream().collect(Collectors.joining(";")));
		post.addHeader("Accept", "application/json;odata=verbose");
		post.addHeader("X-ClientService-ClientTag", "SDK-JAVA");
		post.setEntity(new StringEntity(""));
		HttpResponse response = client.execute(post);
		JSONObject json = new JSONObject(EntityUtils.toString(response.getEntity()));

		return json.getJSONObject("d").getJSONObject("GetContextWebInformation").getString("FormDigestValue");
	}
	
	/**
	 * @throws Exception
	 */
	public void init() throws Exception {
		String securityToken = receiveSecurityToken();
		this.cookies = getSignInCookies(securityToken);
		formDigestValue = getFormDigestValue(this.cookies);
	}


	/**
	 * The security token to use in Authorization Bearer  header or X-RequestDigest header 
	 * (depending on operation called from Rest API).
	 * 
	 * @return
	 */
	public String getFormDigestValue() {
		return formDigestValue;
	}

	/**
	 * Retrieves session cookies to use in communication with the Sharepoint Online Rest API.
	 * 
	 * @return
	 */
	public List<String> getCookies() {
		return this.cookies;
	}
	
	/**
	 * Mounts the sharepoint online site url, composed by the protocol, domain and spSiteUri.
	 * 
	 * @return
	 * @throws URISyntaxException 
	 */
	public URI getSharepointSiteUrl(String apiPath) throws URISyntaxException {
		return new URI("https",
				this.domain,
				this.spSiteUri + apiPath,
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
		if (!query.startsWith("$filter=")) {
			LOG.debug("Missing $filter in query string, adding");
			query = String.format("%s%s", "$filter=", query);
		}
		return new URI("https",
				this.domain,
				this.spSiteUri + apiPath,
				query,
				null
				);
	}
}
