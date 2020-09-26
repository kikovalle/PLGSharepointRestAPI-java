package com.panxoloto.sharepoint.rest.helper;

import java.net.URI;
import java.net.URISyntaxException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class AuthTokenHelperOnPremises {

	private static final Logger LOG = LoggerFactory.getLogger(AuthTokenHelperOnPremises.class);
	private String spDomain;
	private String spSitePrefix;

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
	
	/**
	 * Mounts the sharepoint online site url, composed by the protocol, domain and spSiteUri.
	 * 
	 * @return
	 * @throws URISyntaxException 
	 */
	public URI getSharepointSiteUrl(String apiPath) throws URISyntaxException {
		return new URI("https",
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
		if (!query.startsWith("$filter=")) {
			LOG.debug("Missing $filter in query string, adding");
			query = String.format("%s%s", "$filter=", query);
		}
		return new URI("https",
				this.spDomain,
				this.spSitePrefix + apiPath,
				query,
				null
				);
	}
}
