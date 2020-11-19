package com.panxoloto.sharepoint.rest;

import java.util.ArrayList;
import java.util.List;

import org.apache.http.auth.AuthScope;
import org.apache.http.auth.NTCredentials;
import org.apache.http.client.CredentialsProvider;
import org.apache.http.impl.client.BasicCredentialsProvider;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.json.JSONArray;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpMethod;
import org.springframework.http.RequestEntity;
import org.springframework.http.ResponseEntity;
import org.springframework.http.client.HttpComponentsClientHttpRequestFactory;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.RestTemplate;

import com.panxoloto.sharepoint.rest.helper.AuthTokenHelperOnPremises;
import com.panxoloto.sharepoint.rest.helper.HeadersOnPremiseHelper;
import com.panxoloto.sharepoint.rest.helper.HttpProtocols;
import com.panxoloto.sharepoint.rest.helper.Permission;

public class PLGSharepointOnPremisesClient implements PLGSharepointClient {


	private static final Logger LOG = LoggerFactory.getLogger(PLGSharepointOnPremisesClient.class);
	private MultiValueMap<String, String> headers;
	private RestTemplate restTemplate;
	private String spSiteUrl;
	private HeadersOnPremiseHelper headerHelper;
	private AuthTokenHelperOnPremises tokenHelper;
	private HttpProtocols protocol = HttpProtocols.HTTPS;


	/**
	 * @param spSiteUr.- The sharepoint site URL like https://contoso.sharepoint.com/sites/contososite
	 */
	/**
	 * @param user - The user email to access sharepoint online site.
	 * @param passwd - the user password to access sharepoint online site.
	 * @param domain - the domain without protocol and no uri like contoso.sharepoint.com
	 * @param spSiteUrl - The sharepoint site URI - host part
	 * @param spSitePrefix - The sharepoint site URI - path part /sites/contososite
	 */
	public PLGSharepointOnPremisesClient(String user, 
			String passwd, String domain, String spSiteUrl, String spSitePrefix) {
		super();
		
		CredentialsProvider credsProvider = new BasicCredentialsProvider();
		credsProvider.setCredentials(AuthScope.ANY, new NTCredentials(user, passwd, spSiteUrl, domain));
		CloseableHttpClient httpClient = HttpClients.custom()
		        .setDefaultCredentialsProvider(credsProvider)
		        .build();
		HttpComponentsClientHttpRequestFactory requestFactory = new HttpComponentsClientHttpRequestFactory();
		requestFactory.setHttpClient(httpClient);
		this.restTemplate = new RestTemplate(requestFactory);

		this.spSiteUrl = spSiteUrl;
		if (this.spSiteUrl.endsWith("/")) {
			LOG.debug("spSiteUri ends with /, removing character");
			this.spSiteUrl = this.spSiteUrl.substring(0, this.spSiteUrl.length() - 1);
		}
		if (!this.spSiteUrl.startsWith("/")) {
			LOG.debug("spSiteUri doesnt start with /, adding character");
			this.spSiteUrl = String.format("%s%s", "/", this.spSiteUrl);
		}
		try {
			LOG.debug("Wrapper auth initialization performed successfully. Now you can perform actions on the site.");
			this.headerHelper = new HeadersOnPremiseHelper();
			this.tokenHelper = new AuthTokenHelperOnPremises(spSitePrefix, spSiteUrl);
		} catch (Exception e) {
			LOG.error("Initialization failed!! Please check the user, pass, domain and spSiteUri parameters you provided", e);
		}
	}

	public HttpProtocols getProtocol() {
		return protocol;
	}

	public void setProtocol(HttpProtocols protocol) {
		this.protocol = protocol;
		tokenHelper.setProtocol(protocol);
	}

	/**
	 * Method to get json string wich you can transform to a JSONObject and get data from it.
	 * 
	 * 
	 * @param data - Data to be sent as query (look at the rest api documentation on how to include search filters).
	 * @return.- String representing a json object if the auth is correct.
	 * @throws Exception
	 */
	@Override
	public JSONObject getAllLists(String data) throws Exception {
		LOG.debug("getAllLists {}", data);
	    headers = headerHelper.getGetHeaders(false);

	    RequestEntity<String> requestEntity = new RequestEntity<>(data, 
	        headers, HttpMethod.GET, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/lists")
	        );

	    ResponseEntity<String> responseEntity = 
	        restTemplate.exchange(requestEntity, String.class);

	    return new JSONObject(responseEntity.getBody());
	}
	
	/**
	 * Retrieves list info by list title.
	 * 
	 * @param title - Site list title to query info.
	 * @param jsonExtendedAttrs - form json body for advanced query (take a look at ms documentation about rest api).
	 * @return json string with list information that can be used to get a JSONObject to use this info.
	 * @throws Exception thrown when something went wrong.
	 */
	@Override
	public JSONObject getListByTitle(String title, String jsonExtendedAttrs) throws Exception {
		LOG.debug("getListByTitle {} jsonExtendedAttrs {}", new Object[] {title, jsonExtendedAttrs});
	    headers = headerHelper.getGetHeaders(false);

	    RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs, 
	        headers, HttpMethod.GET, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + title + "')")
	        );

	    ResponseEntity<String> responseEntity = 
	        restTemplate.exchange(requestEntity, String.class);

	    return new JSONObject(responseEntity.getBody());
	}

	/**
	 * Retrieves list info by list title.
	 * 
	 * @param title - Site list title to query info.
	 * @return json string with list information that can be used to get a JSONObject to use this info.
	 * @throws Exception thrown when something went wrong.
	 */
	@Override
	public JSONObject getListFields(String title) throws Exception {
		LOG.debug("getListByTitle {} ", new Object[] {title});
	    headers = headerHelper.getGetHeaders(false);

	    RequestEntity<String> requestEntity = new RequestEntity<>("{}", 
	        headers, HttpMethod.GET, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + title + "')/Fields")
	        );

	    ResponseEntity<String> responseEntity = 
	        restTemplate.exchange(requestEntity, String.class);

	    return new JSONObject(responseEntity.getBody());
	}


	
	
	/**
	 * @param listTitle
	 * @param description
	 * @return
	 * @throws Exception
	 */
	@Override
	public JSONObject createList(String listTitle, String description) throws Exception {
		LOG.debug("createList siteUrl {} listTitle {} description {}", new Object[] {listTitle, description});
		JSONObject payload = new JSONObject();
		JSONObject meta = new JSONObject();
		meta.put("type", "SP.List");
		payload.put("__metadata", meta);
		payload.put("AllowContentTypes", true);
		payload.put("BaseTemplate", 100);
		payload.put("ContentTypesEnabled", true);
		payload.put("Description", description);
		payload.put("Title", listTitle);
		
		String payloadStr = payload.toString();
		headers = headerHelper.getPostHeaders(payloadStr);
	    
	    RequestEntity<String> requestEntity = new RequestEntity<>(payloadStr, 
    			headers, HttpMethod.POST, 
    			this.tokenHelper.getSharepointSiteUrl("/_api/web/lists")
    			);
	    ResponseEntity<String> responseEntity =  restTemplate.exchange(requestEntity, String.class);
	    return new JSONObject(responseEntity.getBody());
	}

	/**
	 * @param listTitle
	 * @param newDescription
	 * @return
	 * @throws Exception
	 */
	@Override
	public JSONObject updateList(String listTitle, String newDescription) throws Exception {
		LOG.debug("createList siteUrl {} listTitle {} description {}", new Object[] {listTitle, newDescription});
		JSONObject payload = new JSONObject();
		JSONObject meta = new JSONObject();
		meta.put("type", "SP.List");
		payload.put("__metadata", meta);
		if (newDescription != null) {
			payload.put("Description", newDescription);
		}

		String payloadStr = payload.toString();
		headers = headerHelper.getUpdateHeaders(payloadStr);
	    
	    RequestEntity<String> requestEntity = new RequestEntity<>(payloadStr, 
    			headers, HttpMethod.POST, 
    			this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + listTitle + "')")
    			);
	    ResponseEntity<String> responseEntity =  restTemplate.exchange(requestEntity, String.class);
	    return new JSONObject(responseEntity.getBody());
	}

	
	/**
	 * @param title
	 * @param jsonExtendedAttrs
	 * @param filter
	 * @return
	 * @throws Exception
	 */
	@Override
	public JSONObject getListItems(String title, String jsonExtendedAttrs, String filter) throws Exception {
		LOG.debug("getListByTitle {} jsonExtendedAttrs {}", new Object[] {title, jsonExtendedAttrs});
	    headers = headerHelper.getGetHeaders(true);

	    RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs, 
	        headers, HttpMethod.GET, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/lists/GetByTitle('" + title + "')/items", filter)
	        );

	    ResponseEntity<String> responseEntity = 
	        restTemplate.exchange(requestEntity, String.class);

	    return new JSONObject(responseEntity.getBody());
	}
	
	/**
	 * @param folder folder server relative URL to retrieve (/SITEURL/folder)
	 * @param jsonExtendedAttrs extended body for the query.
	 * @return json string representing folder info.
	 * @throws Exception thrown when something went wrong.
	 */
	@Override
	public JSONObject getFolderByRelativeUrl(String folder, String jsonExtendedAttrs) throws Exception {
		LOG.debug("getFolderByRelativeUrl {} jsonExtendedAttrs {}", new Object[] {folder, jsonExtendedAttrs});
	    headers = headerHelper.getGetHeaders(false);

	    RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs, 
	        headers, HttpMethod.GET, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')")
	        );

	    ResponseEntity<String> responseEntity = 
	        restTemplate.exchange(requestEntity, String.class);

	    return new JSONObject(responseEntity.getBody());
	}


	/**
	 * @param folder folder server relative URL to retrieve (/SITEURL/folder)
	 * @param jsonExtendedAttrs extended body for the query.
	 * @return json string representing list of folders
	 * @throws Exception thrown when something went wrong.
	 */
	@Override
	public JSONObject getFolderFoldersByRelativeUrl(String folder, String jsonExtendedAttrs) throws Exception {
		LOG.debug("getFolderByRelativeUrl {} jsonExtendedAttrs {}", new Object[] {folder, jsonExtendedAttrs});
		headers = headerHelper.getGetHeaders(false);

		RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs,
			  headers, HttpMethod.GET,
			  this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/Folders")
		);

		ResponseEntity<String> responseEntity =
				restTemplate.exchange(requestEntity, String.class);

		return new JSONObject(responseEntity.getBody());
	}


	/**
	 * @param folder folder server relative URL to retrieve (/SITEURL/folder)
	 * @param jsonExtendedAttrs extended body for the query.
	 * @return json string representing list of files
	 * @throws Exception thrown when something went wrong.
	 */
	@Override
	public JSONObject getFolderFilesByRelativeUrl(String folder, String jsonExtendedAttrs) throws Exception {
		LOG.debug("getFolderByRelativeUrl {} jsonExtendedAttrs {}", new Object[] {folder, jsonExtendedAttrs});
		headers = headerHelper.getGetHeaders(false);

		RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs,
			  headers, HttpMethod.GET,
			  this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/Files")
		);

		ResponseEntity<String> responseEntity =
				restTemplate.exchange(requestEntity, String.class);

		return new JSONObject(responseEntity.getBody());
	}

	/**
	 * @param fileServerRelativeUrl
	 * @return
	 * @throws Exception
	 */
	@Override
	public Boolean deleteFile(String fileServerRelativeUrl) throws Exception {
		LOG.debug("Deleting file {} ", fileServerRelativeUrl);

	    headers = headerHelper.getDeleteHeaders();
	    
	    RequestEntity<String> requestEntity = new RequestEntity<>("{}", 
	        headers, HttpMethod.POST, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + fileServerRelativeUrl +"')")
	        );

	    restTemplate.exchange(requestEntity, String.class);
	    return Boolean.TRUE;
	}

	/**
	 * @param fileServerRelativeUrl
	 * @return
	 * @throws Exception
	 */
	@Override
	public Resource downloadFile(String fileServerRelativeUrl) throws Exception {
		LOG.debug("Downloading file {} ", fileServerRelativeUrl);

	    headers = headerHelper.getGetHeaders(true);
	    
	    RequestEntity<String> requestEntity = new RequestEntity<>("", 
	        headers, HttpMethod.GET, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + fileServerRelativeUrl +"')/$value")
	        );

	    ResponseEntity<Resource> response = restTemplate.exchange(requestEntity, Resource.class);
	    return response.getBody();
	}

	/**
	 * @param folder
	 * @param resource
	 * @param jsonMetadata
	 * @return
	 * @throws Exception
	 */
	@Override
	public JSONObject uploadFile(String folder, Resource resource, JSONObject jsonMetadata) throws Exception {
		LOG.debug("Uploading file {} to folder {}", resource.getFilename(), folder);
		JSONObject submeta = new JSONObject();
		submeta.put("type", "SP.ListItem");
		jsonMetadata.put("__metadata", submeta);
		
	    headers = headerHelper.getPostHeaders("");
	    headers.remove("Content-Length");

	    RequestEntity<Resource> requestEntity = new RequestEntity<>(resource, 
	        headers, HttpMethod.POST, 
	        this.tokenHelper.getSharepointSiteUrl(
		    		"/_api/web/GetFolderByServerRelativeUrl('" + folder
							+"')/Files/add(url='" + resource.getFilename() + "',overwrite=true)"
		    		)
	        );

	    ResponseEntity<String> responseEntity = 
	        restTemplate.exchange(requestEntity, String.class);

	    String fileInfoStr = responseEntity.getBody();
	    
	    LOG.debug("Retrieved response from server with json");
	    
	    JSONObject jsonFileInfo = new JSONObject(fileInfoStr);
	    String serverRelFileUrl = jsonFileInfo.getJSONObject("d").getString("ServerRelativeUrl");

	    LOG.debug("File uploaded to URI", serverRelFileUrl);
	    String metadata = jsonMetadata.toString();
	    headers = headerHelper.getUpdateHeaders(metadata);

	    LOG.debug("Updating file adding metadata {}", jsonMetadata);

	    RequestEntity<String> requestEntity1 = new RequestEntity<>(metadata, 
	        headers, HttpMethod.POST, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + serverRelFileUrl + "')/listitemallfields")
	        );
	    ResponseEntity<String> responseEntity1 = 
		        restTemplate.exchange(requestEntity1, String.class);
	    LOG.debug("Updated file metadata Status {}", responseEntity1.getStatusCode());
	    return jsonFileInfo;
	}

	/**
	 * @param fileServerRelatUrl
	 * @param jsonMetadata
	 * @return
	 * @throws Exception
	 */
	@Override
	public JSONObject updateFileMetadata(String fileServerRelatUrl, JSONObject jsonMetadata) throws Exception {
		JSONObject meta = new JSONObject();
		meta.put("type", "SP.Folder");
		jsonMetadata.put("__metadata", meta);
	    LOG.debug("File uploaded to URI", fileServerRelatUrl);
	    String metadata = jsonMetadata.toString();
	    headers = headerHelper.getUpdateHeaders(metadata);
	    LOG.debug("Updating file adding metadata {}", jsonMetadata);

	    RequestEntity<String> requestEntity1 = new RequestEntity<>(metadata, 
	        headers, HttpMethod.POST, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + fileServerRelatUrl + "')/listitemallfields")
	        );
	    ResponseEntity<String> responseEntity1 = 
		        restTemplate.exchange(requestEntity1, String.class);
	    LOG.debug("Updated file metadata Status {}", responseEntity1.getStatusCode());
	    return new JSONObject(responseEntity1);
	}
	
	/**
	 * @param folderServerRelatUrl
	 * @param jsonMetadata
	 * @return
	 * @throws Exception
	 */
	@Override
	public JSONObject updateFolderMetadata(String folderServerRelatUrl, JSONObject jsonMetadata) throws Exception {
		JSONObject meta = new JSONObject();
		meta.put("type", "SP.Folder");
		jsonMetadata.put("__metadata", meta);
	    LOG.debug("File uploaded to URI", folderServerRelatUrl);
	    String metadata = jsonMetadata.toString();
	    headers = headerHelper.getUpdateHeaders(metadata);
	    LOG.debug("Updating file adding metadata {}", jsonMetadata);

	    RequestEntity<String> requestEntity1 = new RequestEntity<>(metadata, 
	        headers, HttpMethod.POST, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folderServerRelatUrl + "')/listitemallfields")
	        );
	    ResponseEntity<String> responseEntity1 = 
		        restTemplate.exchange(requestEntity1, String.class);
	    LOG.debug("Updated file metadata Status {}", responseEntity1.getStatusCode());
	    return new JSONObject(responseEntity1);
	}
	
	/**
	 * @param folder
	 * @return
	 * @throws Exception
	 */
	@Override
	public JSONObject breakRoleInheritance(String folder) throws Exception {
		LOG.debug("Breaking role inheritance on folder {}", folder);
	    headers = headerHelper.getPostHeaders("");

	    RequestEntity<String> requestEntity1 = new RequestEntity<>("", 
	        headers, HttpMethod.POST, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)")
        );

	    ResponseEntity<String> responseEntity1 =  restTemplate.exchange(requestEntity1, String.class);
	    return new JSONObject(responseEntity1.getBody());
	}

	/**
	 * @param baseFolderRemoteRelativeUrl
	 * @param folder
	 * @return
	 * @throws Exception
	 */
	@Override
	public JSONObject createFolder(String baseFolderRemoteRelativeUrl, String folder, JSONObject payload) throws Exception {
		LOG.debug("createFolder baseFolderRemoteRelativeUrl {} folder {}", new Object[] {baseFolderRemoteRelativeUrl, folder});
		if (payload == null) {
			payload = new JSONObject();
		}
		JSONObject meta = new JSONObject();
		meta.put("type", "SP.Folder");
		payload.put("__metadata", meta);
		payload.put("ServerRelativeUrl", baseFolderRemoteRelativeUrl + "/" + folder);
		String payloadStr = payload.toString();
		headers = headerHelper.getPostHeaders(payloadStr);
		
	    RequestEntity<String> requestEntity = new RequestEntity<>(payloadStr, 
    			headers, HttpMethod.POST, 
    			this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" +  baseFolderRemoteRelativeUrl + "')/folders")
    			);
	    ResponseEntity<String> responseEntity =  restTemplate.exchange(requestEntity, String.class);
	    return new JSONObject(responseEntity.getBody());
	}
	
	/**
	 * @param sourceRelativeServerUrl
	 * @param destinyRelativeServerUrl
	 * @return
	 * @throws Exception
	 */
	@Override
	public JSONObject moveFolder(String sourceRelativeServerUrl, String destinyRelativeServerUrl) throws Exception {
		LOG.debug("createFolder sourceRelativeServerUrl {} destinyRelativeServerUrl {}", new Object[] {sourceRelativeServerUrl, destinyRelativeServerUrl});
		headers = headerHelper.getPostHeaders("");
	    
	    RequestEntity<String> requestEntity = new RequestEntity<>("", 
    			headers, HttpMethod.POST, 
    			this.tokenHelper.getSharepointSiteUrl(
    		    		"/_api/web/GetFolderByServerRelativeUrl('" + sourceRelativeServerUrl
    		    	    		+ "')/moveto(newUrl='" + destinyRelativeServerUrl + "',flags=1)"
    		    	    		)
    			);
	    ResponseEntity<String> responseEntity =  restTemplate.exchange(requestEntity, String.class);
	    return new JSONObject(responseEntity.getBody());
	}
	
	/**
	 * @param sourceRelativeServerUrl
	 * @param destinyRelativeServerUrl
	 * @return
	 * @throws Exception
	 */
	@Override
	public JSONObject moveFile(String sourceRelativeServerUrl, String destinyRelativeServerUrl) throws Exception {
		LOG.debug("createFolder sourceRelativeServerUrl {} destinyRelativeServerUrl {}", new Object[] {sourceRelativeServerUrl, destinyRelativeServerUrl});
		headers = headerHelper.getPostHeaders("");
	    
	    RequestEntity<String> requestEntity = new RequestEntity<>("", 
    			headers, HttpMethod.POST, 
    			this.tokenHelper.getSharepointSiteUrl(
    		    		"/_api/web/GetFileByServerRelativeUrl('" + sourceRelativeServerUrl
    		    	    		+ "')/moveto(newUrl='" + destinyRelativeServerUrl + "',flags=1)"
    		    		)
    			);
	    ResponseEntity<String> responseEntity =  restTemplate.exchange(requestEntity, String.class);
	    return new JSONObject(responseEntity.getBody());
	}
	
	/**
	 * @param folderRemoteRelativeUrl
	 * @return
	 * @throws Exception
	 */
	@Override
	public Boolean removeFolder(String folderRemoteRelativeUrl) throws Exception {
		LOG.debug("Deleting folder {}", folderRemoteRelativeUrl);
		headers = headerHelper.getDeleteHeaders();

	    RequestEntity<String> requestEntity = new RequestEntity<>("", 
    			headers, HttpMethod.POST, 
    			this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folderRemoteRelativeUrl + "')")
    			);
	    restTemplate.exchange(requestEntity, String.class);
	    return Boolean.TRUE;
	}
	
	/**
	 * @param folder
	 * @param users
	 * @param permission
	 * @return
	 * @throws Exception
	 */
	@Override
	public Boolean grantPermissionToUsers(String folder, List<String> users, Permission permission) throws Exception {
		LOG.debug("Granting {} permission to users {} in folder {}", new Object[] {permission, users, folder});

	    headers = headerHelper.getGetHeaders(false);

	    List<Integer> userIds = new ArrayList<>();
	    for (String user : users) {
	    	RequestEntity<String> requestEntity = new RequestEntity<>("{}", 
	    			headers, HttpMethod.GET, 
	    			this.tokenHelper.getSharepointSiteUrl("/_api/web/SiteUsers/getByEmail('" + user+ "')")
	    			);
	    	ResponseEntity<String> responseEntity =  restTemplate.exchange(requestEntity, String.class);
	    	JSONObject objJson = new JSONObject(responseEntity.getBody());
	    	LOG.debug("json object retrieved for user {}", user);
	    	Integer userId = (Integer) objJson.getJSONObject("d").get("Id");
	    	userIds.add(userId);
	    }
	    
	    headers = headerHelper.getPostHeaders("{}");

	    for (Integer userId : userIds) {
	    	RequestEntity<String> requestEntity1 = new RequestEntity<>("{}", 
	    			headers, HttpMethod.POST, 
	    			this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/roleAssignments/addroleassignment(principalid=" + userId +",roleDefId=" + permission +")")
    			);
	    	
	    	restTemplate.exchange(requestEntity1, String.class);
	    }
	    return Boolean.TRUE;
	}
	
	/**
	 * @param folder
	 * @return
	 * @throws Exception
	 */
	@Override
	public JSONObject getFolderPermissions(String folder) throws Exception {
		headers = headerHelper.getGetHeaders(false);
	    RequestEntity<String> requestEntity1 = new RequestEntity<>("{}", 
	    		headers, HttpMethod.GET, 
	    			this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/roleAssignments")
	    		);
	    
	    ResponseEntity<String> response = restTemplate.exchange(requestEntity1, String.class);

	    return new JSONObject(response.getBody());
	}
	
	/**
	 * @param folder
	 * @param permission
	 * @return
	 * @throws Exception
	 */
	@Override
	public Boolean removePermissionToFolder(String folder, Permission permission) throws Exception {
	    List<Integer> userIds = new ArrayList<>();
	    JSONObject permissions = getFolderPermissions(folder);
	    JSONArray results = permissions.getJSONObject("d").getJSONArray("results");
	    for (int i = 0 ; i < results.length() ; i++) {
    		JSONObject jObj = results.getJSONObject(i);
    		Integer principalId = jObj.getInt("PrincipalId");
    		if (principalId != null && !userIds.contains(principalId)) {
    			userIds.add(principalId);
    		}
    		LOG.debug("JSON payload retrieved from server for user {}", "");
	    }
	    
	    headers = headerHelper.getDeleteHeaders();
	    for (Integer userId : userIds) {
	    	RequestEntity<String> requestEntity1 = new RequestEntity<>("{}", 
	    			headers, HttpMethod.POST, 
	    			this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/roleAssignments/getbyprincipalid(" + userId  +")")
			);
	    	
	    	restTemplate.exchange(requestEntity1, String.class);
	    }
	    return Boolean.TRUE;
	}
	
	/**
	 * @param folder
	 * @param users
	 * @param permission
	 * @return
	 * @throws Exception
	 */
	@Override
	public Boolean removePermissionToUsers(String folder, List<String> users, Permission permission) throws Exception {
		LOG.debug("Revoking {} permission to users {} in folder {}", new Object[] {permission, users, folder});
		
	    headers = headerHelper.getGetHeaders(false);

	    List<Integer> userIds = new ArrayList<>();
	    for (String user : users) {
	    	RequestEntity<String> requestEntity = new RequestEntity<>("{}", 
	    			headers, HttpMethod.GET, 
	    			this.tokenHelper.getSharepointSiteUrl("/_api/web/SiteUsers/getByEmail('" + user+ "')")
	    			);
	    	ResponseEntity<String> responseEntity =  restTemplate.exchange(requestEntity, String.class);
	    	LOG.debug("JSON payload retrieved from server for user {}", user);
	    	JSONObject objJson = new JSONObject(responseEntity.getBody());
	    	Integer userId = (Integer) objJson.getJSONObject("d").get("Id");
	    	userIds.add(userId);
	    }
	    
	    headers = headerHelper.getDeleteHeaders();
	    for (Integer userId : userIds) {
	    	RequestEntity<String> requestEntity1 = new RequestEntity<>("{}", 
	    			headers, HttpMethod.POST, 
	    			this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/roleAssignments/getbyprincipalid(" + userId  +")")
			);
	    	
	    	restTemplate.exchange(requestEntity1, String.class);
	    }
	    return Boolean.TRUE;
	}

	@Override
	public void refreshToken() throws Exception {
		LOG.debug("Nothing to do here as we are using a credentials provider on the rest template");
	}

}
