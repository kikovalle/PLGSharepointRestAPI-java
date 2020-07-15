package com.panxoloto.sharepoint.rest;

import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.commons.io.IOUtils;
import org.json.JSONArray;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpMethod;
import org.springframework.http.RequestEntity;
import org.springframework.http.ResponseEntity;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.util.UriUtils;

import com.panxoloto.sharepoint.rest.helper.AuthTokenHelper;
import com.panxoloto.sharepoint.rest.helper.Permission;

public class PLGSharepointClient {


	private static final Logger LOG = LoggerFactory.getLogger(PLGSharepointClient.class);
	private MultiValueMap<String, String> headers;
	private RestTemplate restTemplate;
	private String spSiteUrl;
	private AuthTokenHelper tokenHelper;
	
	
	/**
	 * @param spSiteUr.- The sharepoint site URL like https://contoso.sharepoint.com/sites/contososite
	 */
	/**
	 * @param user.- The user email to access sharepoint online site.
	 * @param passwd.- the user password to access sharepoint online site.
	 * @param domain.- the domain without protocol and no uri like contoso.sharepoint.com
	 * @param spSiteUr.- The sharepoint site URI like /sites/contososite
	 */
	public PLGSharepointClient(String user, 
			String passwd, String domain, String spSiteUrl) {
		super();
		this.restTemplate = new RestTemplate();
		this.spSiteUrl = spSiteUrl;
		if (this.spSiteUrl.endsWith("/")) {
			LOG.debug("spSiteUri ends with /, removing character");
			this.spSiteUrl = this.spSiteUrl.substring(0, this.spSiteUrl.length() - 1);
		}
		if (!this.spSiteUrl.startsWith("/")) {
			LOG.debug("spSiteUri doesnt start with /, adding character");
			this.spSiteUrl = String.format("%s%s", "/", this.spSiteUrl);
		}
		this.tokenHelper = new AuthTokenHelper(this.restTemplate, user, passwd, domain, spSiteUrl);
		try {
			LOG.debug("Wrapper auth initialization performed successfully. Now you can perform actions on the site.");
			this.tokenHelper.init();
		} catch (Exception e) {
			LOG.error("Initialization failed!! Please check the user, pass, domain and spSiteUri parameters you provided", e);
		}
	}

	/**
	 * @throws Exception
	 */
	public void refreshToken() throws Exception {
		this.tokenHelper.init();
	}
	
	/**
	 * Method to get json string wich you can transform to a JSONObject and get data from it.
	 * 
	 * 
	 * @param data.- Data to be sent as query (look at the rest api documentation on how to include search filters).
	 * @return.- String representing a json object if the auth is correct.
	 * @throws Exception
	 */
	public JSONObject getAllLists(String data) throws Exception {
		LOG.debug("getAllLists {}", data);
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("X-RequestDigest", this.tokenHelper.getFormDigestValue());

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
	 * @param title.- Site list title to query info.
	 * @param jsonExtendedAttrs.- form json body for advanced query (take a look at ms documentation about rest api).
	 * @return json string with list information that can be used to get a JSONObject to use this info.
	 * @throws Exception thrown when something went wrong.
	 */
	public JSONObject getListByTitle(String title, String jsonExtendedAttrs) throws Exception {
		LOG.debug("getListByTitle {} jsonExtendedAttrs {}", new Object[] {title, jsonExtendedAttrs});
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("X-RequestDigest", this.tokenHelper.getFormDigestValue());

	    RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs, 
	        headers, HttpMethod.GET, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + UriUtils.encodeQuery(title, StandardCharsets.UTF_8) + "')")
	        );

	    ResponseEntity<String> responseEntity = 
	        restTemplate.exchange(requestEntity, String.class);

	    return new JSONObject(responseEntity.getBody());
	}

	/**
	 * @param title
	 * @param jsonExtendedAttrs
	 * @param filter
	 * @return
	 * @throws Exception
	 */
	public JSONObject getListItems(String title, String jsonExtendedAttrs, String filter) throws Exception {
		LOG.debug("getListByTitle {} jsonExtendedAttrs {}", new Object[] {title, jsonExtendedAttrs});
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("X-RequestDigest", this.tokenHelper.getFormDigestValue());

	    RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs, 
	        headers, HttpMethod.GET, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + UriUtils.encodeQuery(title, StandardCharsets.UTF_8) + "')/items", filter)
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
	public JSONObject getFolderByRelativeUrl(String folder, String jsonExtendedAttrs) throws Exception {
		LOG.debug("getFolderByRelativeUrl {} jsonExtendedAttrs {}", new Object[] {folder, jsonExtendedAttrs});
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("X-RequestDigest", this.tokenHelper.getFormDigestValue());

	    RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs, 
	        headers, HttpMethod.GET, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + UriUtils.encodeQuery(folder, StandardCharsets.UTF_8) + "')")
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
	public Boolean deleteFile(String fileServerRelativeUrl) throws Exception {
		LOG.debug("Deleting file {} ", fileServerRelativeUrl);

	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    headers.add("X-HTTP-Method", "DELETE");
	    headers.add("IF-Match", "*");
	    
	    RequestEntity<String> requestEntity = new RequestEntity<>("{}", 
	        headers, HttpMethod.POST, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + UriUtils.encodeQuery(fileServerRelativeUrl, StandardCharsets.UTF_8) +"')")
	        );

	    restTemplate.exchange(requestEntity, String.class);
	    return Boolean.TRUE;
	}
	
	/**
	 * @param fileServerRelativeUrl
	 * @return
	 * @throws Exception
	 */
	public Resource downloadFile(String fileServerRelativeUrl) throws Exception {
		LOG.debug("Downloading file {} ", fileServerRelativeUrl);

	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    
	    RequestEntity<String> requestEntity = new RequestEntity<>("{}", 
	        headers, HttpMethod.GET, 
	        this.tokenHelper.getSharepointSiteUrl("_api/web/GetFileByServerRelativeUrl('" + UriUtils.encodeQuery(fileServerRelativeUrl, StandardCharsets.UTF_8) +"')/$value") 
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
	public JSONObject uploadFile(String folder, Resource resource, JSONObject jsonMetadata) throws Exception {
		LOG.debug("Uploading file {} to folder {}", resource.getFilename(), folder);
		JSONObject submeta = new JSONObject();
		submeta.put("type", "SP.ListItem");
		jsonMetadata.put("__metadata", submeta);
		
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    headers.add("Content-Length", "" + resource.contentLength());
	    
	    byte[] resBytes = IOUtils.readFully(resource.getInputStream(), (int) resource.contentLength());
 
	    RequestEntity<byte[]> requestEntity = new RequestEntity<>(resBytes, 
	        headers, HttpMethod.POST, 
	        this.tokenHelper.getSharepointSiteUrl(
		    		"_api/web/GetFolderByServerRelativeUrl('" + UriUtils.encodeQuery(folder, StandardCharsets.UTF_8) +"')/Files/add(url='" 
					+ UriUtils.encodeQuery(resource.getFilename(), StandardCharsets.UTF_8) + "',overwrite=true)"
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
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("content-length", "" + metadata.getBytes().length);
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
		headers.add("X-HTTP-Method", "MERGE");
		headers.add("IF-MATCH", "*");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());

	    LOG.debug("Updating file adding metadata {}", jsonMetadata);

	    RequestEntity<String> requestEntity1 = new RequestEntity<>(metadata, 
	        headers, HttpMethod.POST, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + UriUtils.encodeQuery(serverRelFileUrl, StandardCharsets.UTF_8) + "')/listitemallfields")
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
	public JSONObject updateFileMetadata(String fileServerRelatUrl, JSONObject jsonMetadata) throws Exception {
		JSONObject meta = new JSONObject();
		meta.put("type", "SP.Folder");
		jsonMetadata.put("__metadata", meta);
	    LOG.debug("File uploaded to URI", fileServerRelatUrl);
	    String metadata = jsonMetadata.toString();
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("content-length", "" + metadata.getBytes().length);
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
		headers.add("X-HTTP-Method", "MERGE");
		headers.add("IF-MATCH", "*");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());

	    LOG.debug("Updating file adding metadata {}", jsonMetadata);

	    RequestEntity<String> requestEntity1 = new RequestEntity<>(metadata, 
	        headers, HttpMethod.POST, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + UriUtils.encodeQuery(fileServerRelatUrl, StandardCharsets.UTF_8) + "')/listitemallfields")
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
	public JSONObject updateFolderMetadata(String folderServerRelatUrl, JSONObject jsonMetadata) throws Exception {
		JSONObject meta = new JSONObject();
		meta.put("type", "SP.Folder");
		jsonMetadata.put("__metadata", meta);
	    LOG.debug("File uploaded to URI", folderServerRelatUrl);
	    String metadata = jsonMetadata.toString();
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("content-length", "" + metadata.getBytes().length);
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
		headers.add("X-HTTP-Method", "MERGE");
		headers.add("IF-MATCH", "*");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());

	    LOG.debug("Updating file adding metadata {}", jsonMetadata);

	    RequestEntity<String> requestEntity1 = new RequestEntity<>(metadata, 
	        headers, HttpMethod.POST, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + UriUtils.encodeQuery(folderServerRelatUrl, StandardCharsets.UTF_8) + "')/listitemallfields")
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
	public JSONObject breakRoleInheritance(String folder) throws Exception {
		LOG.debug("Breaking role inheritance on folder {}", folder);
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());

	    RequestEntity<String> requestEntity1 = new RequestEntity<>("", 
	        headers, HttpMethod.POST, 
	        this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + UriUtils.encodeQuery(folder, StandardCharsets.UTF_8) + "')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)")
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
		headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("Content-length", "" + payloadStr.getBytes().length);
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    
	    RequestEntity<String> requestEntity = new RequestEntity<>(payloadStr, 
    			headers, HttpMethod.POST, 
    			this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" +  UriUtils.encodeQuery(baseFolderRemoteRelativeUrl, StandardCharsets.UTF_8) + "')/folders")
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
	public JSONObject moveFolder(String sourceRelativeServerUrl, String destinyRelativeServerUrl) throws Exception {
		LOG.debug("createFolder sourceRelativeServerUrl {} destinyRelativeServerUrl {}", new Object[] {sourceRelativeServerUrl, destinyRelativeServerUrl});
		headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    headers.add("IF-Match", "*");
	    
	    RequestEntity<String> requestEntity = new RequestEntity<>("", 
    			headers, HttpMethod.POST, 
    			this.tokenHelper.getSharepointSiteUrl(
    		    		"/_api/web/GetFolderByServerRelativeUrl('" 
    		    	    		+ UriUtils.encodeQuery(sourceRelativeServerUrl , StandardCharsets.UTF_8)
    		    	    		+ "')/moveto(newUrl='" 
    		    	    		+ UriUtils.encodeQuery(destinyRelativeServerUrl, StandardCharsets.UTF_8)
    		    	    		+"',flags=1)"
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
	public JSONObject moveFile(String sourceRelativeServerUrl, String destinyRelativeServerUrl) throws Exception {
		LOG.debug("createFolder sourceRelativeServerUrl {} destinyRelativeServerUrl {}", new Object[] {sourceRelativeServerUrl, destinyRelativeServerUrl});
		headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    headers.add("IF-Match", "*");
	    
	    RequestEntity<String> requestEntity = new RequestEntity<>("", 
    			headers, HttpMethod.POST, 
    			this.tokenHelper.getSharepointSiteUrl(
    		    		"/_api/web/GetFileByServerRelativeUrl('" 
    		    	    		+ UriUtils.encodeQuery(sourceRelativeServerUrl, StandardCharsets.UTF_8) 
    		    	    		+ "')/moveto(newUrl='" 
    		    	    		+ UriUtils.encodeQuery(destinyRelativeServerUrl, StandardCharsets.UTF_8)
    		    	    		+"',flags=1)"
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
	public Boolean removeFolder(String folderRemoteRelativeUrl) throws Exception {
		LOG.debug("Deleting folder {}", folderRemoteRelativeUrl);
		headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
		headers.add("X-HTTP-Method", "DELETE");
		headers.add("IF-Match", "*");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());

	    RequestEntity<String> requestEntity = new RequestEntity<>("", 
    			headers, HttpMethod.POST, 
    			this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + UriUtils.encodeQuery(folderRemoteRelativeUrl, StandardCharsets.UTF_8) + "')")
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
	public Boolean grantPermissionToUsers(String folder, List<String> users, Permission permission) throws Exception {
		LOG.debug("Granting {} permission to users {} in folder {}", new Object[] {permission, users, folder});

	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());

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
	    
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());

	    for (Integer userId : userIds) {
	    	RequestEntity<String> requestEntity1 = new RequestEntity<>("{}", 
	    			headers, HttpMethod.POST, 
	    			this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + UriUtils.encodeQuery(folder, StandardCharsets.UTF_8) + "')/ListItemAllFields/roleAssignments/addroleassignment(principalid=" + userId +",roleDefId=" + permission +")")
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
	public JSONObject getFolderPermissions(String folder) throws Exception {
		headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    RequestEntity<String> requestEntity1 = new RequestEntity<>("{}", 
	    		headers, HttpMethod.GET, 
	    			this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + UriUtils.encodeQuery(folder, StandardCharsets.UTF_8) + "')/ListItemAllFields/roleAssignments")
	    		);
	    
	    ResponseEntity<String> response = restTemplate.exchange(requestEntity1, String.class);

	    return new JSONObject(response.getBody());
	}
	
	/**
	 * @param folder
	 * @param users
	 * @param permission
	 * @return
	 * @throws Exception
	 */
	public Boolean removePermissionToFolder(String folder, Permission permission) throws Exception {
		
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());

	    List<Integer> userIds = new ArrayList<>();
	    JSONObject permissions = getFolderPermissions(folder);
	    JSONArray results = permissions.getJSONObject("d").getJSONArray("results");
	    for (Object obj : results) {
	    	if (obj instanceof JSONObject) {
	    		System.out.println("Vamos ben....");
	    		JSONObject jObj = (JSONObject) obj;
	    		Integer principalId = jObj.getInt("PrincipalId");
	    		if (principalId != null && !userIds.contains(principalId)) {
	    			userIds.add(principalId);
	    		}
	    		LOG.debug("JSON payload retrieved from server for user {}", "");
	    	}
	    }
	    
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    headers.add("X-HTTP-Method", "DELETE");
	    for (Integer userId : userIds) {
	    	RequestEntity<String> requestEntity1 = new RequestEntity<>("{}", 
	    			headers, HttpMethod.POST, 
	    			this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + UriUtils.encodeQuery(folder, StandardCharsets.UTF_8) + "')/ListItemAllFields/roleAssignments/getbyprincipalid(" + userId  +")")
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
	public Boolean removePermissionToUsers(String folder, List<String> users, Permission permission) throws Exception {
		LOG.debug("Revoking {} permission to users {} in folder {}", new Object[] {permission, users, folder});
		
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());

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
	    
	    headers = new LinkedMultiValueMap<>();
		headers.add("Cookie",  this.tokenHelper.getCookies().stream().collect(Collectors.joining(";")) );
		headers.add("Accept", "application/json;odata=verbose");
		headers.add("Content-Type", "application/json;odata=verbose");
		headers.add("X-ClientService-ClientTag", "SDK-JAVA");
	    headers.add("Authorization", "Bearer " + this.tokenHelper.getFormDigestValue());
	    headers.add("X-HTTP-Method", "DELETE");
	    for (Integer userId : userIds) {
	    	RequestEntity<String> requestEntity1 = new RequestEntity<>("{}", 
	    			headers, HttpMethod.POST, 
	    			this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + UriUtils.encodeQuery(folder, StandardCharsets.UTF_8) + "')/ListItemAllFields/roleAssignments/getbyprincipalid(" + userId  +")")
			);
	    	
	    	restTemplate.exchange(requestEntity1, String.class);
	    }
	    return Boolean.TRUE;
	}

}
