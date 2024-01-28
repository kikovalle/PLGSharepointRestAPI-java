package com.panxoloto.sharepoint.rest;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Supplier;
import java.util.stream.Collectors;

import org.apache.http.Header;
import org.apache.http.HttpResponse;
import org.apache.http.HttpStatus;
import org.apache.http.NameValuePair;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.client.utils.URIBuilder;
import org.apache.http.entity.StringEntity;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.entity.mime.content.ByteArrayBody;
import org.apache.http.entity.mime.content.InputStreamBody;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.message.BasicHeader;
import org.apache.http.util.EntityUtils;
import org.json.JSONArray;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.panxoloto.sharepoint.rest.helper.AuthTokenHelperOnline;
import com.panxoloto.sharepoint.rest.helper.HeadersHelper;
import com.panxoloto.sharepoint.rest.helper.Permission;

public class PLGSharepointClientOnline implements PLGSharepointClient {


	private static final Logger LOG = LoggerFactory.getLogger(PLGSharepointClientOnline.class);
	private List<Header> headers;
	private String spSiteUrl;
	private AuthTokenHelperOnline tokenHelper;
	private HeadersHelper headerHelper;
	private Supplier<HttpClientBuilder> httpClientBuilderSupplier;
	private static final String METADATA = "__metadata";

	/**
	 * @param spSiteUr.- The sharepoint site URL like https://contoso.sharepoint.com/sites/contososite
	 */
	/**
	 * @param user - The user email to access sharepoint online site.
	 * @param passwd - the user password to access sharepoint online site.
	 * @param domain - the domain without protocol and no uri like contoso.sharepoint.com
	 * @param spSiteUrl - The sharepoint site URI like /sites/contososite
	 */
	public PLGSharepointClientOnline(String user, 
			String passwd, String domain, String spSiteUrl) throws Exception {
		this(user, passwd, domain, spSiteUrl,false);
	}

	public PLGSharepointClientOnline(String user, String passwd, String domain, String site, boolean useClienId) throws Exception {
		this(user, passwd, domain, site, useClienId, HttpClients::custom);
	}

	public PLGSharepointClientOnline(String user, String passwd, String domain, String site, boolean useClienId, Supplier<HttpClientBuilder> httpClientBuilderSupplier) throws Exception {
		super();
		init(user, passwd, domain, site, useClienId, httpClientBuilderSupplier);
	}

	private void init(String user, String passwd, String domain, String spSiteUrl, boolean useClienId, Supplier<HttpClientBuilder> httpClientBuilderSupplier) throws Exception {
		this.spSiteUrl = spSiteUrl;
		this.httpClientBuilderSupplier = httpClientBuilderSupplier;
		if (this.spSiteUrl.endsWith("/")) {
			LOG.debug("spSiteUri ends with /, removing character");
			this.spSiteUrl = this.spSiteUrl.substring(0, this.spSiteUrl.length() - 1);
		}
		if (!this.spSiteUrl.startsWith("/")) {
			LOG.debug("spSiteUri doesnt start with /, adding character");
			this.spSiteUrl = String.format("%s%s", "/", this.spSiteUrl);
		}
		if (useClienId) {
			this.tokenHelper = new AuthTokenHelperOnline(true, user, passwd, domain, spSiteUrl, httpClientBuilderSupplier);
		} else {
			this.tokenHelper = new AuthTokenHelperOnline(user, passwd, domain, spSiteUrl, httpClientBuilderSupplier);
		}
		this.tokenHelper.init();
		this.headerHelper = new HeadersHelper(this.tokenHelper);
	}


	/**
	 * {@inheritDoc}
	 */
	@Override
	public void refreshToken() throws Exception {
		this.tokenHelper.init();
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject getAllLists(List<NameValuePair> data) throws Exception {
		LOG.debug("getAllLists {}", data);
	    headers = headerHelper.getGetHeaders(false);
	    try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/web/lists"));
			headers.stream().forEach(header -> httpGet.addHeader(header));
			httpGet.setURI(new URIBuilder(httpGet.getURI()).addParameters(data).build());
			HttpResponse response = client.execute(httpGet);
		    return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject getListByTitle(String title, List<NameValuePair> filterQueryData) throws Exception {
		LOG.debug("getListByTitle {} jsonExtendedAttrs", new Object[] {title});
	    headers = headerHelper.getGetHeaders(false);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + title + "')"));
			headers.stream().forEach(header -> httpGet.addHeader(header));
			httpGet.setURI(new URIBuilder(httpGet.getURI()).addParameters(filterQueryData).build());
			HttpResponse response = client.execute(httpGet);
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject getListFields(String title) throws Exception {
		LOG.debug("getListByTitle {} ", new Object[] {title});
	    headers = headerHelper.getGetHeaders(false);
	    try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
	    	HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + title + "')/Fields"));
	    	headers.stream().forEach(header -> httpGet.addHeader(header));
	    	HttpResponse response = client.execute(httpGet);
	    	return new JSONObject(EntityUtils.toString(response.getEntity()));
	    }
	}

	/**
	 * {@inheritDoc}
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
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/lists"));
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setEntity(new StringEntity(payloadStr));
			HttpResponse response = client.execute(httpPost);
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}

	/**
	 * {@inheritDoc}
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
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + listTitle + "')"));
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setEntity(new StringEntity(payloadStr));
			HttpResponse response = client.execute(httpPost);
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}

	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject getListItems(String title, List<NameValuePair> searchExtraQuery, String filter) throws Exception {
		LOG.debug("getListByTitle {} filter {}", new Object[] {title, filter});
		if (!filter.startsWith("$filter=")) {
			LOG.debug("Missing $filter in filter string, adding");
			filter = String.format("%s%s", "$filter=", filter);
		}
	    headers = headerHelper.getGetHeaders(true);
	    try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
	    	HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/lists/GetByTitle('" + title + "')/items", filter));
	    	headers.stream().forEach(header -> httpGet.addHeader(header));
	    	httpGet.setURI(new URIBuilder(httpGet.getURI()).addParameters(searchExtraQuery).build());
	    	HttpResponse response = client.execute(httpGet);
	    	return new JSONObject(EntityUtils.toString(response.getEntity()));
	    }
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject getListItem(String title, int itemId, List<NameValuePair> searchExtraQuery, String query) throws Exception {
		LOG.debug("getListItem {} itemId {} jsonExtendedAttrs {} query {}", new Object[] {title, itemId, searchExtraQuery, query});
		headers = headerHelper.getGetHeaders(true);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/lists/GetByTitle('" + title + "')/items(" + itemId + ")", query));
			headers.stream().forEach(header -> httpGet.addHeader(header));
			httpGet.setURI(new URIBuilder(httpGet.getURI()).addParameters(searchExtraQuery).build());
			HttpResponse response = client.execute(httpGet);
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject createListItem(String listTitle, String itemType, JSONObject data) throws Exception {
		LOG.debug("updateListItem list {} itemType {} data {}", new Object[] {listTitle, itemType, data});
		JSONObject payload = new JSONObject(data, JSONObject.getNames(data));
		if (itemType != null && !payload.has(METADATA)) {
			JSONObject meta = new JSONObject();
			meta.put("type", itemType);
			payload.put(METADATA, meta);
		}

		String payloadStr = payload.toString();
		headers = headerHelper.getPostHeaders(payloadStr);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + listTitle + "')/items"));
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setEntity(new StringEntity(payloadStr));
			HttpResponse response = client.execute(httpPost);
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public boolean updateListItem(String listTitle, int itemId, String itemType, JSONObject data) throws Exception {
		LOG.debug("updateListItem list {} itemId {} itemType {} data {}", new Object[] {listTitle, itemId, itemType, data});
		JSONObject payload = new JSONObject(data, JSONObject.getNames(data));
		if (itemType != null && !payload.has(METADATA)) {
			JSONObject meta = new JSONObject();
			meta.put("type", itemType);
			payload.put(METADATA, meta);
		}

		String payloadStr = payload.toString();
		headers = headerHelper.getUpdateHeaders(payloadStr);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + listTitle + "')/items(" + itemId +")"));
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setEntity(new StringEntity(payloadStr));
			CloseableHttpResponse response = client.execute(httpPost);
			return HttpStatus.SC_OK == response.getStatusLine().getStatusCode();
		}
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject getFolderByRelativeUrl(String folder, List<NameValuePair> searchExtraQuery) throws Exception {
		LOG.debug("getFolderByRelativeUrl {} jsonExtendedAttrs {}", new Object[] {folder, searchExtraQuery});
	    headers = headerHelper.getGetHeaders(false);
	    try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
	    	HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')"));
	    	headers.stream().forEach(header -> httpGet.addHeader(header));
	    	httpGet.setURI(new URIBuilder(httpGet.getURI()).addParameters(searchExtraQuery).build());
	    	HttpResponse response = client.execute(httpGet);
	    	return new JSONObject(EntityUtils.toString(response.getEntity()));
	    }
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject getFolderFoldersByRelativeUrl(String folder, List<NameValuePair> searchExtraQuery) throws Exception {
		LOG.debug("getFolderFoldersByRelativeUrl {} jsonExtendedAttrs {}", new Object[] {folder, searchExtraQuery});
		headers = headerHelper.getGetHeaders(false);
	    try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
	    	HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/Folders"));
	    	headers.stream().forEach(header -> httpGet.addHeader(header));
	    	httpGet.setURI(new URIBuilder(httpGet.getURI()).addParameters(searchExtraQuery).build());
	    	HttpResponse response = client.execute(httpGet);
	    	return new JSONObject(EntityUtils.toString(response.getEntity()));
	    }
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject getFolderFilesByRelativeUrl(String folderServerRelativeUrl) throws Exception {
		LOG.debug("getFolderFilesByRelativeUrl {} ", new Object[] {folderServerRelativeUrl});
		headers = headerHelper.getGetHeaders(false);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folderServerRelativeUrl + "')/Files"));
			headers.stream().forEach(header -> httpGet.addHeader(header));
			HttpResponse response = client.execute(httpGet);
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject getFolderFilesByRelativeUrl(String folder, List<NameValuePair> searchExtraQuery) throws Exception {
		LOG.debug("getFolderFilesByRelativeUrl {} jsonExtendedAttrs {}", new Object[] {folder, searchExtraQuery});
		headers = headerHelper.getGetHeaders(false);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/Files"));
			headers.stream().forEach(header -> httpGet.addHeader(header));
			httpGet.setURI(new URIBuilder(httpGet.getURI()).addParameters(searchExtraQuery).build());
			HttpResponse response = client.execute(httpGet);
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public Boolean deleteFile(String fileServerRelativeUrl) throws Exception {
		LOG.debug("Deleting file {} ", fileServerRelativeUrl);
	    headers = headerHelper.getDeleteHeaders();
	    try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
	    	HttpPost httpDelete = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + fileServerRelativeUrl +"')"));
	    	headers.stream().forEach(header -> httpDelete.addHeader(header));
	    	client.execute(httpDelete);
	    	return Boolean.TRUE;
	    }
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject getFileInfo(String fileServerRelativeUrl) throws Exception {
		LOG.debug("Getting file info {} ", fileServerRelativeUrl);
		headers = headerHelper.getGetHeaders(true);
	    try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
	    	HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + fileServerRelativeUrl +"')"));
	    	headers.stream().forEach(header -> httpGet.addHeader(header));
	    	HttpResponse response = client.execute(httpGet);
	    	return new JSONObject(EntityUtils.toString(response.getEntity()));
	    }
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public InputStream downloadFile(String fileServerRelativeUrl) throws Exception {
		return downloadFileWithReponse(fileServerRelativeUrl);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public InputStream downloadFileWithReponse(String fileServerRelativeUrl) throws Exception {
		LOG.debug("Downloading file {} ", fileServerRelativeUrl);
		headers = headerHelper.getGetHeaders(true);
	    try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
	    	HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + fileServerRelativeUrl +"')/$value"));
	    	headers.stream().forEach(header -> httpGet.addHeader(header));
	    	HttpResponse response = client.execute(httpGet);
	    	return response.getEntity().getContent();
	    }
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject uploadBigFile(String folder, InputStreamBody resource, JSONObject jsonMetadata, int chunkSize, String fileName)  throws Exception {
		LOG.debug("Uploading Big file {} to folder {}", resource.getFilename(), folder);
		JSONObject submeta = new JSONObject();
		if (jsonMetadata.has("type")) {
			submeta.put("type", jsonMetadata.get("type"));
		} else {
			submeta.put("type", "SP.ListItem");
		}
		jsonMetadata.put("__metadata", submeta);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			java.util.UUID uuid = java.util.UUID.randomUUID();
			String cleanFolderName = folder.startsWith(spSiteUrl) ? folder.substring(spSiteUrl.length() + 1) : folder;
			
			headers = headerHelper.getPostHeaders("");
			headers = headers.stream().filter(header -> !"content-length".equals(header.getName().toLowerCase())).collect(Collectors.toList());
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl(
					"/_api/web/GetFolderByServerRelativeUrl('" + cleanFolderName +"')/Files/add(url='"
							+ fileName + "',overwrite=true)"
					));
			httpPost.setEntity(new StringEntity(""));
			HttpResponse response = client.execute(httpPost);
			String fileInfoStr = EntityUtils.toString(response.getEntity());
			
			LOG.debug("Empty file created for chunked file upload");
			
			JSONObject jsonFileInfo = new JSONObject(fileInfoStr);
			String serverRelativeUrl = jsonFileInfo.getJSONObject("d").getString("ServerRelativeUrl");
			
			headers = headerHelper.getPostHeaders("");
			headers = headers.stream()
					.filter(header -> !"content-length".equals(header.getName().toLowerCase()))
					.filter(header -> !"accept".equals(header.getName().toLowerCase()))
					.filter(header -> !"content-type".equals(header.getName().toLowerCase()))
					.collect(Collectors.toList());
			headers.add(new BasicHeader("Content-Type", "application/octet-stream"));
			headers.add(new BasicHeader("Accept", "application/json;odata=verbose"));
			headers.add(new BasicHeader("X-RequestDigest", this.tokenHelper.getFormDigestValue()));
			byte[] bytes = new byte[chunkSize];
			try (InputStream is = resource.getInputStream();) {
				boolean firstChunk = true;
				int totalLength = is.available();
				int readed = 0;
				while (is.read(bytes) != -1) {
					readed += bytes.length;
					headers = headers.stream()
							.filter(header -> !"content-length".equals(header.getName().toLowerCase()))
							.collect(Collectors.toList());
					if (firstChunk) {
						headers.add(new BasicHeader("Content-Length", "" + bytes.length));
						ByteArrayBody body = new ByteArrayBody(bytes, resource.getFilename());
						httpPost.setEntity(MultipartEntityBuilder.create().addPart("source", body).build());
						httpPost.setURI(this.tokenHelper.getSharepointSiteUrl(
								"_api/web/getfilebyserverrelativeurl('" + ( serverRelativeUrl) +"')/startupload(uploadId=guid'" + uuid.toString() + "')"
								));
						client.execute(httpPost);
						LOG.debug("Uploaded {} of {} bytes, {} completed", new Object[] {
								readed,
								totalLength,
								(readed * 1.0) / (totalLength * 1.0)
						});
						firstChunk = false;
					} else if (readed < totalLength) {
						headers.add(new BasicHeader("Content-Length", "" + bytes.length));
						ByteArrayBody body = new ByteArrayBody(bytes, resource.getFilename());
						httpPost.setEntity(MultipartEntityBuilder.create().addPart("source", body).build());
						httpPost.setURI(this.tokenHelper.getSharepointSiteUrl(
								"/_api/web/getfilebyserverrelativeurl('" + (serverRelativeUrl) +"')/continueupload(uploadId=guid'" + uuid.toString() 
								+"',fileOffset=" 
								+ (readed -bytes.length)
								+ ")"
								));
						client.execute(httpPost);
						LOG.debug("Uploaded {} of {} bytes, {} completed", new Object[] {
								readed,
								totalLength,
								(readed * 1.0) / (totalLength * 1.0)
						});
					} else {
						headers.add(new BasicHeader("Content-Length", "" + bytes.length));
						ByteArrayBody body = new ByteArrayBody(bytes, resource.getFilename());
						httpPost.setEntity(MultipartEntityBuilder.create().addPart("source", body).build());
						httpPost.setURI(this.tokenHelper.getSharepointSiteUrl(
								"/_api/web/getfilebyserverrelativeurl('" + (serverRelativeUrl) +"')/finishupload(uploadId=guid'" + uuid.toString() + "',fileOffset="
										+ ( readed - bytes.length)
										+")"
								));
						client.execute(httpPost);
						LOG.debug("Chunked upload completed, next step is to update metadata");
					}
					
				}
			}
			
			String metadata = jsonMetadata.toString();
			headers = headerHelper.getUpdateHeaders(metadata);
			LOG.debug("Updating file adding metadata {}", jsonMetadata);
			httpPost.setEntity(new StringEntity(metadata));
			httpPost.setURI(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + serverRelativeUrl + "')/listitemallfields"));
			client.execute(httpPost);
			return jsonFileInfo;
		}
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject uploadBigFile(String folder, InputStreamBody resource, JSONObject jsonMetadata, int chunkSize) throws Exception {
		return uploadBigFile(folder, resource, jsonMetadata, chunkSize, resource.getFilename());
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject uploadFile(String folder, InputStreamBody resource, JSONObject jsonMetadata) throws Exception {
		LOG.debug("Uploading file {} to folder {}", resource.getFilename(), folder);
		return this.uploadFile(folder, resource, resource.getFilename(), jsonMetadata);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject uploadFile(String folder, InputStreamBody resource, String fileName, JSONObject jsonMetadata) throws Exception {
		LOG.debug("Uploading file {} to folder {}", fileName, folder);
		JSONObject submeta = new JSONObject();
		submeta.put("type", "SP.ListItem");
		jsonMetadata.put("__metadata", submeta);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			headers = headerHelper.getPostHeaders("");
			headers = headers.stream().filter(header -> !"content-type".equals(header.getName().toLowerCase())).collect(Collectors.toList());
			headers.add(new BasicHeader("Content-Type", "multipart/form-data"));
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl(
					"/_api/web/GetFolderByServerRelativeUrl('" + folder +"')/Files/add(url='"
							+ fileName + "',overwrite=true)"
					));
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setEntity(MultipartEntityBuilder.create()
					.setMode(HttpMultipartMode.STRICT)
					.addPart("source", resource)
					.build());
			HttpResponse response = client.execute(httpPost);
			String fileInfoStr = EntityUtils.toString(response.getEntity());
			LOG.debug("Retrieved response from server with json");
			JSONObject jsonFileInfo = new JSONObject(fileInfoStr);
			String serverRelFileUrl = jsonFileInfo.getJSONObject("d").getString("ServerRelativeUrl");
			
			LOG.debug("File uploaded to URI {}", serverRelFileUrl);
			String metadata = jsonMetadata.toString();
			headers = headerHelper.getUpdateHeaders(metadata);
			httpPost.reset();
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setURI(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + serverRelFileUrl + "')/listitemallfields"));
			httpPost.setEntity(new StringEntity(metadata));
			response = client.execute(httpPost);
			LOG.debug("Updating file adding metadata {}", jsonMetadata);
			
			LOG.debug("Updated file metadata Status {}", response.getStatusLine().getStatusCode());
			return jsonFileInfo;
		}
		
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject updateFileMetadata(String fileServerRelatUrl, JSONObject jsonMetadata) throws Exception {
		JSONObject meta = new JSONObject();
		if (jsonMetadata.has("type")) {
			meta.put("type", jsonMetadata.get("type"));
		} else {
			meta.put("type", "SP.File");
		}
		jsonMetadata.put("__metadata", meta);
	    LOG.debug("File uploaded to URI {}", fileServerRelatUrl);
	    String metadata = jsonMetadata.toString();
	    headers = headerHelper.getUpdateHeaders(metadata);
	    LOG.debug("Updating file adding metadata {}", jsonMetadata);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + fileServerRelatUrl + "')/listitemallfields"));
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setEntity(new StringEntity(metadata));
			CloseableHttpResponse response = client.execute(httpPost);
			LOG.debug("Updated file metadata Status {}", response.getStatusLine().getStatusCode());
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject updateFolderMetadata(String folderServerRelatUrl, JSONObject jsonMetadata) throws Exception {
		JSONObject meta = new JSONObject();
		if (jsonMetadata.has("type")) {
			meta.put("type", jsonMetadata.get("type"));
		} else {
			meta.put("type", "SP.Folder");
		}
		jsonMetadata.put("__metadata", meta);
	    LOG.debug("File uploaded to URI {}", folderServerRelatUrl);
	    String metadata = jsonMetadata.toString();
	    headers = headerHelper.getUpdateHeaders(metadata);
	    LOG.debug("Updating file adding metadata {}", jsonMetadata);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folderServerRelatUrl + "')/listitemallfields"));
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setEntity(new StringEntity(metadata));
			CloseableHttpResponse response = client.execute(httpPost);
			LOG.debug("Updated file metadata Status {}", response.getStatusLine().getStatusCode());
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject breakRoleInheritance(String folder) throws Exception {
		LOG.debug("Breaking role inheritance on folder {}", folder);
	    headers = headerHelper.getPostHeaders("");
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)"));
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setEntity(new StringEntity(""));
			CloseableHttpResponse response = client.execute(httpPost);
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject createFolder(String baseFolderRemoteRelativeUrl, String folder, JSONObject payload) throws Exception {
		LOG.debug("createFolder baseFolderRemoteRelativeUrl {} folder {}", new Object[] {baseFolderRemoteRelativeUrl, folder});
		if (payload == null) {
			payload = new JSONObject();
		}
		JSONObject meta = new JSONObject();
		if (payload.has("type")) {
			meta.put("type", payload.get("type"));
		} else {
			meta.put("type", "SP.Folder");
		}
		payload.put("__metadata", meta);
		payload.put("ServerRelativeUrl", baseFolderRemoteRelativeUrl + "/" + folder);
		String payloadStr = payload.toString();
		headers = headerHelper.getPostHeaders(payloadStr);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" +  baseFolderRemoteRelativeUrl + "')/folders"));
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setEntity(new StringEntity(payloadStr));
			CloseableHttpResponse response = client.execute(httpPost);
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject moveFolder(String sourceRelativeServerUrl, String destinyRelativeServerUrl) throws Exception {
		LOG.debug("createFolder sourceRelativeServerUrl {} destinyRelativeServerUrl {}", new Object[] {sourceRelativeServerUrl, destinyRelativeServerUrl});
		headers = headerHelper.getPostHeaders("");
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl(
					"/_api/web/GetFolderByServerRelativeUrl('" + sourceRelativeServerUrl
					+ "')/moveto(newUrl='" + destinyRelativeServerUrl + "',flags=1)"
					));
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setEntity(new StringEntity(""));
			CloseableHttpResponse response = client.execute(httpPost);
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject moveFile(String sourceRelativeServerUrl, String destinyRelativeServerUrl) throws Exception {
		LOG.debug("createFolder sourceRelativeServerUrl {} destinyRelativeServerUrl {}", new Object[] {sourceRelativeServerUrl, destinyRelativeServerUrl});
		headers = headerHelper.getPostHeaders("");
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl(
					"/_api/web/GetFileByServerRelativeUrl('" + sourceRelativeServerUrl
					+ "')/moveto(newUrl='"  + destinyRelativeServerUrl + "',flags=1)"
					));
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setEntity(new StringEntity(""));
			CloseableHttpResponse response = client.execute(httpPost);
			return new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}
	
	/**
	 * {@inheritDoc}
	 */
	@Override
	public Boolean removeFolder(String folderRemoteRelativeUrl) throws Exception {
		LOG.debug("Deleting folder {}", folderRemoteRelativeUrl);
		headers = headerHelper.getDeleteHeaders();
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folderRemoteRelativeUrl + "')"));
			headers.stream().forEach(header -> httpPost.addHeader(header));
			httpPost.setEntity(new StringEntity(""));
			client.execute(httpPost);
			return Boolean.TRUE;
		}
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public Boolean grantPermissionToUsers(String folder, List<String> users, Permission permission) throws Exception {
		LOG.debug("Granting {} permission to users {} in folder {}", new Object[] {permission, users, folder});

	    headers = headerHelper.getGetHeaders(false);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			List<Integer> userIds = new ArrayList<>();
			for (String user : users) {
				HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/web/SiteUsers/getByEmail('" + user+ "')"));
				headers.stream().forEach(header -> httpGet.addHeader(header));
				HttpResponse response = client.execute(httpGet);
				JSONObject objJson = new JSONObject(EntityUtils.toString(response.getEntity()));
				LOG.debug("json object retrieved for user {}", user);
				Integer userId = (Integer) objJson.getJSONObject("d").get("Id");
				userIds.add(userId);
			}
			
			headers = headerHelper.getPostHeaders("{}");
			for (Integer userId : userIds) {
				HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/roleAssignments/addroleassignment(principalid=" + userId +",roleDefId=" + permission +")"));
				headers.stream().forEach(header -> httpPost.addHeader(header));
				httpPost.setEntity(new StringEntity(""));
				client.execute(httpPost);
			}
			return Boolean.TRUE;
		}
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public JSONObject getFolderPermissions(String folder) throws Exception {
		headers = headerHelper.getGetHeaders(false);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/roleAssignments"));
			headers.stream().forEach(header -> httpGet.addHeader(header));
			HttpResponse response = client.execute(httpGet);
			return  new JSONObject(EntityUtils.toString(response.getEntity()));
		}
	}

	/**
	 * {@inheritDoc}
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
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			for (Integer userId : userIds) {
				HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/roleAssignments/getbyprincipalid(" + userId  +")"));
				headers.stream().forEach(header -> httpPost.addHeader(header));
				httpPost.setEntity(new StringEntity("{}"));
				client.execute(httpPost);
			}
			return Boolean.TRUE;
		}
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public Boolean removePermissionToUsers(String folder, List<String> users, Permission permission) throws Exception {
		LOG.debug("Revoking {} permission to users {} in folder {}", new Object[] {permission, users, folder});
		
	    headers = headerHelper.getGetHeaders(false);
		try (CloseableHttpClient client = this.httpClientBuilderSupplier.get().build()) {
			List<Integer> userIds = new ArrayList<>();
			for (String user : users) {
				HttpGet httpGet = new HttpGet(this.tokenHelper.getSharepointSiteUrl("/_api/web/SiteUsers/getByEmail('" + user+ "')"));
				headers.stream().forEach(header -> httpGet.addHeader(header));
				HttpResponse response = client.execute(httpGet);
				LOG.debug("JSON payload retrieved from server for user {}", user);
				JSONObject objJson = new JSONObject(EntityUtils.toString(response.getEntity()));
				Integer userId = (Integer) objJson.getJSONObject("d").get("Id");
				userIds.add(userId);
			}
			
			headers = headerHelper.getDeleteHeaders();
			for (Integer userId : userIds) {
				HttpPost httpPost = new HttpPost(this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/roleAssignments/getbyprincipalid(" + userId  +")"));
				headers.stream().forEach(header -> httpPost.addHeader(header));
				httpPost.setEntity(new StringEntity("{}"));
				client.execute(httpPost);
			}
			return Boolean.TRUE;
		}
	}

}
