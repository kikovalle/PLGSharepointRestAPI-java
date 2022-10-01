package com.panxoloto.sharepoint.rest;

import com.panxoloto.sharepoint.rest.helper.AuthTokenHelperOnline;
import com.panxoloto.sharepoint.rest.helper.HeadersHelper;
import com.panxoloto.sharepoint.rest.helper.Permission;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.impl.client.HttpClients;
import org.json.JSONArray;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpMethod;
import org.springframework.http.RequestEntity;
import org.springframework.http.ResponseEntity;
import org.springframework.http.client.HttpComponentsClientHttpRequestFactory;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.RestTemplate;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Supplier;

public class PLGSharepointClientOnline implements PLGSharepointClient {


    private static final Logger LOG = LoggerFactory.getLogger(PLGSharepointClientOnline.class);
    private static final String METADATA = "__metadata";
    private MultiValueMap<String, String> headers;
    private RestTemplate restTemplate;
    private String spSiteUrl;
    private AuthTokenHelperOnline tokenHelper;
    private HeadersHelper headerHelper;

    /**
     * @param user      The user email to access sharepoint online site.
     * @param password  the user password to access sharepoint online site.
     * @param domain    the domain without protocol and no uri like contoso.sharepoint.com
     * @param spSiteUrl The sharepoint site URI like /sites/contososite
     */
    public PLGSharepointClientOnline(String user,
                                     String password, String domain, String spSiteUrl) throws Exception {
        this(user, password, domain, spSiteUrl, false);
    }

    public PLGSharepointClientOnline(String user, String password, String domain, String site, boolean useClientId) throws Exception {
        this(user, password, domain, site, useClientId, HttpClients::custom);
    }

    public PLGSharepointClientOnline(String user, String password, String domain, String site, boolean useClientId, Supplier<HttpClientBuilder> httpClientBuilderSupplier) throws Exception {
        super();
        init(user, password, domain, site, useClientId, httpClientBuilderSupplier);
    }

    private void init(String user, String password, String domain, String spSiteUrl, boolean useClientId, Supplier<HttpClientBuilder> httpClientBuilderSupplier) throws Exception {
        CloseableHttpClient httpClient = httpClientBuilderSupplier.get().build();
        HttpComponentsClientHttpRequestFactory requestFactory = new HttpComponentsClientHttpRequestFactory();
        requestFactory.setHttpClient(httpClient);
        this.restTemplate = new StreamRestTemplate(requestFactory);

        this.spSiteUrl = spSiteUrl;
        if (this.spSiteUrl.endsWith("/")) {
            LOG.debug("spSiteUri ends with /, removing character");
            this.spSiteUrl = this.spSiteUrl.substring(0, this.spSiteUrl.length() - 1);
        }
        if (!this.spSiteUrl.startsWith("/")) {
            LOG.debug("spSiteUri doesn't start with /, adding character");
            this.spSiteUrl = String.format("%s%s", "/", this.spSiteUrl);
        }
        if (useClientId) {
            this.tokenHelper = new AuthTokenHelperOnline(true, this.restTemplate, user, password, domain, spSiteUrl, httpClientBuilderSupplier);
        } else {
            this.tokenHelper = new AuthTokenHelperOnline(this.restTemplate, user, password, domain, spSiteUrl, httpClientBuilderSupplier);
        }
        this.tokenHelper.init();
        this.headerHelper = new HeadersHelper(this.tokenHelper);
    }

    @Override
    public void refreshToken() throws Exception {
        this.tokenHelper.init();
    }


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


    @Override
    public JSONObject getListByTitle(String title, String jsonExtendedAttrs) throws Exception {
        LOG.debug("getListByTitle {} jsonExtendedAttrs {}", title, jsonExtendedAttrs);
        headers = headerHelper.getGetHeaders(false);

        RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs,
                headers, HttpMethod.GET,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + title + "')")
        );

        ResponseEntity<String> responseEntity =
                restTemplate.exchange(requestEntity, String.class);

        return new JSONObject(responseEntity.getBody());
    }


    @Override
    public JSONObject getListFields(String title) throws Exception {
        LOG.debug("getListByTitle {} ", title);
        headers = headerHelper.getGetHeaders(false);

        RequestEntity<String> requestEntity = new RequestEntity<>("{}",
                headers, HttpMethod.GET,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + title + "')/Fields")
        );

        ResponseEntity<String> responseEntity =
                restTemplate.exchange(requestEntity, String.class);

        return new JSONObject(responseEntity.getBody());
    }


    @Override
    public JSONObject createList(String listTitle, String description) throws Exception {
        LOG.debug("createList listTitle {} description {}", listTitle, description);
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
        ResponseEntity<String> responseEntity = restTemplate.exchange(requestEntity, String.class);
        return new JSONObject(responseEntity.getBody());
    }


    @Override
    public JSONObject updateList(String listTitle, String newDescription) throws Exception {
        LOG.debug("update List listTitle {} description {}", listTitle, newDescription);
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
        ResponseEntity<String> responseEntity = restTemplate.exchange(requestEntity, String.class);
        return new JSONObject(responseEntity.getBody());
    }


    @Override
    public JSONObject getListItems(String title, String jsonExtendedAttrs, String filter) throws Exception {
        LOG.debug("getListItems from list {} jsonExtendedAttrs {}", title, jsonExtendedAttrs);
        headers = headerHelper.getGetHeaders(true);

        URI request = this.tokenHelper.getSharepointSiteUrl("/_api/lists/GetByTitle('" + title + "')/items", filter);
        JSONArray results = new JSONArray();
        while (request != null) {
            RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs,
                    headers, HttpMethod.GET, request);

            ResponseEntity<String> responseEntity = restTemplate.exchange(requestEntity, String.class);
            if (requestEntity.hasBody()) {
                JSONObject temp = new JSONObject(Objects.requireNonNull(responseEntity.getBody())).getJSONObject("d");
                temp.getJSONArray("results").forEach(results::put);
                if (temp.has("__next")) {
                    LOG.debug("There's another part. Let's explore it.");
                    request = new URI(temp.getString("__next"));
                } else request = null;
            }
            else request = null;
        }
        return new JSONObject("{\"d\":{\"results\":" + results + "}}");
    }


    @Override
    public JSONObject getListItem(String title, int itemId, String jsonExtendedAttrs, String query) throws Exception {
        LOG.debug("getListItem {} itemId {} jsonExtendedAttrs {} query {}", title, itemId, jsonExtendedAttrs, query);
        MultiValueMap<String, String> headers = headerHelper.getGetHeaders(true);

        RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs,
                headers, HttpMethod.GET,
                this.tokenHelper.getSharepointSiteUrl("/_api/lists/GetByTitle('" + title + "')/items(" + itemId + ")", query)
        );

        ResponseEntity<String> responseEntity =
                restTemplate.exchange(requestEntity, String.class);

        return new JSONObject(responseEntity.getBody());
    }

    @Override
    public JSONObject createListItem(String listTitle, String itemType, JSONObject data) throws Exception {
        LOG.debug("updateListItem list {} itemType {} data {}", listTitle, itemType, data);
        JSONObject payload = new JSONObject(data, JSONObject.getNames(data));
        if (itemType != null && !payload.has(METADATA)) {
            JSONObject meta = new JSONObject();
            meta.put("type", itemType);
            payload.put(METADATA, meta);
        }

        String payloadStr = payload.toString();
        MultiValueMap<String, String> headers = headerHelper.getPostHeaders(payloadStr);

        RequestEntity<String> requestEntity = new RequestEntity<>(payloadStr,
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + listTitle + "')/items")
        );
        ResponseEntity<String> responseEntity = restTemplate.exchange(requestEntity, String.class);
        return new JSONObject(responseEntity.getBody());
    }

    @Override
    public boolean updateListItem(String listTitle, int itemId, String itemType, JSONObject data) throws Exception {
        LOG.debug("updateListItem list {} itemId {} itemType {} data {}", listTitle, itemId, itemType, data);
        JSONObject payload = new JSONObject(data, JSONObject.getNames(data));
        if (itemType != null && !payload.has(METADATA)) {
            JSONObject meta = new JSONObject();
            meta.put("type", itemType);
            payload.put(METADATA, meta);
        }

        String payloadStr = payload.toString();
        MultiValueMap<String, String> headers = headerHelper.getUpdateHeaders(payloadStr);

        RequestEntity<String> requestEntity = new RequestEntity<>(payloadStr,
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/lists/GetByTitle('" + listTitle + "')/items(" + itemId + ")")
        );
        ResponseEntity<String> responseEntity = restTemplate.exchange(requestEntity, String.class);
        return responseEntity.getStatusCode().is2xxSuccessful();
    }


    @Override
    public JSONObject getFolderByRelativeUrl(String folder, String jsonExtendedAttrs) throws Exception {
        LOG.debug("getFolderByRelativeUrl {} jsonExtendedAttrs {}", folder, jsonExtendedAttrs);
        headers = headerHelper.getGetHeaders(false);

        RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs,
                headers, HttpMethod.GET,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')")
        );

        ResponseEntity<String> responseEntity =
                restTemplate.exchange(requestEntity, String.class);

        return new JSONObject(responseEntity.getBody());
    }

    @Override
    public JSONObject getFolderFoldersByRelativeUrl(String folder, String jsonExtendedAttrs) throws Exception {
        LOG.debug("getFolderFoldersByRelativeUrl {} jsonExtendedAttrs {}", folder, jsonExtendedAttrs);
        headers = headerHelper.getGetHeaders(false);

        RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs,
                headers, HttpMethod.GET,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/Folders")
        );

        ResponseEntity<String> responseEntity =
                restTemplate.exchange(requestEntity, String.class);

        return new JSONObject(responseEntity.getBody());
    }

    @Override
    public JSONObject getFolderFilesByRelativeUrl(String folderServerRelativeUrl) throws Exception {
        LOG.debug("getFolderFilesByRelativeUrl {} ", folderServerRelativeUrl);
        headers = headerHelper.getGetHeaders(false);

        RequestEntity<String> requestEntity = new RequestEntity<>("{}",
                headers, HttpMethod.GET,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folderServerRelativeUrl + "')/Files")
        );

        ResponseEntity<String> responseEntity =
                restTemplate.exchange(requestEntity, String.class);

        return new JSONObject(responseEntity.getBody());
    }

    @Override
    public JSONObject getFolderFilesByRelativeUrl(String folder, String jsonExtendedAttrs) throws Exception {
        LOG.debug("getFolderFilesByRelativeUrl {} jsonExtendedAttrs {}", folder, jsonExtendedAttrs);
        headers = headerHelper.getGetHeaders(false);

        RequestEntity<String> requestEntity = new RequestEntity<>(jsonExtendedAttrs,
                headers, HttpMethod.GET,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/Files")
        );

        ResponseEntity<String> responseEntity =
                restTemplate.exchange(requestEntity, String.class);

        return new JSONObject(responseEntity.getBody());
    }

    @Override
    public Boolean deleteFile(String fileServerRelativeUrl) throws Exception {
        LOG.debug("Deleting file {} ", fileServerRelativeUrl);

        headers = headerHelper.getDeleteHeaders();

        RequestEntity<String> requestEntity = new RequestEntity<>("{}",
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + fileServerRelativeUrl + "')")
        );

        restTemplate.exchange(requestEntity, String.class);
        return Boolean.TRUE;
    }

    @Override
    public JSONObject getFileInfo(String fileServerRelativeUrl) throws Exception {
        LOG.debug("Getting file info {} ", fileServerRelativeUrl);

        headers = headerHelper.getGetHeaders(true);

        RequestEntity<String> requestEntity = new RequestEntity<>("",
                headers, HttpMethod.GET,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + fileServerRelativeUrl + "')")
        );

        ResponseEntity<String> responseEntity = restTemplate.exchange(requestEntity, String.class);
        return new JSONObject(responseEntity.getBody());
    }

    @Override
    public InputStreamResource downloadFile(String fileServerRelativeUrl) throws Exception {
        return downloadFileWithResponse(fileServerRelativeUrl).getBody();
    }

    @Override
    public ResponseEntity<InputStreamResource> downloadFileWithResponse(String fileServerRelativeUrl) throws Exception {
        LOG.debug("Downloading file {} ", fileServerRelativeUrl);

        MultiValueMap<String, String> headers = headerHelper.getGetHeaders(true);

        RequestEntity<String> requestEntity = new RequestEntity<>("",
                headers, HttpMethod.GET,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + fileServerRelativeUrl + "')/$value")
        );

        return restTemplate.exchange(requestEntity, InputStreamResource.class);
    }

    @Override
    public JSONObject uploadBigFile(String folder, Resource resource, JSONObject jsonMetadata, int chunkSize, String fileName) throws Exception {
        LOG.debug("Uploading Big file {} to folder {}", resource.getFilename(), folder);
        JSONObject subMeta = new JSONObject();
        if (jsonMetadata.has("type")) {
            subMeta.put("type", jsonMetadata.get("type"));
        } else {
            subMeta.put("type", "SP.ListItem");
        }
        jsonMetadata.put("__metadata", subMeta);
        java.util.UUID uuid = java.util.UUID.randomUUID();
        String cleanFolderName = folder.startsWith(spSiteUrl) ? folder.substring(spSiteUrl.length() + 1) : folder;

        Resource tmpRes = new ByteArrayResource(new byte[0]);
        headers = headerHelper.getPostHeaders("");
        headers.remove("Content-Length");

        RequestEntity<Resource> requestEntityCreate = new RequestEntity<>(tmpRes,
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl(
                        "/_api/web/GetFolderByServerRelativeUrl('" + cleanFolderName + "')/Files/add(url='"
                                + fileName + "',overwrite=true)"
                )
        );

        ResponseEntity<String> tmpResponse = restTemplate.exchange(requestEntityCreate, String.class);
        String fileInfoStr = tmpResponse.getBody();

        LOG.debug("Empty file created for chunked file upload");

        JSONObject jsonFileInfo = new JSONObject(fileInfoStr);
        String serverRelativeUrl = jsonFileInfo.getJSONObject("d").getString("ServerRelativeUrl");

        headers = headerHelper.getPostHeaders("");
        headers.remove("Content-Length");
        headers.remove("Content-length");
        headers.remove("Accept");
        headers.remove("Content-Type");
        headers.add("Content-Type", "application/octet-stream");
        headers.add("Accept", "application/json;odata=verbose");
        headers.add("X-RequestDigest", this.tokenHelper.getFormDigestValue());
        byte[] bytes = new byte[chunkSize];
        try (InputStream inputStream = resource.getInputStream()) {
            boolean firstChunk = true;
            int totalLength = inputStream.available();
            int nbrBytesRead = 0;
            while (inputStream.read(bytes) != -1) {
                nbrBytesRead += bytes.length;
                headers.remove("Content-Length");
                if (firstChunk) {
                    headers.add("Content-Length", "" + bytes.length);
                    RequestEntity<byte[]> requestEntity = new RequestEntity<>(bytes,
                            headers, HttpMethod.POST,
                            this.tokenHelper.getSharepointSiteUrl(
                                    "_api/web/getfilebyserverrelativeurl('" + (serverRelativeUrl) + "')/startupload(uploadId=guid'" + uuid + "')"
                            )
                    );
                    restTemplate.exchange(requestEntity, String.class);
                    LOG.debug("Uploaded {} of {} bytes, {} completed", nbrBytesRead,
                            totalLength,
                            (nbrBytesRead * 1.0) / (totalLength * 1.0));
                    firstChunk = false;
                } else if (nbrBytesRead < totalLength) {
                    RequestEntity<byte[]> requestEntity = new RequestEntity<>(bytes,
                            headers, HttpMethod.POST,
                            this.tokenHelper.getSharepointSiteUrl(
                                    "/_api/web/getfilebyserverrelativeurl('" + (serverRelativeUrl) + "')/continueupload(uploadId=guid'" + uuid
                                            + "',fileOffset="
                                            + (nbrBytesRead - bytes.length)
                                            + ")"
                            )
                    );
                    restTemplate.exchange(requestEntity, String.class);
                    LOG.debug("Uploaded {} of {} bytes, {} completed", nbrBytesRead,
                            totalLength,
                            (nbrBytesRead * 1.0) / (totalLength * 1.0));
                } else {
                    RequestEntity<byte[]> requestEntity = new RequestEntity<>(bytes,
                            headers, HttpMethod.POST,
                            this.tokenHelper.getSharepointSiteUrl(
                                    "/_api/web/getfilebyserverrelativeurl('" + (serverRelativeUrl) + "')/finishupload(uploadId=guid'" + uuid + "',fileOffset="
                                            + (nbrBytesRead - bytes.length)
                                            + ")"
                            )
                    );
                    restTemplate.exchange(requestEntity, String.class);
                    LOG.debug("Chunked upload completed, next step is to update metadata");
                }

            }
        }

        String metadata = jsonMetadata.toString();
        headers = headerHelper.getUpdateHeaders(metadata);
        LOG.debug("Updating file adding metadata {}", jsonMetadata);

        RequestEntity<String> requestEntity1 = new RequestEntity<>(metadata,
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + serverRelativeUrl + "')/listitemallfields")
        );
        restTemplate.exchange(requestEntity1, String.class);
        return jsonFileInfo;
    }


    @Override
    public JSONObject uploadBigFile(String folder, Resource resource, JSONObject jsonMetadata, int chunkSize) throws Exception {
        return uploadBigFile(folder, resource, jsonMetadata, chunkSize, resource.getFilename());
    }

    @Override
    public JSONObject uploadFile(String folder, Resource resource, JSONObject jsonMetadata) throws Exception {
        LOG.debug("Uploading file {} to folder {}", resource.getFilename(), folder);
        JSONObject subMeta = new JSONObject();
        if (jsonMetadata.has("type")) {
            subMeta.put("type", jsonMetadata.get("type"));
        } else {
            subMeta.put("type", "SP.ListItem");
        }
        jsonMetadata.put("__metadata", subMeta);

        headers = headerHelper.getPostHeaders("");
        headers.remove("Content-length");
        headers.remove("Content-Type");
        headers.add("Content-Type", "multipart/form-data");

        RequestEntity<Resource> requestEntity = new RequestEntity<>(resource,
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl(
                        "/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/Files/add(url='"
                                + resource.getFilename() + "',overwrite=true)"
                )
        );

        ResponseEntity<String> responseEntity =
                restTemplate.exchange(requestEntity, String.class);

        String fileInfoStr = responseEntity.getBody();

        LOG.debug("Retrieved response from server with json");

        JSONObject jsonFileInfo = new JSONObject(fileInfoStr);
        String serverRelFileUrl = jsonFileInfo.getJSONObject("d").getString("ServerRelativeUrl");

        LOG.debug("File uploaded to URI {}", serverRelFileUrl);
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

    @Override
    public JSONObject uploadFile(String folder, Resource resource, String fileName, JSONObject jsonMetadata) throws Exception {
        LOG.debug("Uploading file {} to folder {}", fileName, folder);
        JSONObject subMeta = new JSONObject();
        subMeta.put("type", "SP.ListItem");
        jsonMetadata.put("__metadata", subMeta);

        headers = headerHelper.getPostHeaders("");
        headers.remove("Content-length");
        headers.remove("Content-Type");
        headers.add("Content-Type", "multipart/form-data");

        RequestEntity<Resource> requestEntity = new RequestEntity<>(resource,
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl(
                        "/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/Files/add(url='"
                                + fileName + "',overwrite=true)"
                )
        );

        ResponseEntity<String> responseEntity =
                restTemplate.exchange(requestEntity, String.class);

        String fileInfoStr = responseEntity.getBody();

        LOG.debug("Retrieved response from server with json");

        JSONObject jsonFileInfo = new JSONObject(fileInfoStr);
        String serverRelFileUrl = jsonFileInfo.getJSONObject("d").getString("ServerRelativeUrl");

        LOG.debug("File uploaded to URI {}", serverRelFileUrl);
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

    @Override
    public JSONObject updateFileMetadata(String fileServerRelativeUrl, JSONObject jsonMetadata) throws Exception {
        JSONObject meta = new JSONObject();
        if (jsonMetadata.has("type")) {
            meta.put("type", jsonMetadata.get("type"));
        } else {
            meta.put("type", "SP.File");
        }
        jsonMetadata.put("__metadata", meta);
        LOG.debug("File uploaded to URI {}", fileServerRelativeUrl);
        String metadata = jsonMetadata.toString();
        headers = headerHelper.getUpdateHeaders(metadata);
        LOG.debug("Updating file adding metadata {}", jsonMetadata);

        RequestEntity<String> requestEntity1 = new RequestEntity<>(metadata,
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFileByServerRelativeUrl('" + fileServerRelativeUrl + "')/listitemallfields")
        );
        ResponseEntity<String> responseEntity1 =
                restTemplate.exchange(requestEntity1, String.class);
        LOG.debug("Updated file metadata Status {}", responseEntity1.getStatusCode());
        return new JSONObject(responseEntity1);
    }

    @Override
    public JSONObject updateFolderMetadata(String folderServerRelativeUrl, JSONObject jsonMetadata) throws Exception {
        JSONObject meta = new JSONObject();
        if (jsonMetadata.has("type")) {
            meta.put("type", jsonMetadata.get("type"));
        } else {
            meta.put("type", "SP.Folder");
        }
        jsonMetadata.put("__metadata", meta);
        LOG.debug("File uploaded to URI {}", folderServerRelativeUrl);
        String metadata = jsonMetadata.toString();
        headers = headerHelper.getUpdateHeaders(metadata);
        LOG.debug("Updating file adding metadata {}", jsonMetadata);

        RequestEntity<String> requestEntity1 = new RequestEntity<>(metadata,
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folderServerRelativeUrl + "')/listitemallfields")
        );
        ResponseEntity<String> responseEntity1 =
                restTemplate.exchange(requestEntity1, String.class);
        LOG.debug("Updated file metadata Status {}", responseEntity1.getStatusCode());
        return new JSONObject(responseEntity1);
    }

    @Override
    public JSONObject breakRoleInheritance(String folder) throws Exception {
        LOG.debug("Breaking role inheritance on folder {}", folder);
        headers = headerHelper.getPostHeaders("");

        RequestEntity<String> requestEntity1 = new RequestEntity<>("",
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)")
        );

        ResponseEntity<String> responseEntity1 = restTemplate.exchange(requestEntity1, String.class);
        return new JSONObject(responseEntity1.getBody());
    }

    @Override
    public JSONObject createFolder(String baseFolderRemoteRelativeUrl, String folder, JSONObject payload) throws Exception {
        LOG.debug("createFolder baseFolderRemoteRelativeUrl {} folder {}", baseFolderRemoteRelativeUrl, folder);
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

        RequestEntity<String> requestEntity = new RequestEntity<>(payloadStr,
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + baseFolderRemoteRelativeUrl + "')/folders")
        );
        ResponseEntity<String> responseEntity = restTemplate.exchange(requestEntity, String.class);
        return new JSONObject(responseEntity.getBody());
    }

    @Override
    public JSONObject moveFolder(String sourceRelativeServerUrl, String destinyRelativeServerUrl) throws Exception {
        LOG.debug("createFolder sourceRelativeServerUrl {} destinyRelativeServerUrl {}", sourceRelativeServerUrl, destinyRelativeServerUrl);
        headers = headerHelper.getPostHeaders("");

        RequestEntity<String> requestEntity = new RequestEntity<>("",
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl(
                        "/_api/web/GetFolderByServerRelativeUrl('" + sourceRelativeServerUrl
                                + "')/moveto(newUrl='" + destinyRelativeServerUrl + "',flags=1)"
                )
        );
        ResponseEntity<String> responseEntity = restTemplate.exchange(requestEntity, String.class);
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
        LOG.debug("createFolder sourceRelativeServerUrl {} destinyRelativeServerUrl {}", sourceRelativeServerUrl, destinyRelativeServerUrl);
        headers = headerHelper.getPostHeaders("");

        RequestEntity<String> requestEntity = new RequestEntity<>("",
                headers, HttpMethod.POST,
                this.tokenHelper.getSharepointSiteUrl(
                        "/_api/web/GetFileByServerRelativeUrl('" + sourceRelativeServerUrl
                                + "')/moveto(newUrl='" + destinyRelativeServerUrl + "',flags=1)"
                )
        );
        ResponseEntity<String> responseEntity = restTemplate.exchange(requestEntity, String.class);
        return new JSONObject(responseEntity.getBody());
    }

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

    @Override
    public Boolean grantPermissionToUsers(String folder, List<String> users, Permission permission) throws Exception {
        LOG.debug("Granting {} permission to users {} in folder {}", permission, users, folder);

        headers = headerHelper.getGetHeaders(false);

        List<Integer> userIds = new ArrayList<>();
        for (String user : users) {
            RequestEntity<String> requestEntity = new RequestEntity<>("{}",
                    headers, HttpMethod.GET,
                    this.tokenHelper.getSharepointSiteUrl("/_api/web/SiteUsers/getByEmail('" + user + "')")
            );
            ResponseEntity<String> responseEntity = restTemplate.exchange(requestEntity, String.class);
            JSONObject objJson = new JSONObject(responseEntity.getBody());
            LOG.debug("json object retrieved for user {}", user);
            Integer userId = (Integer) objJson.getJSONObject("d").get("Id");
            userIds.add(userId);
        }

        headers = headerHelper.getPostHeaders("{}");

        for (Integer userId : userIds) {
            RequestEntity<String> requestEntity1 = new RequestEntity<>("{}",
                    headers, HttpMethod.POST,
                    this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/roleAssignments/addroleassignment(principalid=" + userId + ",roleDefId=" + permission + ")")
            );

            restTemplate.exchange(requestEntity1, String.class);
        }
        return Boolean.TRUE;
    }

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

    @Override
    public Boolean removePermissionToFolder(String folder, Permission permission) throws Exception {
        List<Integer> userIds = new ArrayList<>();
        JSONObject permissions = getFolderPermissions(folder);
        JSONArray results = permissions.getJSONObject("d").getJSONArray("results");
        for (int i = 0; i < results.length(); i++) {
            JSONObject jObj = results.getJSONObject(i);
            Integer principalId = jObj.getInt("PrincipalId");
            // This will throw an exception if the key isn't in the JSON.
            if (!userIds.contains(principalId)) {
                userIds.add(principalId);
            }
            LOG.debug("JSON payload retrieved from server for user {}", "");
        }

        headers = headerHelper.getDeleteHeaders();
        for (Integer userId : userIds) {
            RequestEntity<String> requestEntity1 = new RequestEntity<>("{}",
                    headers, HttpMethod.POST,
                    this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/roleAssignments/getbyprincipalid(" + userId + ")")
            );

            restTemplate.exchange(requestEntity1, String.class);
        }
        return Boolean.TRUE;
    }

    @Override
    public Boolean removePermissionToUsers(String folder, List<String> users, Permission permission) throws Exception {
        LOG.debug("Revoking {} permission to users {} in folder {}", permission, users, folder);

        headers = headerHelper.getGetHeaders(false);

        List<Integer> userIds = new ArrayList<>();
        for (String user : users) {
            RequestEntity<String> requestEntity = new RequestEntity<>("{}",
                    headers, HttpMethod.GET,
                    this.tokenHelper.getSharepointSiteUrl("/_api/web/SiteUsers/getByEmail('" + user + "')")
            );
            ResponseEntity<String> responseEntity = restTemplate.exchange(requestEntity, String.class);
            LOG.debug("JSON payload retrieved from server for user {}", user);
            JSONObject objJson = new JSONObject(responseEntity.getBody());
            Integer userId = (Integer) objJson.getJSONObject("d").get("Id");
            userIds.add(userId);
        }

        headers = headerHelper.getDeleteHeaders();
        for (Integer userId : userIds) {
            RequestEntity<String> requestEntity1 = new RequestEntity<>("{}",
                    headers, HttpMethod.POST,
                    this.tokenHelper.getSharepointSiteUrl("/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/ListItemAllFields/roleAssignments/getbyprincipalid(" + userId + ")")
            );

            restTemplate.exchange(requestEntity1, String.class);
        }
        return Boolean.TRUE;
    }

    public final ChunkFileUploader createChunkFileUploader() {
        return new ChunkFileUploader(this.tokenHelper);
    }

}
