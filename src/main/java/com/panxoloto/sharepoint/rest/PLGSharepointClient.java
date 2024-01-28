package com.panxoloto.sharepoint.rest;

import java.io.InputStream;
import java.util.List;

import org.apache.http.NameValuePair;
import org.apache.http.entity.mime.content.InputStreamBody;
import org.json.JSONObject;

import com.panxoloto.sharepoint.rest.helper.Permission;

public interface PLGSharepointClient {
	/**
	 * @throws Exception
	 */
	void refreshToken() throws Exception;
	
	/**
	 * Method to get json string wich you can transform to a JSONObject and get data from it.
	 * 
	 * 
	 * @param data - Data to be sent as query (look at the rest api documentation on how to include search filters).
	 * @return.- String representing a json object if the auth is correct.
	 * @throws Exception
	 */
	JSONObject getAllLists(List<NameValuePair> data) throws Exception;
	
	/**
	 * Retrieves list info by list title.
	 * 
	 * @param title - Site list title to query info.
	 * @param filterQueryData -query string with http client name value pairs.
	 * @return json string with list information that can be used to get a JSONObject to use this info.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getListByTitle(String title, List<NameValuePair> filterQueryData) throws Exception;

	/**
	 * Retrieves list info by list title.
	 * 
	 * @param title - Site list title to query info.
	 * @return json string with list information that can be used to get a JSONObject to use this info.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getListFields(String title) throws Exception;


	
	
	/**
	 * @param listTitle
	 * @param description
	 * @return
	 * @throws Exception
	 */
	JSONObject createList(String listTitle, String description) throws Exception;

	
	
	
	/**
	 * @param listTitle
	 * @param newDescription
	 * @return
	 * @throws Exception
	 */
	JSONObject updateList(String listTitle, String newDescription) throws Exception;

	
	/**
	 * @param title List title
	 * @param queryExtFilter Extra filter query params for extended search
	 * @param filter filter string, will be added to the query parameters as $filter, only need contenct
	 * @return
	 * @throws Exception
	 */
	JSONObject getListItems(String title, List<NameValuePair> queryExtFilter, String filter) throws Exception;
	
	/**
	 * @param title
	 * @param itemId
	 * @param queryExtFilter
	 * @param query
	 * @return
	 * @throws Exception
	 */
	JSONObject getListItem(String title, int itemId, List<NameValuePair> queryExtFilter, String query) throws Exception;
	
	/**
	 * @param folder folder server relative URL to retrieve (/SITEURL/folder)
	 * @param queryExtFilter extended body for the query.
	 * @return json string representing folder info.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getFolderByRelativeUrl(String folder, List<NameValuePair> queryExtFilter) throws Exception;

	/**
	 * @param folder folder server relative URL to retrieve (/SITEURL/folder)
	 * @param queryExtFilter extended body for the query.
	 * @return json string representing list of folders.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getFolderFoldersByRelativeUrl(String folder, List<NameValuePair> queryExtFilter) throws Exception;

	/**
	 * @param folder folder server relative URL to retrieve (/SITEURL/folder)
	 * @param queryExtFilter extended body for the query.
	 * @return json string representing list of files.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getFolderFilesByRelativeUrl(String folder, List<NameValuePair> queryExtFilter) throws Exception;

	/**
	 * @param fileServerRelativeUrl
	 * @return
	 * @throws Exception
	 */
	Boolean deleteFile(String fileServerRelativeUrl) throws Exception;

	JSONObject getFileInfo(String fileServerRelativeUrl) throws Exception;

	/**
	 * @param fileServerRelativeUrl
	 * @return
	 * @throws Exception
	 */
	InputStream downloadFile(String fileServerRelativeUrl) throws Exception;

	/**
	 * @param fileServerRelativeUrl
	 * @return
	 * @throws Exception
	 */
	InputStream downloadFileWithReponse(String fileServerRelativeUrl) throws Exception;

	
	/**
	 * @param folder
	 * @param resource
	 * @param jsonMetadata
	 * @return
	 * @throws Exception
	 */
	JSONObject uploadFile(String folder, InputStreamBody resource, JSONObject jsonMetadata) throws Exception;

	/**
	 * @param folder
	 * @param resource
	 * @param fileName
	 * @param jsonMetadata
	 * @return
	 * @throws Exception
	 */
	JSONObject uploadFile(String folder, InputStreamBody resource, String fileName, JSONObject jsonMetadata) throws Exception;
	
	/**
	 * @param fileServerRelatUrl
	 * @param jsonMetadata
	 * @return
	 * @throws Exception
	 */
	JSONObject updateFileMetadata(String fileServerRelatUrl, JSONObject jsonMetadata) throws Exception;
	
	/**
	 * @param folderServerRelatUrl
	 * @param jsonMetadata
	 * @return
	 * @throws Exception
	 */
	JSONObject updateFolderMetadata(String folderServerRelatUrl, JSONObject jsonMetadata) throws Exception;
	
	/**
	 * @param folder
	 * @return
	 * @throws Exception
	 */
	JSONObject breakRoleInheritance(String folder) throws Exception;

	/**
	 * @param baseFolderRemoteRelativeUrl
	 * @param folder
	 * @return
	 * @throws Exception
	 */
	JSONObject createFolder(String baseFolderRemoteRelativeUrl, String folder, JSONObject payload) throws Exception;
	
	/**
	 * @param sourceRelativeServerUrl
	 * @param destinyRelativeServerUrl
	 * @return
	 * @throws Exception
	 */
	JSONObject moveFolder(String sourceRelativeServerUrl, String destinyRelativeServerUrl) throws Exception;
	
	/**
	 * @param sourceRelativeServerUrl
	 * @param destinyRelativeServerUrl
	 * @return
	 * @throws Exception
	 */
	JSONObject moveFile(String sourceRelativeServerUrl, String destinyRelativeServerUrl) throws Exception;
	
	/**
	 * @param folderRemoteRelativeUrl
	 * @return
	 * @throws Exception
	 */
	Boolean removeFolder(String folderRemoteRelativeUrl) throws Exception;
	
	
	
	/**
	 * @param folder
	 * @param users
	 * @param permission
	 * @return
	 * @throws Exception
	 */
	Boolean grantPermissionToUsers(String folder, List<String> users, Permission permission) throws Exception;
	
	/**
	 * @param folder
	 * @return
	 * @throws Exception
	 */
	JSONObject getFolderPermissions(String folder) throws Exception;
	
	/**
	 * @param folder
	 * @param permission
	 * @return
	 * @throws Exception
	 */
	Boolean removePermissionToFolder(String folder, Permission permission) throws Exception;
	
	
	/**
	 * @param folder
	 * @param users
	 * @param permission
	 * @return
	 * @throws Exception
	 */
	Boolean removePermissionToUsers(String folder, List<String> users, Permission permission) throws Exception;

	/**
	 * @param folderServerRelativeUrl
	 * @return
	 * @throws Exception
	 */
	JSONObject getFolderFilesByRelativeUrl(String folderServerRelativeUrl) throws Exception;

	/**
	 * @param folder
	 * @param resource
	 * @param jsonMetadata
	 * @param chunkFileSize
	 * @return
	 * @throws Exception
	 */
	JSONObject uploadBigFile(String folder, InputStreamBody resource, JSONObject jsonMetadata, int chunkFileSize) throws Exception;

	/**
	 * @param folder
	 * @param resource
	 * @param jsonMetadata
	 * @param chunkFileSize
	 * @param fileName
	 * @return
	 */
	JSONObject uploadBigFile(String folder, InputStreamBody resource, JSONObject jsonMetadata, int chunkFileSize, String fileName) throws Exception;

	/**
	 * @param listTitle
	 * @param itemId
	 * @param itemType
	 * @param data
	 * @return
	 * @throws Exception
	 */
	boolean updateListItem(String listTitle, int itemId, String itemType, JSONObject data) throws Exception;

	/**
	 * @param listTitle
	 * @param itemType
	 * @param data
	 * @return
	 * @throws Exception
	 */
	JSONObject createListItem(String listTitle, String itemType, JSONObject data) throws Exception;
}
