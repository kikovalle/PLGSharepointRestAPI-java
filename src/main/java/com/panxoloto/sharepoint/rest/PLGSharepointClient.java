package com.panxoloto.sharepoint.rest;

import java.util.List;

import org.json.JSONObject;
import org.springframework.core.io.Resource;

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
	 * @param data.- Data to be sent as query (look at the rest api documentation on how to include search filters).
	 * @return.- String representing a json object if the auth is correct.
	 * @throws Exception
	 */
	JSONObject getAllLists(String data) throws Exception;
	
	/**
	 * Retrieves list info by list title.
	 * 
	 * @param title.- Site list title to query info.
	 * @param jsonExtendedAttrs.- form json body for advanced query (take a look at ms documentation about rest api).
	 * @return json string with list information that can be used to get a JSONObject to use this info.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getListByTitle(String title, String jsonExtendedAttrs) throws Exception;

	/**
	 * Retrieves list info by list title.
	 * 
	 * @param title.- Site list title to query info.
	 * @param jsonExtendedAttrs.- form json body for advanced query (take a look at ms documentation about rest api).
	 * @return json string with list information that can be used to get a JSONObject to use this info.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getListFields(String title) throws Exception;


	
	
	/**
	 * @param baseFolderRemoteRelativeUrl
	 * @param folder
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
	 * @param title
	 * @param jsonExtendedAttrs
	 * @param filter
	 * @return
	 * @throws Exception
	 */
	JSONObject getListItems(String title, String jsonExtendedAttrs, String filter) throws Exception;
	
	/**
	 * @param folder folder server relative URL to retrieve (/SITEURL/folder)
	 * @param jsonExtendedAttrs extended body for the query.
	 * @return json string representing folder info.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getFolderByRelativeUrl(String folder, String jsonExtendedAttrs) throws Exception;
	

	/**
	 * @param fileServerRelativeUrl
	 * @return
	 * @throws Exception
	 */
	Boolean deleteFile(String fileServerRelativeUrl) throws Exception;

	
	
	/**
	 * @param fileServerRelativeUrl
	 * @return
	 * @throws Exception
	 */
	Resource downloadFile(String fileServerRelativeUrl) throws Exception;

	/**
	 * @param folder
	 * @param resource
	 * @param jsonMetadata
	 * @return
	 * @throws Exception
	 */
	JSONObject uploadFile(String folder, Resource resource, JSONObject jsonMetadata) throws Exception;

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
	 * @param users
	 * @param permission
	 * @return
	 * @throws Exception
	 */
	JSONObject getFolderPermissions(String folder) throws Exception;
	
	/**
	 * @param folder
	 * @param users
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
}
