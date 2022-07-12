package com.panxoloto.sharepoint.rest;

import java.util.List;

import org.json.JSONObject;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;

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
	 * @param data Data to be sent as query (look at the rest api documentation on how to include search filters).
	 * @return JSONObject if the auth is correct.
	 * @throws Exception
	 */
	JSONObject getAllLists(String data) throws Exception;

	/**
	 * Retrieves list info by list title.
	 *
	 * @param title Site list title to query info.
	 * @param jsonExtendedAttrs form json body for advanced query (take a look at ms documentation about rest api).
	 * @return json with list information.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getListByTitle(String title, String jsonExtendedAttrs) throws Exception;

	/**
	 * Retrieves list fields by list title.
	 *
	 * @param title Site list title to query info.
	 * @return JSON of the list fields.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getListFields(String title) throws Exception;




	/**
	 * Creates a list from a title and a description.
	 * @param listTitle Title of the new list.
	 * @param description Description of the list.
	 * @return JSON of the response from this list creation.
	 * @throws Exception
	 */
	JSONObject createList(String listTitle, String description) throws Exception;




	/**
	 * Updates the list information. There needs to be a list of that title.
	 * @param listTitle Title of the list to query info before update.
	 * @param newDescription New description of the list.
	 * @return JSON of the response from this update.
	 * @throws Exception
	 */
	JSONObject updateList(String listTitle, String newDescription) throws Exception;

	/**
	 * Get all the items from the list.
	 * @param title Title of the list to query info.
	 * @param jsonExtendedAttrs Extended arguments.
	 * @param filter Filter for the items.
	 * @return JSON of the list.
	 * @throws Exception
	 */
	JSONObject getListItems(String title, String jsonExtendedAttrs, String filter) throws Exception;

	/**
	 * Get a specific item from the list.
	 * @param title Title of the list to query info.
	 * @param jsonExtendedAttrs Extended arguments.
	 * @param query Name of the item you want to query.
	 * @return JSON of the queried item.
	 * @throws Exception
	 */
	JSONObject getListItem(String title, int itemId, String jsonExtendedAttrs, String query) throws Exception;
	
	/**
	 * @param folder folder server relative URL to retrieve (/SITEURL/folder)
	 * @param jsonExtendedAttrs extended body for the query.
	 * @return JSONObject representing folder info.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getFolderByRelativeUrl(String folder, String jsonExtendedAttrs) throws Exception;

	/**
	 * @param folder folder server relative URL to retrieve (/SITEURL/folder)
	 * @param jsonExtendedAttrs extended body for the query.
	 * @return json string representing list of folders.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getFolderFoldersByRelativeUrl(String folder, String jsonExtendedAttrs) throws Exception;

	/**
	 * @param folder folder server relative URL to retrieve (/SITEURL/folder)
	 * @param jsonExtendedAttrs extended body for the query.
	 * @return JSONObject representing list of files.
	 * @throws Exception thrown when something went wrong.
	 */
	JSONObject getFolderFilesByRelativeUrl(String folder, String jsonExtendedAttrs) throws Exception;

	/**
	 * Deletes a file.
	 * @param fileServerRelativeUrl Relative URL to the file.
	 * @return True if file was deleted.
	 * @throws Exception
	 */
	Boolean deleteFile(String fileServerRelativeUrl) throws Exception;

	/**
	 * Get info on a file.
	 * @param fileServerRelativeUrl Relative URL to the file.
	 * @return JSON of the info.
	 * @throws Exception
	 */
	JSONObject getFileInfo(String fileServerRelativeUrl) throws Exception;

	/**
	 * Downloads a file.
	 * @param fileServerRelativeUrl Relative URL to the file.
	 * @return InputStreamResource to the file.
	 * @throws Exception
	 */
	InputStreamResource downloadFile(String fileServerRelativeUrl) throws Exception;

	/**
	 * Downloads a file, and includes the whole response.
	 * @param fileServerRelativeUrl Relative URL to the file.
	 * @return ResponseEntity of an InputStreamResource.
	 * @throws Exception
	 */
	ResponseEntity<InputStreamResource> downloadFileWithResponse(String fileServerRelativeUrl) throws Exception;

	
	/**
	 * @param folder
	 * @param resource
	 * @param jsonMetadata
	 * @return
	 * @throws Exception
	 */
	JSONObject uploadFile(String folder, Resource resource, JSONObject jsonMetadata) throws Exception;

	/**
	 * @param folder
	 * @param resource
	 * @param fileName
	 * @param jsonMetadata
	 * @return
	 * @throws Exception
	 */
	JSONObject uploadFile(String folder, Resource resource, String fileName, JSONObject jsonMetadata) throws Exception;
	
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
	JSONObject uploadBigFile(String folder, Resource resource, JSONObject jsonMetadata, int chunkFileSize) throws Exception;

	/**
	 * @param folder
	 * @param resource
	 * @param jsonMetadata
	 * @param chunkFileSize
	 * @param fileName
	 * @return
	 */
	JSONObject uploadBigFile(String folder, Resource resource, JSONObject jsonMetadata, int chunkFileSize, String fileName) throws Exception;

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
