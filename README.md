# PLGSharepointRestAPI-java
Easy to use wrapper for the Sharepoint Rest API v1. Even if this is not a full implementation it covers most common use cases and provides examples to extending this API.

I decided to share this project here because one of the most encouraging issues I've ever found is when i tried to integrate with sharepoint online without chance of using the .Net framework. I found several java APIs that took me into headaches trying to use them. This API is a really easy to use one that covers most frequent operations I needed while integrating with Sharepoint. After a lot of research I finally got this working and I shared it.

If you find this usefull and saves you some time renember you can support me so I can achieve more time to complete this project and prepare other usefull projects that I hope will save time and efforst to someone out there.

<a href="https://www.buymeacoffee.com/kikovalle" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-blue.png" alt="Buy Me A Coffee" style="height: 51px !important;width: 217px !important;" ></a>

This is a maven project that uses spring RestTemplate to communicate with the server, but can be use in a non-spring application as you'll see in the examples I provide.

The API has been finally released to Maven Central repository (https://s01.oss.sonatype.org/#view-repositories;releases~browsestorage~io/github/kikovalle/com/panxoloto/sharepoint/rest/PLGSharepointRestAPI), so from now it is possible to include the dependency in a easier way. Simply add this to your pom.xml replacing the 1.0.3 version with the latest version of the API.

		<dependency>
			<groupId>io.github.kikovalle.com.panxoloto.sharepoint.rest</groupId>
			<artifactId>PLGSharepointRestAPI</artifactId>
			<version>1.0.3</version>
		</dependency>

As this is a maven project you have to clone this repo and compile it with maven. You can modify the pom.xml to include any distribution management repository that you use in your company so you can make use of the library in any other java project
  
    mvn clean install

Once the project is build, you can include the dependency in any other project as follow:
		
		<dependency>
			<groupId>io.github.kikovalle.com.panxoloto.sharepoint.rest</groupId>
			<artifactId>PLGSharepointRestAPI</artifactId>
			<version>1.0.3</version>
		</dependency>
  
Once this is done you can test this simple examples to perform actions in your sharepoint sites.

First step is to instantiate the API client, for this you need a sharepoint user email, a password, a domain and a sharepoint site URI:

    String user = "userwithrights@contososharepoint.com";
    String passwd = "userpasswordforthesharepointsite";
    String domain = "contoso.sharepoint.com";
    String spSiteUrl = "/sites/yoursiteorsubsitepath";

<b>Get all lists of a site</b>


    // Initialize the API
    PLGSharepointClient wrapper = new PLGSharepointClient(user, passwd, domain, spSiteUrl);
    try {
        JSONObject result = wrapper.getAllLists("{}");
        System.out.println(result);
    } catch (Exception e) {
        e.printStackTrace();
    }

<b>Get a list by list title</b>

    PLGSharepointClient wrapper = new PLGSharepointClient(user, passwd, domain, spSiteUrl);
    try {
        JSONObject result = wrapper.getListByTitle("MySharepointList", "{}");
        System.out.println(result);
    } catch (Exception e) {
        e.printStackTrace();
    }

<b>Get items of a list</b>

    PLGSharepointClient wrapper = new PLGSharepointClient(user, passwd, domain, spSiteUrl);
    try {
        // Propertyfieldname is a column name in sharepoint site, and value is the searched value, see SP Rest API to know how to filter a list.
        String queryStr = "$filter=PropertyFieldName eq 'PropertyFieldValue'";
        JSONObject result = wrapper.getListItems("MySharepointList", "{}", queryStr);
        System.out.println(result);
    } catch (Exception e) {
        e.printStackTrace();
    }
    
<b>Get a folder by server relative URL</b>

    PLGSharepointClient wrapper = new PLGSharepointClient(user, passwd, domain, spSiteUrl);
    try {
        JSONObject result = wrapper.getFolderByRelativeUrl("/sites/mysite/FolderName", "{}");
        System.out.println(result);
    } catch (Exception e) {
        e.printStackTrace();
    }

<b>Create a folder</b>

    PLGSharepointClient wrapper = new PLGSharepointClient(user, passwd, domain, spSiteUrl);
    try {
        // payload is a json object where to place metadata properties to associate with file in this example i set Title
        JSONObject payload = new JSONObject();
        payload.put("Title","Document Title set with the API");
        JSONObject result = wrapper.createFolder("/sites/mysite/parentfolderwheretocreatenew", "newfoldername", payload);
        System.out.println(result);
    } catch (Exception e) {
        e.printStackTrace();
    }

Other actions you can perform with this API are the following

<ol>
  <li>Remove a folder</li>
  <li>Upload a file</li>
  <li>Remove a file</li>
  <li>Move a folder</li>
  <li>Move a file</li>
  <li>Update file metadata</li>
  <li>Break folder role inheritance</li>
  <li>Update folder properties</li>
  <li>Grant user permissions on a folder (yet to implement file permissions control)</li>
  <li>Remove user permissions on a folder (yet to implement file permissions control)</li>
</ol>

If you find this project useful you can buy me a coffee to support this initiative

<a href="https://www.buymeacoffee.com/kikovalle" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-blue.png" alt="Buy Me A Coffee" style="height: 51px !important;width: 217px !important;" ></a>

