package com.panxoloto.sharepoint.rest;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.UUID;

import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpMethod;
import org.springframework.http.RequestEntity;
import org.springframework.http.ResponseEntity;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.RestTemplate;

import com.panxoloto.sharepoint.rest.helper.AuthTokenHelperOnline;
import com.panxoloto.sharepoint.rest.helper.HeadersHelper;

public class ChunkFileUploader
	extends Object
{
	private final static Logger            log		= LoggerFactory.getLogger(PLGSharepointClientOnline.class);
	private final static ByteArrayResource empty	= new ChunkResource(new byte[] {});
	
	private final HeadersHelper			headerHelper;
	private final AuthTokenHelperOnline	tokenHelper;
	private final RestTemplate			restTemplate;	
	
	ChunkFileUploader( final AuthTokenHelperOnline tokenHelper )
	{
		super();
		this.tokenHelper	= tokenHelper;
		this.restTemplate	= new RestTemplate();
		this.headerHelper	= new HeadersHelper(this.tokenHelper);
	}
	
	protected long startFileUpload(final String uploadId, final String pathToTargetFile, final Resource resource)
		throws Exception
	{
		final long size = resource.contentLength();
		final String response = this.execute
		(
			"/_api/web/getFileByServerRelativeUrl('" + pathToTargetFile + "')/startupload(uploadId=guid'" + uploadId + "')",
			resource
		);		
	    final long bytesWritten = Long.parseLong( new JSONObject(response).getJSONObject("d").getString("StartUpload") );
	    if ( bytesWritten!=resource.contentLength() ) 
	    {
	    	throw new Exception( "Start chunk has not been transmitted completely [" + bytesWritten +"/" + size + "]" );
	    }
	    return bytesWritten;
	}

	protected long continueFileUpload
	(
		final String uploadId, final String pathToTargetFile, final long offset, final Resource resource
	)
		throws Exception
	{		
	    final String uri = "/_api/web/getFileByServerRelativeUrl('" + pathToTargetFile + "')/continueupload(uploadId=guid'" + uploadId + "', fileOffset=" + offset + ")";
		final String response = this.execute(uri, resource);
	    return Long.parseLong( new JSONObject(response).getJSONObject("d").getString("ContinueUpload") );
	}
	
	protected JSONObject finishFileUpload
	(
		final String uploadId, final String pathToTargetFile, final long offset, final Resource resource
	)
		throws Exception
	{		
	    final String uri = "/_api/web/getFileByServerRelativeUrl('" + pathToTargetFile + "')/finishupload(uploadId=guid'" + uploadId + "', fileOffset=" + offset + ")";
		final String response = this.execute(uri, resource);
	    return new JSONObject(response);
	}

	protected void cancelFileUploadSilently( final String uploadId, final String pathToTargetFile )
	{		
		try
		{
		    final String uri = "/_api/web/getFileByServerRelativeUrl('" + pathToTargetFile + "')/cancelupload(uploadId=guid'" + uploadId + "')";
			this.execute(uri, null);
		}
		catch( final Exception e )
		{
			log.error("Could not cancel chunkedUpload", e );
		}
	}
	
	public JSONObject uploadFile( final String folder, final Resource resource, final int size ) 
		throws Exception 
	{
		final String filename = resource.getFilename();
		log.debug("Uploading file {} to folder {}", filename, folder);
			
		final String id = UUID.randomUUID().toString();
		final String pathToTargetFile = folder + "/" + filename;
		this.createNewEmptyFile(folder, filename);
		log.debug("empty file {} has been created in folder {}", filename, folder);

		long offset = 0L;
		final byte[] buffer = new byte[size];
		try ( final InputStream is = resource.getInputStream() )
		{
			for ( int readBytes=is.read(buffer); readBytes!=-1; readBytes = is.read(buffer) )
			{
				log.debug("offset [" + offset + "] got [" + readBytes + "] bytes");
				final ChunkResource chunkResource = new ChunkResource(filename, buffer, readBytes); 
				if ( offset==0 )
				{
					offset = this.startFileUpload(id, pathToTargetFile, chunkResource);
				}
				else
				{
					offset = this.continueFileUpload(id, pathToTargetFile, offset, chunkResource);
				}
			}
		}
		return this.finishFileUpload(id, pathToTargetFile, offset, empty);
	}
	
	protected final JSONObject createNewEmptyFile( final String folder, final String newFileName )
		throws Exception
	{
		log.debug("creating new empty file {} to folder {}", newFileName, folder);
		final String fileInfoStr = this.execute
	    ( 
	    	"/_api/web/GetFolderByServerRelativeUrl('" + folder +"')/Files/add(url='" + newFileName + "',overwrite=true)",	
	    	empty
	    );
	    return new JSONObject(fileInfoStr);
	}

	private final String execute( final String call, final Resource resource )
		throws URISyntaxException
	{
		final RequestEntity<Resource> request = this.requestEntity(call, resource);
	    final ResponseEntity<String> responseEntity = this.restTemplate.exchange(request, String.class);
	    final String response = responseEntity.getBody();
	    log.debug("json response from server :" );
	    log.debug( response );
	    return response;
	}
	
	private final RequestEntity<Resource> requestEntity( final String call, final Resource resource )
		throws URISyntaxException
	{
	    return new RequestEntity<>
	    (
	    	resource, 
	        headers(), 
	        HttpMethod.POST, 
	        this.uri(call)
	    );
	}
	
	private final URI uri(final String call)
		throws URISyntaxException
	{
        return this.tokenHelper.getSharepointSiteUrl(call);
	}
	
	private final MultiValueMap<String, String> headers()
	{
		final MultiValueMap<String, String> headers = headerHelper.getPostHeaders("");
	    headers.remove("Content-Length");
	    return headers;
	}
	
	public static class ChunkResource
		extends ByteArrayResource
	{
		private final String filename;
		private final int    size;

		ChunkResource( final byte[] bytes )
		{
			this("dummy", bytes, bytes.length);
		}
		
		ChunkResource( final byte[] bytes, final int size )
		{
			this("dummy", bytes, size);
		}
		
		ChunkResource( final String filename, final byte[] bytes, final int size )
		{
			super(bytes);
			this.filename	= filename;
			this.size		= size;
		}

		@Override
		public final String getFilename() 
		{
			return this.filename;
		}

		@Override
		public long contentLength() 
		{
			return this.size;
		}

		@Override
		public InputStream getInputStream() 
			throws IOException 
		{
			return new ByteArrayInputStream(this.getByteArray(), 0, this.size);
		}
	}
}
