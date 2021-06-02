package com.panxoloto.sharepoint.rest.helper;

import javax.xml.soap.SOAPFault;

public class AuthenticationException
	extends RuntimeException
{
	private static final long serialVersionUID = 1L;

	public final SOAPFault reason;
	
	public AuthenticationException( final String message, final Throwable cause ) 
	{
		super(message, cause);
		this.reason = null;
	}

	public AuthenticationException(final String message) 
	{
		super(message);
		this.reason = null;
	}
	
	public AuthenticationException( final SOAPFault fault ) 
	{
		super("Authentication has failed : " + fault.getFaultString());
		this.reason = fault;
	}
	
}
