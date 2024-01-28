package com.panxoloto.sharepoint.rest.helper;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

import javax.xml.namespace.QName;

import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import jakarta.xml.soap.MessageFactory;
import jakarta.xml.soap.SOAPBody;
import jakarta.xml.soap.SOAPConstants;
import jakarta.xml.soap.SOAPException;
import jakarta.xml.soap.SOAPFault;
import jakarta.xml.soap.SOAPMessage;

public class AuthenticationResponseParser
{
	public static final Charset utf8 = StandardCharsets.UTF_8;
	public static final QName binarySecurityTokenName = new QName
	(
		"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd", "BinarySecurityToken"
	);

	private AuthenticationResponseParser() {
	}

	public static String  parseAuthenticationResponse(final String response )
		throws AuthenticationException 
	{
		try ( final InputStream is = new ByteArrayInputStream(response.getBytes(utf8)) )
		{
			final MessageFactory	f		= MessageFactory.newInstance(SOAPConstants.SOAP_1_2_PROTOCOL);
			final SOAPMessage		m		= f.createMessage(null, is);
			final SOAPBody			body	= m.getSOAPBody();
			final SOAPFault			fault	= body.getFault();
			if ( fault!=null )
			{
				throw new AuthenticationException(fault);
			}
			return token(body);
		}
		catch( final SOAPException soapExc )
		{
			throw new AuthenticationException("Could not parse authentication response : " + soapExc.getMessage(), soapExc);
		}
		catch( final IOException ioExc )
		{
			throw new AuthenticationException("Unexpected IO exception occured", ioExc);
		}
	}
	
	private static String token( final SOAPBody body )
	{
		final NodeList list=body.getElementsByTagNameNS(binarySecurityTokenName.getNamespaceURI(), binarySecurityTokenName.getLocalPart());
		if ( list.getLength()>0 )
		{
			final Node node = list.item(0);
			if ( node instanceof Element )
			{
				final Element tokenEl = (Element)node;
				return tokenEl.getTextContent();
			}
		}
		throw new AuthenticationException("Authentication response does not contain mandatory element " + binarySecurityTokenName);
	}
}
