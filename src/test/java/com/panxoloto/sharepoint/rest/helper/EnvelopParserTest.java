package com.panxoloto.sharepoint.rest.helper;
import java.io.CharArrayWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.nio.charset.Charset;

import org.testng.annotations.Test;

import com.panxoloto.sharepoint.rest.helper.AuthenticationException;
import com.panxoloto.sharepoint.rest.helper.AuthenticationResponseParser;

public class EnvelopParserTest
	extends Object
{
	public final static String soapEnvelopeNamespace = "http://www.w3.org/2003/05/soap-envelope";
	
	@Test
	(
		expectedExceptions = AuthenticationException.class,
		expectedExceptionsMessageRegExp = "Authentication has failed : Authentication Failure"
	)
	public final void soapTest_message_with_Fault()
		throws Exception
	{
		final String response = charsUtf8("com/panxoloto/sharepoint/rest/helper/authentication-error-response.xml");
		AuthenticationResponseParser.parseAuthenticationResponse(response);
	}

	@Test
	(
		expectedExceptions = AuthenticationException.class,
		expectedExceptionsMessageRegExp = "Could not parse authentication response : Unable to create envelope from given source: "
	)
	public final void soapTest_not_a_soap_message()
		throws Exception
	{
		AuthenticationResponseParser.parseAuthenticationResponse("<root></root>");
	}

	@Test
	public final void soapTest_withSequrityToken()
		throws Exception
	{
		final String response = charsUtf8("com/panxoloto/sharepoint/rest/helper/authentication-success-response.xml");
		final String token = AuthenticationResponseParser.parseAuthenticationResponse(response);
		assert "t=TOKENp=".equals(token);
	}
	
	public static String charsUtf8(	final String resource ) 
		throws IOException
	{
		return chars(resource, AuthenticationResponseParser.utf8);
	}
	
	public static String chars
	(
		final String resource, final Charset cs
	) 
		throws IOException
	{
		try ( final InputStream is = EnvelopParserTest.class.getClassLoader().getResourceAsStream(resource) )
		{
			try ( final Reader r = new InputStreamReader(is, cs) )
			{
				try ( CharArrayWriter w = new CharArrayWriter() )
				{
					final char[] b = new char[10000];
					int charsWereRead = 0;
					while ( (charsWereRead = r.read(b))!=-1 )
					{
						w.write(b, 0, charsWereRead);
					}
					return new String(w.toCharArray());
				}
			}
		}
	}
}
