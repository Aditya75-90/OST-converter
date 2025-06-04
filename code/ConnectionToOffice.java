package email.code;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;

import com.aspose.email.EWSClient;
import com.aspose.email.EmailClient;
import com.aspose.email.IEWSClient;
import com.aspose.email.ITokenProvider;
import com.aspose.email.ImapClient;
import com.aspose.email.OAuthNetworkCredential;
import com.aspose.email.SecurityOptions;

import com.aspose.email.system.ICredentials;
import com.aspose.email.system.NetworkCredential;
import com.chilkatsoft.CkAuthAzureAD;
import com.chilkatsoft.CkAuthAzureSAS;
import com.chilkatsoft.CkAuthAzureStorage;
import com.chilkatsoft.CkAuthGoogle;
import com.chilkatsoft.CkDateTime;
import com.chilkatsoft.CkGlobal;
import com.chilkatsoft.CkHttp;
import com.chilkatsoft.CkImap;
import com.chilkatsoft.CkJsonObject;
import com.chilkatsoft.CkOAuth2;
import com.chilkatsoft.CkRest;
import com.chilkatsoft.CkStringBuilder;

public class ConnectionToOffice {
	static String username;
	static CkOAuth2 oauth2;
	static String token;
	static CkJsonObject json=null;
	public static boolean auth() {
		CkGlobal	glob = new CkGlobal();
		boolean success1 = glob.UnlockBundle("SNKRWT.CB1122022_Q5uG5AzJlRm1");
		if (success1 != true) {
			System.out.println(glob.lastErrorText());
		
		}

		int status = glob.get_UnlockStatus();
		if (status == 2) {
			System.out.println("Unlocked using purchased unlock code.");
		} else {
			System.out.println("Unlocked in trial mode.");
		}
		System.out.println(success1 + " 42");
		
		oauth2 = new CkOAuth2();
		boolean success;
		String id = "1994b9da-99f0-4c6b-9f78-75c47976d339";
		String secret = "ww.8Q~G7N4.J6DSekM1OHiQX6UogLo.b4aMVfckf";

		oauth2.put_ListenPort(3017);
		oauth2.put_AuthorizationEndpoint("https://login.microsoftonline.com/common/oauth2/v2.0/authorize");
		// oauth2.put_AuthorizationEndpoint("https://login.microsoftonline.com/common/oauth2/v2.0/authorize");

		oauth2.put_TokenEndpoint("https://login.microsoftonline.com/common/oauth2/v2.0/token");

		oauth2.put_ClientId(id);

		oauth2.put_ClientSecret(secret);

		oauth2.put_CodeChallenge(false);

		oauth2.put_CodeChallenge(true);

		oauth2.codeChallengeMethod();// CodeChallengeMethod = "S256";

		// oauth2.put_Scope("openid profile offline_access user.readwrite mail.readwrite
		// mail.send files.readwrite");
//		oauth2.put_Scope(
//				"openid profile offline_access https://outlook.office365.com/SMTP.Send https://outlook.office365.com/POP.AccessAsUser.All https://outlook.office365.com/IMAP.AccessAsUser.All offline_access user.readwrite mail.readwrite mail.send files.readwrite");
//		oauth2.put_Scope(
//				"openid profile offline_access https://outlook.office365.com/SMTP.Send https://outlook.office365.com/POP.AccessAsUser.All https://outlook.office365.com/IMAP.AccessAsUser.All ");
		// oauth2.put_Scope("openid profile offline_access
		// https://outlook.office365.com/SMTP.Send
		// https://outlook.office365.com/POP.AccessAsUser.All
		// https://outlook.office365.com/IMAP.AccessAsUser.All");
//		oauth2.put_Scope("openid profile offline_access user.readwrite mail.readwrite mail.send files.readwrite https://outlook.office.com/EWS.AccessAsUser.All");
		oauth2.put_Scope(
				"openid profile offline_access https://outlook.office365.com/SMTP.Send https://outlook.office365.com/POP.AccessAsUser.All https://outlook.office365.com/IMAP.AccessAsUser.All https://outlook.office365.com/EWS.AccessAsUser.All");
		oauth2.put_RedirectAllowHtml(
				"<div style='text-align: center; display: flex; justify-content: center!important; align-items: center; height: 100%;'><div style='flex: 0 0 33.333333%; max-width: 33.333333%; position: relative; width: 100%; background: #f1f9ff; border-radius: 20px; padding: 1rem; border: 1px solid #ddd;'><img src='https://www.arysontechnologies.com/assets/img/arysonlogo.png' alt='logo'><p style='font-size: 27px;color: #19a300;'>Authentication Successful</p><p>Please, go back to the "+All_Data.messageboxtitle+" !</p></div></div>");
		oauth2.put_RedirectDenyHtml(
				"<div style='text-align: center; display: flex; justify-content: center!important; align-items: center; height: 100%;'><div style='flex: 0 0 33.333333%; max-width: 33.333333%; position: relative; width: 100%; background: #f1f9ff; border-radius: 20px; padding: 1rem; border: 1px solid #ddd;'><img src='https://www.arysontechnologies.com/assets/img/arysonlogo.png' alt='logo'><p style='font-size: 27px;color: #ff0000;'>Access Denied</p><p>Please, go back to the "+All_Data.messageboxtitle+" !</p></div></div>");


		String url = oauth2.startAuth();// StartAuth();
		if (oauth2.get_LastMethodSuccess() != true) {
			System.out.println(oauth2.lastErrorText());
			return false;
		}

		try {
			Desktop.getDesktop().browse(new URI(url));
		} catch (IOException e) {

			e.printStackTrace();
		} catch (URISyntaxException e) {

			e.printStackTrace();
		}

		int numMsWaited = 0;
		while ((numMsWaited < 30000) && (oauth2.get_AuthFlowState() < 3)) {
			oauth2.SleepMs(100);
			numMsWaited = numMsWaited + 100;
		}

		if (oauth2.get_AuthFlowState() < 3) {
			oauth2.Cancel();
			System.out.println("No response from the browser!");
			return false;
		}

		if (oauth2.get_AuthFlowState() == 5) {
			System.out.println("OAuth2 failed to complete.");
			System.out.println(oauth2.failureInfo());

			return false;
		}

		if (oauth2.get_AuthFlowState() == 4) {
			System.out.println("93.. OAuth2 authorization was denied.");
			System.out.println("94.." + oauth2.lastErrorText());

			return false;
		}

		if (oauth2.get_AuthFlowState() != 3) {
			System.out.println("Unexpected AuthFlowState:" + oauth2.get_AuthFlowState());

			return false;
		}

		CkStringBuilder sbJson = new CkStringBuilder();
		sbJson.Append(oauth2.accessTokenResponse());

		File f = new File(System.getProperty("java.io.tmpdir") + File.separator + "ost_pst");
		if (!f.exists()) {
			f.mkdir();
		} else {
			System.out.println("icloud");
		}
		sbJson.WriteFile(f.getAbsolutePath() + File.separator + "ost_pst.json", "utf-8", false);

		System.out.println("OAuth2 authorization granted!");

		System.out.println("Access Token = " + oauth2.accessToken());
		//oauth2.pa
		json = new CkJsonObject();
		json.Load(oauth2.accessTokenResponse());
		json.put_EmitCompact(false);

		if (json.HasMember("expires_on") != true) {
			CkDateTime dtExpire = new CkDateTime();
			dtExpire.SetFromCurrentSystemTime();
			dtExpire.AddSeconds(json.IntOf("expires_in"));
			json.AppendString("expires_on", dtExpire.getAsUnixTimeStr(false));
		}
		System.out.println(json.emit());
		token=json.stringOf("access_token");


//
//           CkHttp http=new CkHttp();
//		http.put_AuthToken(token);
//
//		String resp = http.quickGetStr("https://graph.microsoft.com/v1.0/me");
//	
//	    if (http.get_LastMethodSuccess() != true) {
//	    	
//	        System.out.println("149..."+http.lastErrorText());
//	      
//	        }
//	    CkJsonObject json1 = new CkJsonObject();
//	    json1.put_EmitCompact(false);
//	    json1.Load(resp);
//	    System.out.println(json1.emit());
//username=json1.stringOf("mail");
return true;
	}


	
	public static IEWSClient conntiontooffice365_output() throws Exception {
		boolean auth1=auth();
		
		if (auth1) {
			EWSClient.useSAAJAPI(true);
			NetworkCredential credentials = new OAuthNetworkCredential(token);
			main_multiplefile.clientforexchange_output = EWSClient
					.getEWSClient("https://outlook.office365.com/ews/exchange.asmx", credentials);
			main_multiplefile.clientforexchange_output.setTimeout(3 * 60 * 1000);
			EmailClient.setSocketsLayerVersion2(true);
			EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
			System.out.println("Connection Done 1: ");
		}
		else {
			throw new DemoException("authorization denied!");
		}
		return main_multiplefile.clientforexchange_output;
	}
	@SuppressWarnings("deprecation")
	public static IEWSClient conntiontooffice365_output1() throws Exception {
		
		
		
		
		try {
			EWSClient.useSAAJAPI(true);
NetworkCredential credentials = new OAuthNetworkCredential(token);
main_multiplefile.clientforexchange_output = EWSClient.getEWSClient("https://outlook.office365.com/ews/exchange.asmx",credentials);
				
			
			
			main_multiplefile.clientforexchange_output.setTimeout(3 * 60 * 1000);
			EmailClient.setSocketsLayerVersion2(true);
			EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
			System.out.println("Connection Done : ");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			if(e.getMessage().contains("Bad response")|| e.getMessage().contains("Timeout")||e.getMessage().contains("Authentication failed")){
				   oauth2.put_RefreshToken(json.stringOf("refresh_token"));

				    // Send the HTTP POST to refresh the access token..
				 boolean   success = oauth2.RefreshAccessToken();
				    if (success != true) {
				        System.out.println(oauth2.lastErrorText());
				       
				        }

				    System.out.println("New access token: " + oauth2.accessToken());
				    System.out.println("New refresh token: " + oauth2.refreshToken());

				    // Update the JSON with the new tokens.
				    json.UpdateString("access_token",oauth2.accessToken());
				    json.UpdateString("refresh_token",oauth2.refreshToken());
					token=json.stringOf("access_token");
					EWSClient.useSAAJAPI(true);
					NetworkCredential credentials = new OAuthNetworkCredential(token);
					main_multiplefile.clientforexchange_output = EWSClient.getEWSClient("https://outlook.office365.com/ews/exchange.asmx",credentials);
									
								
								
								main_multiplefile.clientforexchange_output.setTimeout(3 * 60 * 1000);
								EmailClient.setSocketsLayerVersion2(true);
								EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
								System.out.println("Connection Done : new acess token");
			
			}
			throw new DemoException("Operation failed");
		}
		return main_multiplefile.clientforexchange_output;
	}
	public static String refresh() {

		   oauth2.put_RefreshToken(json.stringOf("refresh_token"));

		    // Send the HTTP POST to refresh the access token..
		 boolean   success = oauth2.RefreshAccessToken();
		    if (success != true) {
		        System.out.println(oauth2.lastErrorText());
		       
		        }

		    System.out.println("New access token: " + oauth2.accessToken());
		    System.out.println("New refresh token: " + oauth2.refreshToken());

		    // Update the JSON with the new tokens.
		    json.UpdateString("access_token",oauth2.accessToken());
		    json.UpdateString("refresh_token",oauth2.refreshToken());
			token=json.stringOf("access_token");
	return token;
	}

}
