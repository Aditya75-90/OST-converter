package email.code;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;

import com.chilkatsoft.CkCache;
import com.chilkatsoft.CkDateTime;
import com.chilkatsoft.CkFileAccess;
import com.chilkatsoft.CkGlobal;
import com.chilkatsoft.CkHttp;
import com.chilkatsoft.CkJsonObject;
import com.chilkatsoft.CkOAuth2;
import com.chilkatsoft.CkStringBuilder;

public class One {
	static String token;
	static CkGlobal glob;
	static CkOAuth2 oauth2;

	public static String Chilkat_Connection() throws IOException, URISyntaxException {
		boolean success1;

		glob = new CkGlobal();
		success1 = glob.UnlockBundle("SNKRWT.CB1122022_Q5uG5AzJlRm1");
		if (success1 != true) {
			System.out.println(glob.lastErrorText());
			return null;
		}

		int status = glob.get_UnlockStatus();
		if (status == 2) {
			System.out.println("Unlocked using purchased unlock code.");
		} else {
			System.out.println("Unlocked in trial mode.");
		}
		System.out.println(success1 + " 42");

		oauth2 = new CkOAuth2();
		oauth2.put_ListenPort(3017);
		oauth2.put_AuthorizationEndpoint("https://login.microsoftonline.com/common/oauth2/v2.0/authorize");
		oauth2.put_TokenEndpoint("https://login.microsoftonline.com/common/oauth2/v2.0/token");
		oauth2.put_ClientId("1994b9da-99f0-4c6b-9f78-75c47976d339");
		oauth2.put_ClientSecret("Ub4.5yQl-243bYQyfLx.W4MM-w3rfMom8-");

		oauth2.put_CodeChallenge(false);
//		oauth2.put_Scope(
//				"openid profile offline_access https://outlook.office365.com/SMTP.Send https://outlook.office365.com/POP.AccessAsUser.All https://outlook.office365.com/IMAP.AccessAsUser.All");
		oauth2.put_Scope(
				"openid profile offline_access https://outlook.office365.com/SMTP.Send https://outlook.office365.com/POP.AccessAsUser.All https://outlook.office365.com/IMAP.AccessAsUser.All https://outlook.office365.com/EWS.AccessAsUser.All");

		oauth2.put_RedirectAllowHtml(
				"<div style='text-align: center; display: flex; justify-content: center!important; align-items: center; height: 100%;'><div style='flex: 0 0 33.333333%; max-width: 33.333333%; position: relative; width: 100%; background: #f1f9ff; border-radius: 20px; padding: 1rem; border: 1px solid #ddd;'><img src='https://www.arysontechnologies.com/assets/img/arysonlogo.png' alt='logo'><p style='font-size: 27px;color: #19a300;'>Authentication Successful</p><p>Please, go back to the "+All_Data.messageboxtitle+" !</p></div></div>");
		oauth2.put_RedirectDenyHtml(
				"<div style='text-align: center; display: flex; justify-content: center!important; align-items: center; height: 100%;'><div style='flex: 0 0 33.333333%; max-width: 33.333333%; position: relative; width: 100%; background: #f1f9ff; border-radius: 20px; padding: 1rem; border: 1px solid #ddd;'><img src='https://www.arysontechnologies.com/assets/img/arysonlogo.png' alt='logo'><p style='font-size: 27px;color: #ff0000;'>Access Denied</p><p>Please, go back to the "+All_Data.messageboxtitle+" !</p></div></div>");

		String url = oauth2.startAuth();
		if (oauth2.get_LastMethodSuccess() != true) {
			System.out.println(oauth2.lastErrorText());
			return null;
		}

		Desktop.getDesktop().browse(new URI(url));
		int numMsWaited = 0;
		while ((numMsWaited < 3000000) && (oauth2.get_AuthFlowState() < 3)) {
			oauth2.SleepMs(100);
			numMsWaited = numMsWaited + 100;
		}
		if (oauth2.get_AuthFlowState() < 3) {
			oauth2.Cancel();
			System.out.println("No response from the browser!");
			return null;
		}
		if (oauth2.get_AuthFlowState() == 5) {
			System.out.println("OAuth2 failed to complete.");
			System.out.println(oauth2.failureInfo());
			return null;
		}
		if (oauth2.get_AuthFlowState() == 4) {
			System.out.println("OAuth2 authorization was denied.");
			System.out.println(oauth2.accessTokenResponse());
			return null;
		}
		if (oauth2.get_AuthFlowState() != 3) {
			System.out.println("Unexpected AuthFlowState:" + oauth2.get_AuthFlowState());
			return null;
		}
		System.out.println();
		System.out.println("OAuth2 authorization granted!");
//		System.out.println("Access Token = " + oauth2.accessToken());
		token = oauth2.accessToken();

		CkJsonObject json = new CkJsonObject();
		json.Load(oauth2.accessTokenResponse());
		json.put_EmitCompact(false);
		if (json.HasMember("expires_on") != true) {
			CkDateTime dtExpire = new CkDateTime();
			dtExpire.SetFromCurrentSystemTime();
			dtExpire.AddSeconds(json.IntOf("expires_in"));
			json.AppendString("expires_on", dtExpire.getAsUnixTimeStr(false));
		}
//		System.out.println(json.emit());
		CkFileAccess fac = new CkFileAccess();
		File file = new File(System.getProperty("java.io.tmpdir") + File.separator + "Tenate_To_Tenate_Office365");
		file.mkdir();
		fac.WriteEntireTextFile(System.getProperty("java.io.tmpdir") + File.separator + "Tenate_To_Tenate_Office365"
				+ File.separator + "input.json", json.emit(), "utf-8", false);

		return token;

	}

}
