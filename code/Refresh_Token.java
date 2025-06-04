package email.code;

import java.io.File;

import com.chilkatsoft.CkFileAccess;
import com.chilkatsoft.CkJsonObject;
import com.chilkatsoft.CkOAuth2;
import com.chilkatsoft.CkStringBuilder;

public class Refresh_Token {
	static String access_token;

	public static String refreshinput() {

		CkJsonObject jsonToken = new CkJsonObject();
		boolean success = jsonToken.LoadFile(System.getProperty("java.io.tmpdir") + File.separator
				+ "Tenate_To_Tenate_Office365" + File.separator + "input.json");
		if (success != true) {
			System.out.println("Failed to load office365.json");

		}
		File file = new File(System.getProperty("java.io.tmpdir") + File.separator + "Tenate_To_Tenate_Office365"
				+ File.separator + "input.json");
		access_token = refreshConnection(jsonToken, success, file);
		return access_token;

	}

	// Gmail
	public static String refreshoutput_gmail(File file) {

		CkJsonObject jsonToken = new CkJsonObject();
		boolean success = jsonToken.LoadFile(file.getAbsolutePath() + File.separator + "TokenGmail_Out.json");
		if (success != true) {
			System.err.println("Failed to load googleContacts.json");
		}
		access_token = refreshConnection_Gmail(jsonToken, success, file);
		return access_token;

	}

	// 365
	public static String refreshoutput() {

		CkJsonObject jsonToken = new CkJsonObject();
		boolean success = jsonToken.LoadFile(System.getProperty("java.io.tmpdir") + File.separator
				+ "Tenate_To_Tenate_Office365" + File.separator + "output.json");
		if (success != true) {
			System.out.println("Failed to load office365.json");

		}
		File file = new File(System.getProperty("java.io.tmpdir") + File.separator + "Tenate_To_Tenate_Office365"
				+ File.separator + "output.json");
		access_token = refreshConnection(jsonToken, success, file);
		return access_token;

	}

	// 365
	public static String refreshConnection(CkJsonObject jsonToken, boolean success, File file) {
		try {

			CkOAuth2 oauth2 = new CkOAuth2();
			oauth2.put_ListenPort(3017);
			oauth2.put_AuthorizationEndpoint("https://login.microsoftonline.com/common/oauth2/v2.0/authorize");
			oauth2.put_TokenEndpoint("https://login.microsoftonline.com/common/oauth2/v2.0/token");
			oauth2.put_ClientId("1994b9da-99f0-4c6b-9f78-75c47976d339");
			oauth2.put_ClientSecret("Ub4.5yQl-243bYQyfLx.W4MM-w3rfMom8-");

			oauth2.put_RefreshToken(jsonToken.stringOf("refresh_token"));

			success = oauth2.RefreshAccessToken();
			if (success != true) {
//				System.out.println(oauth2.lastErrorText());
			}

			if (success) {
				System.out.println("OAuth2 authorization granted!");
				jsonToken.UpdateString("access_token", oauth2.accessToken());
				jsonToken.UpdateString("refresh_token", oauth2.refreshToken());
				CkStringBuilder sbJson = new CkStringBuilder();
				jsonToken.put_EmitCompact(false);
				jsonToken.EmitSb(sbJson);
				CkJsonObject json = new CkJsonObject();
				json.Load(oauth2.accessTokenResponse());
				json.put_EmitCompact(false);
				CkFileAccess fac = new CkFileAccess();
				fac.WriteEntireTextFile(file.getAbsolutePath(), json.emit(), "utf-8", false);
			}
			access_token = oauth2.accessToken();
			System.out.println("================================");
			System.out.println(access_token);
			System.out.println("================================");
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return access_token;

	}

	// Gmail
	public static String refreshConnection_Gmail(CkJsonObject jsonToken, boolean success, File file) {
		try {
			CkOAuth2 oauth2 = new CkOAuth2();
			String id = "394398039123-m4akagfftghv055p2u3vf59qodra7k8r.apps.googleusercontent.com";
			String secret = "_JkVQlZQ1ec7pSC6l2YBEoa2";
			oauth2.put_TokenEndpoint("https://oauth2.googleapis.com/token");
			oauth2.put_ClientId(id);
			oauth2.put_ClientSecret(secret);
			oauth2.put_RefreshToken(jsonToken.stringOf("refresh_token"));
			success = oauth2.RefreshAccessToken();
			if (success != true) {
//				System.out.println(oauth2.lastErrorText());
			}

			//
			if (success) {
				jsonToken.UpdateString("access_token", oauth2.accessToken());
				CkStringBuilder sbJson = new CkStringBuilder();
				jsonToken.put_EmitCompact(false);
				jsonToken.EmitSb(sbJson);
				System.out.println("OAuth2 authorization granted!");
				CkJsonObject json = new CkJsonObject();
				json.Load(oauth2.accessTokenResponse());
				json.put_EmitCompact(false);
				CkFileAccess fac = new CkFileAccess();
				fac.WriteEntireTextFile(file.getAbsolutePath(), json.emit(), "utf-8", false);
			}
			access_token = oauth2.accessToken();
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return access_token;

	}

}