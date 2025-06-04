package email.code;

import java.awt.Desktop;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URISyntaxException;
import java.util.logging.FileHandler;
import java.util.logging.SimpleFormatter;

import javax.swing.ImageIcon;
import javax.swing.JOptionPane;

import com.aspose.email.EWSClient;
import com.aspose.email.ExchangeFolderInfo;
import com.aspose.email.ExchangeFolderInfoCollection;
import com.aspose.email.ExchangeMailboxInfo;
import com.aspose.email.IConnection;
import com.aspose.email.IEWSClient;
import com.aspose.email.ImapClient;
import com.aspose.email.ListFoldersOptions;
import com.aspose.email.OAuthNetworkCredential;
import com.aspose.email.SecurityOptions;
import com.aspose.email.system.NetworkCredential;
import com.chilkatsoft.CkAuthGoogle;
import com.chilkatsoft.CkGlobal;
import com.chilkatsoft.CkJsonObject;
import com.chilkatsoft.CkOAuth2;
//import com.sun.org.apache.xerces.internal.util.URI;
import com.chilkatsoft.CkRest;
import com.chilkatsoft.CkStringBuilder;

import email.activation.Starting_Frame;

public class GetToken {
	All_Data st = new All_Data();
	static Main_Frame mf;
	static {
		if (System.getProperty("os.name").toLowerCase().contains("windows")) {
			FileHandler fh;
			try {
				fh = new FileHandler(System.getProperty("java.io.tmpdir") + File.separator + "chilkat.log");
				// logger.addHandler(fh);
				SimpleFormatter formatter = new SimpleFormatter();
				fh.setFormatter(formatter);
				// logger.info("My Log File");
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
			try {
				InputStream in = Main_Frame.class.getResourceAsStream("/chilkat.dll");
				byte[] buffer = new byte[1024];
				int read = -1;
				File temp = new File(new File(System.getProperty("java.io.tmpdir")), "chilkat.dll");
				int i = 0;
				while (temp.exists()) {
					temp = new File(new File(System.getProperty("java.io.tmpdir")), "chilkat" + i + ".dll");
					i++;
				}
				FileOutputStream fos = null;
				try {
					fos = new FileOutputStream(temp);
				} catch (FileNotFoundException e) {
					e.printStackTrace();
					// logger.warning(e.getMessage());
				}
				try {
					while ((read = in.read(buffer)) != -1) {
						fos.write(buffer, 0, read);
					}
				} catch (IOException e) {
					// logger.warning(e.getMessage());
				}
				try {
					fos.close();
				} catch (IOException e) {
					// logger.warning(e.getMessage());
				}
				try {
					in.close();
				} catch (IOException e) {
					// logger.warning(e.getMessage());
				}
				try {
					System.load(temp.getAbsolutePath());
				} catch (Error er) {
					in = Main_Frame.class.getResourceAsStream("/chilkatX64.dll");
					buffer = new byte[1024];
					read = -1;
					temp = new File(new File(System.getProperty("java.io.tmpdir")), "chilkatX64.dll");
					i = 0;
					while (temp.exists()) {
						temp = new File(new File(System.getProperty("java.io.tmpdir")), "chilkatX64" + i + ".dll");
						i++;
					}
					fos = null;
					try {
						fos = new FileOutputStream(temp);
					} catch (FileNotFoundException e) {
						e.printStackTrace();
						// logger.warning(e.getMessage());
					}
					try {
						while ((read = in.read(buffer)) != -1) {
							fos.write(buffer, 0, read);
						}
					} catch (IOException e) {
						// logger.warning(e.getMessage());
					}
					try {
						fos.close();
					} catch (IOException e) {
						// logger.warning(e.getMessage());
					}
					try {
						in.close();
					} catch (IOException e) {
						// logger.warning(e.getMessage());
					}
					System.load(temp.getAbsolutePath());
				}
			} catch (UnsatisfiedLinkError | Exception e) {
				e.printStackTrace();
				// logger.warning(e.getMessage());
				// System.loadLibrary("chilkat64");
				// logger.warning(e.getMessage());
				// System.loadLibrary("chilkat");
				// System.load("‪D:\\EXE\\chilkat-9.5.0-jdk11-x64\\chilkat.dll");
			}
		}else {
			// Logger logger = Logger.getLogger("MyLog1");
			FileHandler fh;
			try {
				
				fh = new FileHandler(System.getProperty("java.io.tmpdir") + File.separator + "chilkat.log");
				// logger.addHandler(fh);
				SimpleFormatter formatter = new SimpleFormatter();
				fh.setFormatter(formatter);
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
			try {
				System.out.println("605..");
				InputStream in = Main_Frame.class.getResourceAsStream("/libchilkat.jnilib");
				byte[] buffer = new byte[1024];
				int read = -1;
				File temp = new File(new File(System.getProperty("java.io.tmpdir")), "libchilkat.jnilib");
				System.out.println("610");
				int i = 0;
				while (temp.exists()) {
					temp.delete();
					temp = new File(new File(System.getProperty("java.io.tmpdir")), "libchilkat" + i + ".jnilib");
					i++;
				}
				FileOutputStream fos = null;
				try {
					fos = new FileOutputStream(temp);
				} catch (FileNotFoundException e) {
					e.printStackTrace();
					// logger.warning(e.getMessage());
				}
				try {
					while ((read = in.read(buffer)) != -1) {
						fos.write(buffer, 0, read);
					}
				} catch (IOException e) {
					// logger.warning(e.getMessage());
				}
				try {
					fos.close();
				} catch (IOException e) {
					// logger.warning(e.getMessage());
				}
				try {
					in.close();
				} catch (IOException e) {
					
					// logger.warning(e.getMessage());
				}
				System.out.println("640.."+temp.getAbsolutePath());
				System.load(temp.getAbsolutePath());
			
			temp.delete();	
			} catch (UnsatisfiedLinkError | Exception e) {
				e.printStackTrace();
				// logger.warning(e.getMessage());
				// System.loadLibrary("chilkat64");
				// logger.warning(e.getMessage());
				// System.loadLibrary("chilkat");
				// System.load("â€ªD:\\EXE\\chilkat-9.5.0-jdk11-x64\\chilkat.dll");
			}
		
		}
	}
	private static String token;
	static IConnection iconnforimap_output;
	IConnection iconnforimap_input;
	static CkOAuth2 oauth2 = null;
	static CkGlobal glob;
	static boolean success1;

	static public String tokenForGmail_output() throws IOException, URISyntaxException {

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

		oauth2 = new CkOAuth2();
		boolean success;

		oauth2.put_ListenPort(3017);

		oauth2.put_AuthorizationEndpoint("https://accounts.google.com/o/oauth2/v2/auth");
		oauth2.put_TokenEndpoint("https://oauth2.googleapis.com/token");
		String id = "394398039123-m4akagfftghv055p2u3vf59qodra7k8r.apps.googleusercontent.com";
		String secret = "_JkVQlZQ1ec7pSC6l2YBEoa2";
		oauth2.put_ClientId(id);
		oauth2.put_ClientSecret(secret);

		oauth2.put_CodeChallenge(true);
		oauth2.codeChallengeMethod();

		oauth2.put_Scope("https://www.googleapis.com/auth/drive https://mail.google.com");
		oauth2.put_RedirectAllowHtml(
				"<div style='text-align: center; display: flex; justify-content: center!important; align-items: center; height: 100%;'><div style='flex: 0 0 33.333333%; max-width: 33.333333%; position: relative; width: 100%; background: #F1F9FF; border-radius: 20px; padding: 1rem; border: 1px solid #ddd;'><img src='https://www.arysontechnologies.com/assets/img/arysonlogo.png' alt='logo'><p style='font-size: 27px;color: #19A300;'>Authentication Successful</p><p>Please, go back to the "+All_Data.messageboxtitle+"!</p></div></div>");
		oauth2.put_RedirectDenyHtml(
				"<div style='text-align: center; display: flex; justify-content: center!important; align-items: center; height: 100%;'><div style='flex: 0 0 33.333333%; max-width: 33.333333%; position: relative; width: 100%; background: #F1F9FF; border-radius: 20px; padding: 1rem; border: 1px solid #ddd;'><img src='https://www.arysontechnologies.com/assets/img/arysonlogo.png' alt='logo'><p style='font-size: 27px;color: #FF0000;'>Access Denied</p><p>Please, go back to the "+All_Data.messageboxtitle+"!</p></div></div>");
		String url = oauth2.startAuth();
		if (oauth2.get_LastMethodSuccess() != true) {
			System.err.println(oauth2.lastErrorText());
		}
		try {
			Desktop.getDesktop().browse(new java.net.URI(url));
		} catch (Exception e) {

		}

		int numMsWaited = 0;
		while ((numMsWaited < 90000) && (oauth2.get_AuthFlowState() < 3)) {
			oauth2.SleepMs(100);
			numMsWaited = numMsWaited + 100;
		}

		if (oauth2.get_AuthFlowState() < 3) {
			oauth2.Cancel();
			System.err.println("No response from the browser!");
			JOptionPane.showMessageDialog(mf, "No response from the browser!", All_Data.messageboxtitle,
					JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
//		        return;
		}

		else if (oauth2.get_AuthFlowState() == 5) {
			System.err.println("OAuth2 failed to complete.");
			System.err.println(oauth2.failureInfo());
			JOptionPane.showMessageDialog(mf, "OAuth2 failed to complete.", All_Data.messageboxtitle,
					JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
//		        return;
		}

		else if (oauth2.get_AuthFlowState() == 4) {
			System.err.println("OAuth2 authorization was denied.");
			System.err.println(oauth2.accessTokenResponse());
			JOptionPane.showMessageDialog(mf, "OAuth2 authorization was denied.", All_Data.messageboxtitle,
					JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
			
			
//		        return;
		}

		else if (oauth2.get_AuthFlowState() != 3) {
			// System.out.println("Unexpected AuthFlowState:" + oauth2.get_AuthFlowState());
			JOptionPane.showMessageDialog(mf, "Unexpected AuthFlowState:" + oauth2.get_AuthFlowState(),
					All_Data.messageboxtitle, JOptionPane.ERROR_MESSAGE,
					new ImageIcon(Main_Frame.class.getResource("/information.png")));
//		        return;
		}
		CkStringBuilder sbJson = new CkStringBuilder();
		sbJson.Append(oauth2.accessTokenResponse());
		String path;
		if (System.getProperty("os.name").toLowerCase().contains("windows")) {

			path = System.getenv("APPDATA") + File.separator + All_Data.messageboxtitle + File.separator + "Token";

		} else {

			path = System.getProperty("user.home") + File.separator + "Library" + File.separator + "Application Support"
					+ File.separator + All_Data.messageboxtitle + File.separator + "Token";
		}
		File file = new File(path);
		file.mkdirs();
		sbJson.WriteFile(file.getAbsolutePath() + File.separator + "TokenGmail_Out.json", "utf-8", false);

		return oauth2.accessToken();
	}

	public static ImapClient loginGmail_output(String accessToken) throws Exception {

		String username = userName_Out(accessToken);
		ImapClient clientforimap_output = new ImapClient("imap.gmail.com", 993, username, accessToken, true);
		clientforimap_output.setSecurityOptions(SecurityOptions.SSLAuto);
		clientforimap_output.setSecurityOptions(SecurityOptions.SSLImplicit);

		clientforimap_output.setTimeout(60 * 1000);

		iconnforimap_output = clientforimap_output.createConnection();

		return clientforimap_output;
	}

	public static String userName_Out(String accessToken) {

		String username = null;
		CkAuthGoogle gAuth = new CkAuthGoogle();
		gAuth.put_AccessToken(accessToken);

		CkRest rest = new CkRest();
		@SuppressWarnings("unused")
		boolean success3 = true;
		boolean bAutoReconnect = true;
		success3 = rest.Connect("www.googleapis.com", 443, true, bAutoReconnect);
		rest.SetAuthGoogle(gAuth);
		CkJsonObject json1 = new CkJsonObject();
		String jsonResponse1 = rest.fullRequestNoBody("GET", "/gmail/v1/users/me/profile");
		json1.Load(jsonResponse1);
		username = json1.stringOf("emailAddress");
		return username;
	}

	public static String refreshToken_Gmail_Output() {
		String path;
		if (System.getProperty("os.name").toLowerCase().contains("windows")) {

			path = System.getenv("APPDATA") + File.separator + All_Data.messageboxtitle + File.separator + "Token";

		} else {
			path = System.getProperty("user.home") + File.separator + "Library" + File.separator + "Application Support"
					+ File.separator + All_Data.messageboxtitle + File.separator + "Token";
		}
		File file = new File(path);
		String token = Refresh_Token.refreshoutput_gmail(file);
		return token;
	}

}
