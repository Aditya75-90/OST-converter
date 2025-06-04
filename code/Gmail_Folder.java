package email.code;

import java.io.File;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import javax.swing.ImageIcon;

import com.aspose.email.EmailClient;
import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.IConnection;
import com.aspose.email.IEWSClient;
import com.aspose.email.ImapClient;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.PersonalStorage;
import com.aspose.email.SecurityOptions;

public class Gmail_Folder implements Runnable {

	String path4 = "";
	String x1 = "";
	static String calendertime;
	static Calendar cal;
	String parentfolder;
	PersonalStorage ost;
	int splitcount = 0;
	String splitpath = "";
	long foldermessagecount;
	List<String> listduplicacy = new ArrayList<String>();
	static Date fromdate;
	static Date todate;
	String path3 = "";
	ArrayList<Date> fromList = new ArrayList<Date>();
	ArrayList<Date> toList = new ArrayList<Date>();
	String from;
	String to;
	String first = null, middle = null, last = null;
	private Main_Frame mf;
	String path = "";
	private main_multiplefile mm;
	private String filetype = "";
	private String filepath = "";
	private String destination_path = "";
	static long count_destination;
	static PersonalStorage pst;
	String Folder;
	List<String> pstfolderlist;
	static boolean datevalidflag = false;
	String temppathm = "";
	IEWSClient clientforexchange_output;
	ImapClient clientforimap_output;
	IConnection iconnforimap_output;
	String path1 = "";
	long maxsize = 0;
	int pstindex = 0;
	String Folderuri;
	String fa = "";
	static String mailboxUri = "https://outlook.office365.com/EWS/Exchange.asmx";
	String username_p3;
	String match;
	String password_p3;
	String parent, parent1 = "";

	public Gmail_Folder(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm, String username_p3, String password_p3, String Folderuri,
			ImapClient clientforimap_output, IConnection iconnforimap_output, String path, String fname, String match,
			String domain_p, int portnofiletype) {

		this.mf = mf;
		this.mm = mm;
		this.filetype = filetype;
		this.filepath = filepath;
		this.destination_path = destination_path;
		Gmail_Folder.count_destination = count_destination;
		this.pstfolderlist = pstfolderlist;
		this.fromList = fromList;
		this.toList = toList;
		this.temppathm = temppathm;
		this.username_p3 = username_p3;
		this.password_p3 = password_p3;
		this.Folderuri = Folderuri;
		this.clientforimap_output = clientforimap_output;
		this.iconnforimap_output = iconnforimap_output;
		this.path = path;
	}

	@Override
	public void run() {

		Gmail_Folder(mf, filetype, destination_path, count_destination, filepath, mm, pstfolderlist, fromList, toList,
				temppathm, Folderuri, clientforimap_output, iconnforimap_output, path);
		main_multiplefile.count_destination = Gmail_Folder.count_destination;
	}

	private void Gmail_Folder(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm, String folderuri, ImapClient clientforimap_output,
			IConnection iconnforimap_output, String path) {
		Gmail_Folder.count_destination = 0;
		match = path;
		if (filetype.equalsIgnoreCase("GoDaddy email")) {
			main_multiplefile.fname = main_multiplefile.fname.replaceAll("[^a-zA-Z0-9]", "");

		}
		path4 = path;
		path = path + "/" + main_multiplefile.fname;
		parent = path;
		System.out.println(path + "  Path ");
		clientforimap_output.createFolder(iconnforimap_output, path);
		clientforimap_output.selectFolder(iconnforimap_output, path);
		pst = PersonalStorage.fromFile(filepath);

		FolderInfo folderInfo2 = pst.getRootFolder();
		String Folder = folderInfo2.getDisplayName();
		Folder = Folder.replace(",", "").replace(".", "");
		Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
		Folder = Folder.replaceAll("[\\[\\]]", "");
		Folder = Folder.trim();
		if (Folder.equalsIgnoreCase("")) {
			Folder = "Root Folder";
		}
		if (filetype.equalsIgnoreCase("GoDaddy email")) {
			Folder = Folder.replaceAll("[^a-zA-Z0-9]", "");

		}
		String path1 = Folder;
//
		parent1 = path + "/" + Folder;
		clientforimap_output.createFolder(iconnforimap_output, parent1);
		clientforimap_output.selectFolder(iconnforimap_output, parent1);
		System.out.println(parent1 + "  141 ");
		//
		FolderInfoCollection folderInf = pst.getRootFolder().getSubFolders();

		for (int j = 0; j < folderInf.size(); j++) {
			try {

				if (mm.stop) {
					break;
				}
				FolderInfo folderInfo = folderInf.get_Item(j);
				String Folder1 = folderInfo.getDisplayName();
				Folder1 = Folder1.replace(",", "").replace(".", "");
				Folder1 = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder1);
				Folder1 = Folder1.replaceAll("[\\[\\]]", "");
				Folder1 = Folder1.trim();

				Folder = path1 + File.separator + Folder1;
				String sfolder = Folder;
				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}
					if (pstfolderlist.get(l).equalsIgnoreCase(Folder)) {
						mm.lbl_progressreport.setText("Getting Folder " + Folder);

						String fol = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");

						path = parent1 + "/" + fol;
						System.out.println(path + "  179 ");
						if (clientforimap_output.existFolder(path)) {
							clientforimap_output.selectFolder(path);
						} else {
							clientforimap_output.createFolder(iconnforimap_output, path);
							clientforimap_output.selectFolder(iconnforimap_output, path);
						}

					}
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderforpstost_gmail(folderInfo, sfolder, path);
				}
				path = mm.removefoldergmail(path);
			} catch (Exception e) {
				continue;
			}

		}

	}

	private void getsubfolderforpstost_gmail(FolderInfo f, String sfolder, String path1) {

		FolderInfoCollection subfolder = f.getSubFolders();
		String path11 = "";
		for (int k = 0; k < subfolder.size(); k++) {
			try {
				if (mm.stop) {
					break;
				}
				FolderInfo folderf = subfolder.get_Item(k);

				String Folder = folderf.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();

				sfolder = sfolder + File.separator + Folder;
				System.out.println(sfolder + "  223 ");
				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}

					if (pstfolderlist.get(l).equalsIgnoreCase(sfolder)) {

						//
						String new_path = path4 + "\\" + main_multiplefile.fname + "\\" + sfolder;
						System.out.println(new_path);
						mm.lbl_progressreport.setText("Getting : " + Folder);
						String[] p1 = new_path.split("\\\\");
						x1 = path4;
						for (int i1 = 1; i1 < p1.length; i1++) {

							x1 = x1 + "/" + p1[i1];

							path = x1;

							try {
								if (clientforimap_output.existFolder(x1)) {
									clientforimap_output.selectFolder(x1);
								} else {
									clientforimap_output.createFolder(x1);
									clientforimap_output.selectFolder(x1);
								}
							} catch (Exception e) {
							}
						}

						//

						mm.lbl_progressreport.setText("Getting : " + Folder);
					}
				}
				if (folderf.hasSubFolders()) {

					getsubfolderforpstost_gmail(folderf, sfolder, path11);
				}

				path = mm.removefoldergmail(path);
				sfolder = mm.removefolder(sfolder);
			} catch (Exception e) {
				continue;
			}
		}

	}

	public void connectionHandle() {
		System.out.println("Connection Lost We are trying to Connect using of Connection Class.");
		mm.lbl_progressreport.setText("INTERNET Connection  LOST ");
		System.out.println("Trying to  Connect with Sachin  :" + Calendar.getInstance().getTime());
		mm.label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));

		while (true) {
			try {
				mm.lbl_progressreport.setText("Connecting to Server Please Wait");
				if (filetype.equalsIgnoreCase("GMAIL")) {
					if (main_multiplefile.modern_Authentication.isSelected()) {
						String token = GetToken.refreshToken_Gmail_Output();
						if (token != null) {
							clientforimap_output.dispose();
							clientforimap_output = GetToken.loginGmail_output(token);
						}
					} else {
						clientforimap_output.dispose();
						clientforimap_output = connectiontogmail_output();
					}

				}
				mm.label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));
				mm.lbl_progressreport.setText("Connection extablished Retriving Messasge");
				break;
			} catch (Exception e) {
				mm.lbl_progressreport.setText("INTERNET Connection  LOST ");

			}

		}

		mm.Progressbar.setVisible(true);

	}

	// Gmail
	@SuppressWarnings("deprecation")
	public ImapClient connectiontogmail_output() throws Exception {
		clientforimap_output = new ImapClient("imap.gmail.com", 993, mm.username_p3, mm.password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	// Yahoo
	@SuppressWarnings("deprecation")
	public ImapClient connectiontoyahoo_output(ImapClient clientforimap_output) throws Exception {
		clientforimap_output = new ImapClient("imap.mail.yahoo.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	// Aol
	@SuppressWarnings("deprecation")
	public ImapClient connectiontoaol_output(ImapClient clientforimap_output) throws Exception {
		clientforimap_output = new ImapClient("imap.aol.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

//		clientforimap_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

//	// Amazon
//	@SuppressWarnings("deprecation")
//	public ImapClient connectiontoinaws_output(ImapClient clientforimap_output) throws Exception {
//		clientforimap_output = new ImapClient(domain_p3, portnofiletype, username_p3, password_p3);
//
//		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);
//
//		EmailClient.setSocketsLayerVersion2(true);
//		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
////		clientforimap_output.setTimeout(5 * 60 * 1000);
//		iconnforimap_output = clientforimap_output.createConnection();
//		return clientforimap_output;
//	}

	// Icloud
	@SuppressWarnings("deprecation")
	public ImapClient connectiontoicloud_output() throws Exception {
		clientforimap_output = new ImapClient("imap.mail.me.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		// clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	// Zoho
	@SuppressWarnings("deprecation")
	public ImapClient connectiontozoho_output() throws Exception {
		clientforimap_output = new ImapClient("imap.zoho.in", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		// clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	@SuppressWarnings("deprecation")
	// Yandex
	public ImapClient connectiontoYandex_output() throws Exception {
		clientforimap_output = new ImapClient("imap.yandex.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		// clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public void connectionHandle1() {
		mm.label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));
		while (true) {
			try {
				if (filetype.equalsIgnoreCase("GMAIL")) {
//					if (main_multiplefile.modern_Authentication.isSelected()) {
//						String token = GetToken.refreshToken_Gmail_Output();
//						if (token != null) {
//							clientforimap_output.dispose();
//							clientforimap_output = GetToken.loginGmail_output(token);
//						}
//					} else {
					clientforimap_output.dispose();
					clientforimap_output = connectiontogmail_output();
//					}

				} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {
					clientforimap_output.dispose();
					clientforimap_output = connectiontoyahoo_output(clientforimap_output);

				} else if (filetype.equalsIgnoreCase("AOL")) {
					clientforimap_output.dispose();
					clientforimap_output = connectiontoaol_output(clientforimap_output);
				}

				mm.label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));
				break;

			} catch (Exception e) {
				e.printStackTrace();
			}
		}

	}

}
