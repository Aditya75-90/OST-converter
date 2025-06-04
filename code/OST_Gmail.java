package email.code;

import java.io.File;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import javax.swing.ImageIcon;

import com.aspose.email.AppointmentSaveFormat;
import com.aspose.email.Attachment;
import com.aspose.email.ContactSaveFormat;
import com.aspose.email.EmailClient;
import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.IConnection;
import com.aspose.email.IEWSClient;
import com.aspose.email.ImapClient;
import com.aspose.email.ImapMessageFlags;
import com.aspose.email.MailAddress;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiAttachmentCollection;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiContactElectronicAddress;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiElectronicAddress;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiMessageFlags;
import com.aspose.email.MapiTask;
import com.aspose.email.MessageInfo;
import com.aspose.email.MessageInfoCollection;
import com.aspose.email.PersonalStorage;
import com.aspose.email.SaveOptions;
import com.aspose.email.SecurityOptions;

public class OST_Gmail implements Runnable {

	boolean internet;
	private List<String> listdupliccal = new ArrayList<String>();
	private List<String> listduplictask = new ArrayList<String>();
	private List<String> listdupliccontact = new ArrayList<String>();
	private List<String> listduplicacy = new ArrayList<String>();
	static String calendertime;
	static Calendar cal;
	String parentfolder;
	PersonalStorage ost;
	int splitcount = 0;
	String splitpath = "";
	long foldermessagecount;
	File file11;
	static Date fromdate;
	static Date todate;
	String path3 = "";
	String path4 = "";
	ArrayList<Date> fromList = new ArrayList<Date>();
	ArrayList<Date> toList = new ArrayList<Date>();
	String from;
	String to;
	String first = null, middle = null, last = null;
	private Main_Frame mf;
	private String path = "";
	String x1 = "";
	private main_multiplefile mm;
	private String filetype = "";
	private String filepath = "";
	private String destination_path = "";
	long count_destination;
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
	String fname;
	String password_p3;
	String domain_p3;
	int portnofiletype;

	public OST_Gmail(Main_Frame mf, String filetype, String destination_path, long count_destination, String filepath,
			main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList, ArrayList<Date> toList,
			String temppathm, String username_p3, String password_p3, String Folderuri, ImapClient clientforimap_output,
			IConnection iconnforimap_output, String path, String fname, String match, String domain_p3,
			int portnofiletype) {

		this.mf = mf;
		this.mm = mm;
		this.filetype = filetype;
		this.filepath = filepath;
		this.destination_path = destination_path;
		this.count_destination = count_destination;
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
		this.fname = fname;
		this.match = match;
		this.domain_p3 = domain_p3;
		this.portnofiletype = portnofiletype;
	}

	@Override
	public void run() {
		convertPSTOST_gmail(mf, filetype, destination_path, count_destination, filepath, mm, pstfolderlist, fromList,
				toList, temppathm, Folderuri, clientforimap_output, iconnforimap_output, path, match, domain_p3,
				portnofiletype);
		main_multiplefile.count_destination = count_destination;
		main_multiplefile.match = match;
	}

	private void convertPSTOST_gmail(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm, String folderuri, ImapClient clientforimap_output,
			IConnection iconnforimap_output, String path, String match, String domain_p3, int portnofiletype) {
		clientforimap_output.dispose();
		connectionHandle();
		match = path;
		if (filetype.equalsIgnoreCase("GoDaddy email")) {
			main_multiplefile.fname = main_multiplefile.fname.replaceAll("[^a-zA-Z0-9]", "");

		}
		path4 = path;
		pst = PersonalStorage.fromFile(filepath);
		MailConversionOptions options = new MailConversionOptions();
		FolderInfo folderInfo2 = pst.getRootFolder();
		String Folder = folderInfo2.getDisplayName();
		Folder = Folder.replace(",", "").replace(".", "");
		Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
		Folder = Folder.replaceAll("[\\[\\]]", "");
		Folder = Folder.trim();
		if (Folder.equalsIgnoreCase("")) {
			Folder = "Root Folder";
		}
		System.out.println("136 :" + Folder);
		if (filetype.equalsIgnoreCase("GoDaddy email")) {
			Folder = Folder.replaceAll("[^a-zA-Z0-9]", "");

		}
		String path1 = Folder;
		path = path + "\\" + Folder;

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
						if (filetype.equalsIgnoreCase("GoDaddy email")) {
							fol = fol.replaceAll("[^a-zA-Z0-9]", "");
						}
						path = path + "\\" + fol;

//						path = path4 + "\\" + main_multiplefile.fname + "\\" + fol;
						mm.lbl_progressreport.setText("Getting : " + Folder);
						String[] p1 = path.split("\\\\");
						x1 = path4;
						for (int i1 = 1; i1 < p1.length; i1++) {

							// Zoho
							if (filetype.equalsIgnoreCase("Zoho Mail")) {
								if (p1[i1].equalsIgnoreCase("Inbox") || p1[i1].equalsIgnoreCase("Drafts")
										|| p1[i1].equalsIgnoreCase("Outbox"))
									p1[i1] = p1[i1] + "_" + "zoho";
							}
							// Yandex
							if (filetype.equalsIgnoreCase("Yandex Mail")) {
								x1 = x1 + "|" + p1[i1].substring(0, 4);
							} else {
								x1 = x1 + "/" + p1[i1];
							}
							path = x1;

							try {
								if (clientforimap_output.existFolder(x1)) {
									clientforimap_output.selectFolder(x1);
								} else {
									clientforimap_output.createFolder(x1);
									clientforimap_output.selectFolder(x1);
								}
							} catch (Exception e) {
								e.printStackTrace();
								if (e.getMessage().equals("Object has been disposed.")) {
									System.out.println(e.getMessage() + "  242");
									connectionHandle();
									i1--;
								} else if (e.getMessage().contains("Software caused connection abort: recv failed")
										|| e.getMessage().contains("Network is unreachable: connect")
										|| e.getMessage().contains("Operation failed.")
										|| e.getMessage().equalsIgnoreCase("The operation 'FetchMessage' terminated.")
										|| e.getMessage().equalsIgnoreCase("ConnectFailure")
										|| e.getMessage().equalsIgnoreCase("Operation has been canceled")
										|| e.getMessage().contains("Operation failed")
										|| e.getMessage().contains("Rate limit hit") || e.getMessage()
												.equalsIgnoreCase("The operation 'AppendMessage' terminated.")) {
									System.out.println(e.getMessage() + "  252 ");
									connectionHandle();
								} else {
									System.out.println(x1);
									System.out.println(e.getMessage() + "  254 ");
									System.out.println("Continue");
									continue;
								}

							}
						}
						listduplicacy.clear();
						listdupliccal.clear();
						listdupliccontact.clear();
						listduplictask.clear();

						MessageInfoCollection messageInfoCollection = null;
						try {
							messageInfoCollection = folderInfo.getContents();
						} catch (Exception e1) {

							e1.printStackTrace();
						}

						if (!(messageInfoCollection == null)) {

							boolean s22 = false;
							int messagesize;
							if (main_multiplefile.demo) {
								if (messageInfoCollection.size() <= All_Data.demo_count) {
									messagesize = messageInfoCollection.size();
								} else {
									messagesize = All_Data.demo_count;
								}

							} else {
								messagesize = messageInfoCollection.size();
							}

							for (int i = 0; i < messagesize; i++) {
								try {
									if (mm.stop) {
										break;
									}
									if ((i % 100) == 0) {
										System.gc();
									}
									if (count_destination != 0) {
										if ((count_destination % 500) == 0) {
											if (s22) {
												connectionHandle1();
											}
											s22 = true;
										}
									}
//									MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);
//									MapiMessage message1 = pst.extractMessage(messageInfo);
//									MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
//									MailConversionOptions de = new MailConversionOptions();
//									MailMessage mess = message1.toMailMessage(de);

									MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);
									MapiMessage message1 = pst.extractMessage(messageInfo);
									MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
									MailConversionOptions de = new MailConversionOptions();
									MailMessage mess1 = message1.toMailMessage(de);
									MapiMessage message = MapiMessage.fromMailMessage(mess1, d);

									MailMessage mess = message.toMailMessage(options);
									if (mm.chckbxMigrateOrBackup.isSelected()) {
										mess.getAttachments().clear();
										MapiAttachmentCollection m = message1.getAttachments();
										for (int j1 = 0; j1 < m.size(); j1++) {
											m.clear();
										}
									}
//									MapiMessage message = MapiMessage.fromMailMessage(mess, d);

									if (main_multiplefile.datefilter.isSelected()) {
										datevalidflag = mm.checkdate(message1, mess);
										System.out.println(datevalidflag);
									}
									if (message1.getMessageClass().equals("IPM.Task")
											|| message1.getMessageClass().equals("IPM.StickyNote")
											|| message1.getMessageClass().equals("IPM.Contact")
											|| message1.getMessageClass().equals("IPM.Appointment")
											|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {
										int bct = message1.getBodyType();
										if (bct == 0) {
											message1.setBody(mess1.getBody());
										} else {
											message1.setBody(mess1.getBody());
										}
									}
									if (message1.getMessageClass().equals("IPM.Contact")) {
										MailMessage mapi = new MailMessage();
										try {
											MapiContact con = (MapiContact) message1.toMapiMessageItem();
											try {
												mapi.setSubject(con.getSubject() + "_" + i);
											} catch (Exception e) {

												mapi.setSubject("");
											}
											try {
												mapi.setDate(mess.getDate());
											} catch (Exception e) {
												mapi.setDate(null);
											}

											try {
												message1.setSenderEmailAddress(mess.getFrom().toString());
												mapi.setFrom(mess.getFrom());
											} catch (Exception e) {
											}
											for (Attachment attachment : mess.getAttachments()) {
												attachment.save(temppathm + File.separator + attachment.getName());
												file11 = new File(temppathm + File.separator + attachment.getName());
												mapi.addAttachment(new Attachment(
														temppathm + File.separator + attachment.getName()));
												file11.delete();
											}

											con.save(temppathm + File.separator + mm.namingconventionmapi(message1)
													+ "_" + i + ".vcf", ContactSaveFormat.VCard);
											File file = new File(temppathm + File.separator
													+ mm.namingconventionmapi(message1) + "_" + i + ".vcf");
											mapi.addAttachment(new Attachment(temppathm + File.separator
													+ mm.namingconventionmapi(message1) + "_" + i + ".vcf"));
											file.delete();

											if (mm.chckbxRemoveDuplicacy.isSelected()) {

												String input = mm.duplicacymapiContact(con);
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listdupliccontact.contains(input)) {
													System.out.println("Not a duplicate message");
													listdupliccontact.add(input);

													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforimap_output.appendMessage(iconnforimap_output,
																	path, mapi);
															count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mapi);
														count_destination++;
													}
												} else {
													continue;
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mapi);
														count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													count_destination++;
												}
											}
										} catch (Exception e) {
											if (e.getMessage()
													.equalsIgnoreCase("The operation 'FetchMessage' terminated.")
													|| e.getMessage().equalsIgnoreCase("ConnectFailure")
													|| e.getMessage().equalsIgnoreCase("Operation has been canceled")
													|| e.getMessage().contains("Operation failed")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage().equalsIgnoreCase(
															"The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Network is unreachable: connect")) {
												mm.Progressbar.setVisible(false);
												i--;

											}
											connectionHandle();
											mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
													+ System.lineSeparator());
											continue;
										}

									} else if (message1.getMessageClass().equals("IPM.Appointment")
											|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {
										MailMessage mapi = new MailMessage();
										try {
											MapiCalendar cal = null;
											try {
												cal = (MapiCalendar) message1.toMapiMessageItem();
											} catch (Exception e1) {
												continue;
											}
											try {
												mapi.setSubject(cal.getSubject() + "_" + i);
											} catch (Exception e) {
												mapi.setSubject("");
											}
											try {
												mapi.setDate(mess.getDate());
											} catch (Exception e) {
												mapi.setDate(null);
											}
											try {
												message1.setSenderEmailAddress(cal.getOrganizer().getDisplayName());
												mapi.setFrom(new MailAddress(cal.getOrganizer().getDisplayName()));
											} catch (Exception e) {
												try {
													message1.setSenderEmailAddress(message.getSenderEmailAddress());
													mapi.setFrom(new MailAddress(message.getSenderEmailAddress()));
												} catch (Exception e1) {
													message1.setSenderEmailAddress(
															mess.getFrom().toString() + "@gmail.com");
													mapi.setFrom(
															new MailAddress(mess.getFrom().toString() + "@gmail.com"));
												}
											}
											MapiCalendar cal1 = null;
											try {
												cal1 = (MapiCalendar) message1.toMapiMessageItem();
											} catch (Exception e1) {
												continue;
											}

											//
//											File file1 = new File(
//													temppathm + mm.namingconventionmapi(message1) + i + ".msg");
//											file1.createNewFile();
//											message1.save(file1.getAbsolutePath(), SaveOptions.getDefaultMsg());
//											mapi.addAttachment(new Attachment(
//													temppathm + mm.namingconventionmapi(message1) + i + ".msg"));
//											file1.delete();
											//

											cal1.save(temppathm + File.separator + mm.namingconventionmapi(message1)

													+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
											File file = new File(temppathm + File.separator
													+ mm.namingconventionmapi(message1) + "_" + i + ".ics");

											mapi.addAttachment(new Attachment(temppathm + File.separator
													+ mm.namingconventionmapi(message1) + "_" + i + ".ics"));
											file.delete();

											if (mm.chckbxRemoveDuplicacy.isSelected()) {

												String input = mm.duplicacymapiCal(cal);
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listdupliccal.contains(input)) {
													listdupliccal.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforimap_output.appendMessage(iconnforimap_output,
																	path, mapi);
															count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mapi);
														count_destination++;
													}
												} else {
													continue;
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														System.out.println("date_Filter");
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mapi);
														count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													count_destination++;
												}
											}
										} catch (Exception e) {
											if (e.getMessage()
													.equalsIgnoreCase("The operation 'FetchMessage' terminated.")
													|| e.getMessage().equalsIgnoreCase("ConnectFailure")
													|| e.getMessage().equalsIgnoreCase("Operation has been canceled")
													|| e.getMessage().contains("Operation failed.")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage().equalsIgnoreCase(
															"The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Network is unreachable: connect")) {
												mm.Progressbar.setVisible(false);

												i--;

											}
											connectionHandle();
											mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " "
													+ System.lineSeparator());
											continue;
										}

									} else if (message1.getMessageClass().equals("IPM.Task")) {
										try {
											MapiTask task = (MapiTask) message1.toMapiMessageItem();
											MailMessage mapi = new MailMessage();
											File file = new File(
													temppathm + mm.namingconventionmapi(message1) + i + ".msg");
											file.createNewFile();

											message1.save(file.getAbsolutePath(), SaveOptions.getDefaultMsg());
											mess.addAttachment(new Attachment(
													temppathm + mm.namingconventionmapi(message1) + i + ".msg"));
											file.delete();

											if (mess.getBody() != null) {
												mapi.setBody(mess.getBody());
											} else {
												mapi.setBody(message1.getBody());
											}
											try {
												mapi.setFrom(mess.getFrom());
											} catch (Exception e) {
											}
											try {
												mapi.setDate(mess.getDate());
											} catch (Exception e) {
												mapi.setDate(null);
											}
											if (mm.chckbxRemoveDuplicacy.isSelected()) {
												String input = "";
												if (message1.getMessageClass().equals("IPM.Task")) {
													input = mm.duplicacymapiTask(task);
												}
												input = input.replaceAll("\\s", "");
												input = input.trim();
												if (!listduplictask.contains(input)) {
													listduplictask.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag)
															clientforimap_output.appendMessage(iconnforimap_output,
																	path, mess);
														count_destination++;
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mess);
														count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mess);
														count_destination++;
													}
												} else {
													System.out.println(clientforimap_output.getUsername() + "Task ==>"
															+ count_destination);
													clientforimap_output.appendMessage(iconnforimap_output, path, mess);
													count_destination++;
												}
											}
										} catch (Exception e) {
											if (e.getMessage()
													.equalsIgnoreCase("The operation 'FetchMessage' terminated.")
													|| e.getMessage().equalsIgnoreCase("ConnectFailure")
													|| e.getMessage().equalsIgnoreCase("Operation has been canceled")
													|| e.getMessage().contains("Operation failed")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage().equalsIgnoreCase(
															"The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Network is unreachable: connect")) {
												mm.Progressbar.setVisible(false);

												i--;

											}
											connectionHandle();
											mf.logger.warning("Exception : " + e.getMessage() + "Task" + " "
													+ System.lineSeparator());
											continue;
										}
									} else {
										try {
											String messageid = mailimap(mess, path);
											if (!messageid.equalsIgnoreCase("")) {
												if (((message.getFlags()
														& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
													clientforimap_output.changeMessageFlags(iconnforimap_output,
															messageid, ImapMessageFlags.isRead());
												} else {
													clientforimap_output.removeMessageFlags(iconnforimap_output,
															messageid, ImapMessageFlags.isRead());
												}
											}
										} catch (Error ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mm.namingconventionmapi(message1));
										} catch (Exception e) {
											StringWriter sw = new StringWriter();
											e.printStackTrace(new PrintWriter(sw));
											String exceptionAsString = sw.toString();
											int number = message.getAttachments().size();
											System.out.println(number + " ::>> ");
											System.out.println(e.getMessage());
											if (exceptionAsString.contains("Message too large") || number > 10) {
												File f11 = new File(
														(System.getProperty("user.home") + File.separator + "Desktop")
																+ File.separator + fname + File.separator + "Attachment"
																+ File.separator + mm.namingconventionmapi(message));
												f11.mkdirs();
												mf.logger.info(
														"Message size was greater than allowed size so attachment has been deleted and saved in "
																+ f11.getAbsolutePath());
												for (MapiAttachment attachment : message.getAttachments()) {

													attachment.save(f11.getAbsolutePath() + File.separator
															+ main_multiplefile.getRidOfIllegalFileNameCharacters(
																	attachment.getLongFileName()));

												}
												try {
													mess.getAttachments().clear();
													String messageid = mailimap(mess, path);
													if (!messageid.equalsIgnoreCase("")) {
														if (((message.getFlags()
																& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
															clientforimap_output.changeMessageFlags(iconnforimap_output,
																	messageid, ImapMessageFlags.isRead());
														} else {
															clientforimap_output.removeMessageFlags(iconnforimap_output,
																	messageid, ImapMessageFlags.isRead());
														}
													}
												} catch (Exception e1) {

													e1.printStackTrace();
												}

											} else if (e.getMessage().contains("Operation failed")
													|| e.getMessage().contains("Operation has been canceled")
													|| e.getMessage().contains("Object has been disposed")
													|| e.getMessage().contains("Server logging out")
													|| e.getMessage()
															.contains("The operation 'FetchMessage' terminated.")
													|| e.getMessage().contains(
															"No connection could be made because the target machine actively refused it.")
													|| e.getMessage().contains("ConnectFailure")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage().contains("Operation failed.")
													|| e.getMessage().contains("Operation has been canceled")
													|| e.getMessage()
															.contains("The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Network is unreachable: connect")) {
												mm.Progressbar.setVisible(false);

												i--;
											}
											connectionHandle();
											mf.logger.warning("Exception : " + e.getMessage() + "Message " + " " + i
													+ mm.namingconventionmapi(message) + System.lineSeparator());
											continue;
										}

									}
									mm.lbl_progressreport.setText("Total Message Saved Count  " + count_destination
											+ "  " + Folder + "   Extracting messsage " + message.getSubject());

								} catch (Exception e) {
									System.out.println(e.getMessage() + "   Mails  ===> 695");
									if (e.getMessage().equalsIgnoreCase("The operation 'FetchMessage' terminated.")
											|| e.getMessage().equalsIgnoreCase("ConnectFailure")
											|| e.getMessage().equalsIgnoreCase("Operation has been canceled")
											|| e.getMessage().contains("Operation failed")
											|| e.getMessage().contains("Rate limit hit")
											|| e.getMessage()
													.equalsIgnoreCase("The operation 'AppendMessage' terminated.")
											|| e.getMessage().contains("Software caused connection abort: recv failed")
											|| e.getMessage().contains("Network is unreachable: connect")) {
										mm.Progressbar.setVisible(false);
										i--;
									}
									connectionHandle();
									continue;
								}

							}
						}

					}
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderforpstost_gmail(folderInfo, sfolder);
				}
				path = mm.removefoldergmail(path);
			} catch (Exception e) {
				continue;
			}

		}

	}

	@SuppressWarnings("resource")
	private void getsubfolderforpstost_gmail(FolderInfo f, String sfolder) {
		listduplicacy.clear();
		listdupliccal.clear();
		listdupliccontact.clear();
		listduplictask.clear();
		MailConversionOptions options = new MailConversionOptions();
		FolderInfoCollection subfolder = f.getSubFolders();

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
				if (filetype.equalsIgnoreCase("GoDaddy email")) {
					Folder = Folder.replaceAll("[^a-zA-Z0-9]", "");

				}
				sfolder = sfolder + File.separator + Folder;
				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}
					if (pstfolderlist.get(l).equalsIgnoreCase(sfolder)) {
						path = path4 + "\\" + main_multiplefile.fname + "\\" + sfolder;
						mm.lbl_progressreport.setText("Getting : " + Folder);
						String[] p1 = path.split("\\\\");
						x1 = path4;
						for (int i1 = 1; i1 < p1.length; i1++) {

							// Zoho
							if (filetype.equalsIgnoreCase("Zoho Mail")) {
								if (p1[i1].equalsIgnoreCase("Inbox") || p1[i1].equalsIgnoreCase("Drafts")
										|| p1[i1].equalsIgnoreCase("Outbox"))
									p1[i1] = p1[i1] + "_" + "zoho";
							}
							// Yandex
							if (filetype.equalsIgnoreCase("Yandex Mail")) {
								x1 = x1 + "|" + p1[i1].substring(0, 4);
							} else {
								x1 = x1 + "/" + p1[i1];
							}
							path = x1;

							try {
								if (clientforimap_output.existFolder(x1)) {
									clientforimap_output.selectFolder(x1);
								} else {
									clientforimap_output.createFolder(x1);
									clientforimap_output.selectFolder(x1);
								}
							} catch (Exception e) {
								e.printStackTrace();
								if (e.getMessage().equals("Object has been disposed.")) {
									System.out.println(e.getMessage() + "  242");
									connectionHandle();
									i1--;
								} else if (e.getMessage().contains("Software caused connection abort: recv failed")
										|| e.getMessage().contains("Network is unreachable: connect")
										|| e.getMessage().contains("Operation failed.")
										|| e.getMessage().equalsIgnoreCase("The operation 'FetchMessage' terminated.")
										|| e.getMessage().equalsIgnoreCase("ConnectFailure")
										|| e.getMessage().equalsIgnoreCase("Operation has been canceled")
										|| e.getMessage().contains("Operation failed")
										|| e.getMessage().contains("Rate limit hit") || e.getMessage()
												.equalsIgnoreCase("The operation 'AppendMessage' terminated.")) {

									System.out.println(e.getMessage() + "  252 ");
									connectionHandle();
								} else {
									System.out.println(x1);
									System.out.println(e.getMessage() + "  254 ");
									System.out.println("Continue");
									continue;
								}

							}
						}
						listduplicacy.clear();
						listdupliccal.clear();
						listdupliccontact.clear();
						listduplictask.clear();

						MessageInfoCollection messageInfoCollection = null;
						try {
							messageInfoCollection = folderf.getContents();
						} catch (Exception e1) {

							e1.printStackTrace();
						}

						if (!(messageInfoCollection == null)) {

							boolean s22 = false;
							int messagesize;
							if (main_multiplefile.demo) {
								if (messageInfoCollection.size() <= All_Data.demo_count) {
									messagesize = messageInfoCollection.size();
								} else {
									messagesize = All_Data.demo_count;
								}

							} else {
								messagesize = messageInfoCollection.size();
							}

							for (int i = 0; i < messagesize; i++) {
								try {
									if (mm.stop) {
										break;
									}
									if ((i % 100) == 0) {
										System.gc();
									}
									if (count_destination != 0) {
										if ((count_destination % 500) == 0) {
											if (s22) {
												connectionHandle1();
											}
											s22 = true;
										}
									}
//									MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);
//									MapiMessage message1 = pst.extractMessage(messageInfo);
//									MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
//									MailConversionOptions de = new MailConversionOptions();
//									MailMessage mess = message1.toMailMessage(de);

									MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);
									MapiMessage message1 = pst.extractMessage(messageInfo);
									MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
									MailConversionOptions de = new MailConversionOptions();
									MailMessage mess1 = message1.toMailMessage(de);
									MapiMessage message = MapiMessage.fromMailMessage(mess1, d);

									MailMessage mess = message.toMailMessage(options);
									if (mm.chckbxMigrateOrBackup.isSelected()) {
										mess.getAttachments().clear();
										MapiAttachmentCollection m = message1.getAttachments();
										for (int j = 0; j < m.size(); j++) {
											m.clear();
										}
									}
//									MapiMessage message = MapiMessage.fromMailMessage(mess, d);

									if (main_multiplefile.datefilter.isSelected()) {
										datevalidflag = mm.checkdate(message1, mess);
										System.out.println(datevalidflag);
									}
									if (message1.getMessageClass().equals("IPM.Task")
											|| message1.getMessageClass().equals("IPM.StickyNote")
											|| message1.getMessageClass().equals("IPM.Contact")
											|| message1.getMessageClass().equals("IPM.Appointment")
											|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {
										int bct = message1.getBodyType();
										if (bct == 0) {
											message1.setBody(mess1.getBody());
										} else {
											message1.setBody(mess1.getBody());
										}
									}
									if (message1.getMessageClass().equals("IPM.Contact")) {
										MailMessage mapi = new MailMessage();
										try {
											MapiContact con = (MapiContact) message1.toMapiMessageItem();
											try {
												mapi.setSubject(con.getSubject() + "_" + i);
											} catch (Exception e) {

												mapi.setSubject("");
											}
											try {
												mapi.setDate(mess.getDate());
											} catch (Exception e) {
												mapi.setDate(null);
											}

											try {
												message1.setSenderEmailAddress(mess.getFrom().toString());
												mapi.setFrom(mess.getFrom());
											} catch (Exception e) {
											}
											for (Attachment attachment : mess.getAttachments()) {
												attachment.save(temppathm + File.separator + attachment.getName());
												file11 = new File(temppathm + File.separator + attachment.getName());
												mapi.addAttachment(new Attachment(
														temppathm + File.separator + attachment.getName()));
												file11.delete();
											}

											con.save(temppathm + File.separator + mm.namingconventionmapi(message1)
													+ "_" + i + ".vcf", ContactSaveFormat.VCard);
											File file = new File(temppathm + File.separator
													+ mm.namingconventionmapi(message1) + "_" + i + ".vcf");
											mapi.addAttachment(new Attachment(temppathm + File.separator
													+ mm.namingconventionmapi(message1) + "_" + i + ".vcf"));
											file.delete();

											if (mm.chckbxRemoveDuplicacy.isSelected()) {

												String input = mm.duplicacymapiContact(con);
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listdupliccontact.contains(input)) {
													System.out.println("Not a duplicate message");
													listdupliccontact.add(input);

													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforimap_output.appendMessage(iconnforimap_output,
																	path, mapi);
															count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mapi);
														count_destination++;
													}
												} else {
													continue;
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mapi);
														count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													count_destination++;
												}
											}
										} catch (Exception e) {
											if (e.getMessage()
													.equalsIgnoreCase("The operation 'FetchMessage' terminated.")
													|| e.getMessage().equalsIgnoreCase("ConnectFailure")
													|| e.getMessage().equalsIgnoreCase("Operation has been canceled")
													|| e.getMessage().contains("Operation failed")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage().equalsIgnoreCase(
															"The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Network is unreachable: connect")) {
												mm.Progressbar.setVisible(false);
												i--;

											}
											connectionHandle();
											mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
													+ System.lineSeparator());
											continue;
										}

									} else if (message1.getMessageClass().equals("IPM.Appointment")
											|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {
										MailMessage mapi = new MailMessage();
										try {
											MapiCalendar cal = null;
											try {
												cal = (MapiCalendar) message1.toMapiMessageItem();
											} catch (Exception e1) {
												continue;
											}
											try {
												mapi.setSubject(cal.getSubject() + "_" + i);
											} catch (Exception e) {
												mapi.setSubject("");
											}
											try {
												mapi.setDate(mess.getDate());
											} catch (Exception e) {
												mapi.setDate(null);
											}
											try {
												message1.setSenderEmailAddress(cal.getOrganizer().getDisplayName());
												mapi.setFrom(new MailAddress(cal.getOrganizer().getDisplayName()));
											} catch (Exception e) {
												try {
													message1.setSenderEmailAddress(message.getSenderEmailAddress());
													mapi.setFrom(new MailAddress(message.getSenderEmailAddress()));
												} catch (Exception e1) {
													message1.setSenderEmailAddress(
															mess.getFrom().toString() + "@gmail.com");
													mapi.setFrom(
															new MailAddress(mess.getFrom().toString() + "@gmail.com"));
												}
											}
											MapiCalendar cal1 = null;
											try {
												cal1 = (MapiCalendar) message1.toMapiMessageItem();
											} catch (Exception e1) {
												continue;
											}

											//
//											File file1 = new File(
//													temppathm + mm.namingconventionmapi(message1) + i + ".msg");
//											file1.createNewFile();
//											message1.save(file1.getAbsolutePath(), SaveOptions.getDefaultMsg());
//											mapi.addAttachment(new Attachment(
//													temppathm + mm.namingconventionmapi(message1) + i + ".msg"));
//											file1.delete();
											//

											cal1.save(temppathm + File.separator + mm.namingconventionmapi(message1)

													+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
											File file = new File(temppathm + File.separator
													+ mm.namingconventionmapi(message1) + "_" + i + ".ics");

											mapi.addAttachment(new Attachment(temppathm + File.separator
													+ mm.namingconventionmapi(message1) + "_" + i + ".ics"));
											file.delete();

											if (mm.chckbxRemoveDuplicacy.isSelected()) {

												String input = mm.duplicacymapiCal(cal);
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listdupliccal.contains(input)) {
													listdupliccal.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforimap_output.appendMessage(iconnforimap_output,
																	path, mapi);
															count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mapi);
														count_destination++;
													}
												} else {
													continue;
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														System.out.println("date_Filter");
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mapi);
														count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													count_destination++;
												}
											}
										} catch (Exception e) {
											if (e.getMessage()
													.equalsIgnoreCase("The operation 'FetchMessage' terminated.")
													|| e.getMessage().equalsIgnoreCase("ConnectFailure")
													|| e.getMessage().equalsIgnoreCase("Operation has been canceled")
													|| e.getMessage().contains("Operation failed.")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage().equalsIgnoreCase(
															"The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Network is unreachable: connect")) {
												mm.Progressbar.setVisible(false);

												i--;

											}
											connectionHandle();
											mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " "
													+ System.lineSeparator());
											continue;
										}

									} else if (message1.getMessageClass().equals("IPM.Task")) {
										try {
											MapiTask task = (MapiTask) message1.toMapiMessageItem();
											MailMessage mapi = new MailMessage();
											File file = new File(
													temppathm + mm.namingconventionmapi(message1) + i + ".msg");
											file.createNewFile();

											message1.save(file.getAbsolutePath(), SaveOptions.getDefaultMsg());
											mess.addAttachment(new Attachment(
													temppathm + mm.namingconventionmapi(message1) + i + ".msg"));
											file.delete();

											if (mess.getBody() != null) {
												mapi.setBody(mess.getBody());
											} else {
												mapi.setBody(message1.getBody());
											}
											try {
												mapi.setFrom(mess.getFrom());
											} catch (Exception e) {
											}
											try {
												mapi.setDate(mess.getDate());
											} catch (Exception e) {
												mapi.setDate(null);
											}
											if (mm.chckbxRemoveDuplicacy.isSelected()) {
												String input = "";
												if (message1.getMessageClass().equals("IPM.Task")) {
													input = mm.duplicacymapiTask(task);
												}
												input = input.replaceAll("\\s", "");
												input = input.trim();
												if (!listduplictask.contains(input)) {
													listduplictask.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag)
															clientforimap_output.appendMessage(iconnforimap_output,
																	path, mess);
														count_destination++;
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mess);
														count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mess);
														count_destination++;
													}
												} else {
													System.out.println(clientforimap_output.getUsername() + "Task ==>"
															+ count_destination);
													clientforimap_output.appendMessage(iconnforimap_output, path, mess);
													count_destination++;
												}
											}
										} catch (Exception e) {
											if (e.getMessage()
													.equalsIgnoreCase("The operation 'FetchMessage' terminated.")
													|| e.getMessage().equalsIgnoreCase("ConnectFailure")
													|| e.getMessage().equalsIgnoreCase("Operation has been canceled")
													|| e.getMessage().contains("Operation failed")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage().equalsIgnoreCase(
															"The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Network is unreachable: connect")) {
												mm.Progressbar.setVisible(false);

												i--;

											}
											connectionHandle();
											mf.logger.warning("Exception : " + e.getMessage() + "Task" + " "
													+ System.lineSeparator());
											continue;
										}
									} else {
										try {
											String messageid = mailimap(mess, path);
											if (!messageid.equalsIgnoreCase("")) {
												if (((message.getFlags()
														& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
													clientforimap_output.changeMessageFlags(iconnforimap_output,
															messageid, ImapMessageFlags.isRead());
												} else {
													clientforimap_output.removeMessageFlags(iconnforimap_output,
															messageid, ImapMessageFlags.isRead());
												}
											}
										} catch (Error ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mm.namingconventionmapi(message1));
										} catch (Exception e) {
											StringWriter sw = new StringWriter();
											e.printStackTrace(new PrintWriter(sw));
											String exceptionAsString = sw.toString();
											int number = message.getAttachments().size();
											System.out.println(number + " ::>> ");
											System.out.println(e.getMessage());
											if (exceptionAsString.contains("Message too large") || number > 10) {
												File f11 = new File(
														(System.getProperty("user.home") + File.separator + "Desktop")
																+ File.separator + fname + File.separator + "Attachment"
																+ File.separator + mm.namingconventionmapi(message));
												f11.mkdirs();
												mf.logger.info(
														"Message size was greater than allowed size so attachment has been deleted and saved in "
																+ f11.getAbsolutePath());
												for (MapiAttachment attachment : message.getAttachments()) {

													attachment.save(f11.getAbsolutePath() + File.separator
															+ main_multiplefile.getRidOfIllegalFileNameCharacters(
																	attachment.getLongFileName()));

												}
												try {
													mess.getAttachments().clear();
													String messageid = mailimap(mess, path);
													if (!messageid.equalsIgnoreCase("")) {
														if (((message.getFlags()
																& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
															clientforimap_output.changeMessageFlags(iconnforimap_output,
																	messageid, ImapMessageFlags.isRead());
														} else {
															clientforimap_output.removeMessageFlags(iconnforimap_output,
																	messageid, ImapMessageFlags.isRead());
														}
													}
												} catch (Exception e1) {

													e1.printStackTrace();
												}

											} else if (e.getMessage().contains("Operation failed")
													|| e.getMessage().contains("Operation has been canceled")
													|| e.getMessage().contains("Object has been disposed")
													|| e.getMessage().contains("Server logging out")
													|| e.getMessage()
															.contains("The operation 'FetchMessage' terminated.")
													|| e.getMessage().contains(
															"No connection could be made because the target machine actively refused it.")
													|| e.getMessage().contains("ConnectFailure")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage().contains("Operation failed.")
													|| e.getMessage().contains("Operation has been canceled")
													|| e.getMessage()
															.contains("The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Network is unreachable: connect")) {
												mm.Progressbar.setVisible(false);

												i--;
											}
											connectionHandle();
											mf.logger.warning("Exception : " + e.getMessage() + "Message " + " " + i
													+ mm.namingconventionmapi(message) + System.lineSeparator());
											continue;
										}

									}
									mm.lbl_progressreport.setText("Total Message Saved Count  " + count_destination
											+ "  " + Folder + "   Extracting messsage " + message.getSubject());

								} catch (Exception e) {
									System.out.println(e.getMessage() + "   Mails  ===> 695");
									if (e.getMessage().equalsIgnoreCase("The operation 'FetchMessage' terminated.")
											|| e.getMessage().equalsIgnoreCase("ConnectFailure")
											|| e.getMessage().equalsIgnoreCase("Operation has been canceled")
											|| e.getMessage().contains("Operation failed")
											|| e.getMessage().contains("Rate limit hit")
											|| e.getMessage()
													.equalsIgnoreCase("The operation 'AppendMessage' terminated.")
											|| e.getMessage().contains("Software caused connection abort: recv failed")
											|| e.getMessage().contains("Network is unreachable: connect")) {
										mm.Progressbar.setVisible(false);
										i--;
									}
									connectionHandle();
									continue;
								}

							}
						}

					}
				}
				if (folderf.hasSubFolders()) {

					getsubfolderforpstost_gmail(folderf, sfolder);
				}

				path = mm.removefoldergmail(path);
				sfolder = mm.removefolder(sfolder);

			} catch (Exception e) {
				continue;
			}
		}

	}

	String mailimap(MailMessage message, String path) throws Exception {
		String Messageid = "";
		if (mm.chckbxRemoveDuplicacy.isSelected()) {

			String input = mm.duplicacymail(message);

			if (!listduplicacy.contains(input)) {
				listduplicacy.add(input);

				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
						foldermessagecount++;
						count_destination++;
					}
				} else {
					Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
					foldermessagecount++;
					count_destination++;
				}
			}
		} else {
			if (main_multiplefile.datefilter.isSelected()) {
				if (datevalidflag) {
					Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
					foldermessagecount++;
					count_destination++;
				}
			} else {

				Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
				count_destination++;

			}
		}
		return Messageid;
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
					} else
					{
					clientforimap_output.dispose();
					clientforimap_output = connectiontogmail_output();
					}

				} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {
					clientforimap_output.dispose();
					clientforimap_output = connectiontoyahoo_output(clientforimap_output);

				} else if (filetype.equalsIgnoreCase("AOL")) {
					clientforimap_output.dispose();
					clientforimap_output = connectiontoaol_output(clientforimap_output);
				} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {
					clientforimap_output.dispose();
					clientforimap_output = connectiontoinaws_output(clientforimap_output);
				} else if (filetype.equalsIgnoreCase("Icloud")) {
					clientforimap_output.dispose();
					clientforimap_output = connectiontoicloud_output();
				} else if (filetype.equalsIgnoreCase("Zoho Mail")) {
					clientforimap_output.dispose();
					clientforimap_output = connectiontozoho_output();
				} else if (filetype.equalsIgnoreCase("Yandex Mail")) {
					clientforimap_output.dispose();
					clientforimap_output = connectiontoYandex_output();
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

	// Amazon
	@SuppressWarnings("deprecation")
	public ImapClient connectiontoinaws_output(ImapClient clientforimap_output) throws Exception {
		clientforimap_output = new ImapClient(domain_p3, portnofiletype, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
//		clientforimap_output.setTimeout(5 * 60 * 1000);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

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
