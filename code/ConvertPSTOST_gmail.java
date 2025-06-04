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
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiMessageFlags;
import com.aspose.email.MapiTask;
import com.aspose.email.MessageInfo;
import com.aspose.email.MessageInfoCollection;
import com.aspose.email.PersonalStorage;
import com.aspose.email.SaveOptions;
import com.aspose.email.SecurityOptions;

public class ConvertPSTOST_gmail implements Runnable {
	private List<String> listdupliccal = new ArrayList<String>();
	private List<String> listduplictask = new ArrayList<String>();
	private List<String> listdupliccontact = new ArrayList<String>();
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
	static String path = "";
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
	String domain_p;
	int portnofiletype;

	public ConvertPSTOST_gmail(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm, String username_p3, String password_p3, String Folderuri,
			ImapClient clientforimap_output, IConnection iconnforimap_output, String path, String fname, String match,
			String domain_p, int portnofiletype) {

		this.mf = mf;
		this.mm = mm;
		this.filetype = filetype;
		this.filepath = filepath;
		this.destination_path = destination_path;
		ConvertPSTOST_gmail.count_destination = count_destination;
		this.pstfolderlist = pstfolderlist;
		this.fromList = fromList;
		this.toList = toList;
		this.temppathm = temppathm;
		this.username_p3 = username_p3;
		this.password_p3 = password_p3;
		this.Folderuri = Folderuri;
		this.clientforimap_output = clientforimap_output;
		this.iconnforimap_output = iconnforimap_output;
		ConvertPSTOST_gmail.path = path;
		this.match = match;
		this.domain_p = domain_p;
		this.portnofiletype = portnofiletype;
	}

	@Override
	public void run() {

		convertPSTOST_gmail(mf, filetype, destination_path, count_destination, filepath, mm, pstfolderlist, fromList,
				toList, temppathm, Folderuri, clientforimap_output, iconnforimap_output, path, match, domain_p,
				portnofiletype);
		main_multiplefile.count_destination = ConvertPSTOST_gmail.count_destination;
		main_multiplefile.match = match;
	}

	private void convertPSTOST_gmail(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm, String folderuri, ImapClient clientforimap_output,
			IConnection iconnforimap_output, String path, String match, String domain_p, int portnofiletype) {
		ConvertPSTOST_gmail.count_destination = 0;
		match = path;
		System.out.println(match + " match ");
		if (filetype.equalsIgnoreCase("GoDaddy email")) {
			main_multiplefile.fname = main_multiplefile.fname.replaceAll("[^a-zA-Z0-9]", "");

		}
		path = path + "/" + main_multiplefile.fname;
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
		String path_new = path + "/" + Folder;
		clientforimap_output.createFolder(iconnforimap_output, path_new);
		clientforimap_output.selectFolder(iconnforimap_output, path_new);

		listdupliccal.clear();
		listduplicacy.clear();
		listdupliccontact.clear();
		listduplictask.clear();
		MessageInfoCollection messageInfoCollection1 = folderInfo2.getContents();
		int countr = 0;
		int messagesize1;
		boolean s2 = false;
		if (main_multiplefile.demo) {
			if (messageInfoCollection1.size() <= All_Data.demo_count) {
				messagesize1 = messageInfoCollection1.size();
			} else {
				messagesize1 = All_Data.demo_count;
			}

		} else {
			messagesize1 = messageInfoCollection1.size();
		}
		System.out.println("message size : " + messagesize1);
		for (int i = 0; i < messagesize1; i++) {

			try {

				if (mm.stop) {
					break;
				}
				if ((i % 100) == 0) {
					System.gc();
				}
//				if ((count_destination % 500) == 0) {
//					if (s2) {
//						connectionHandle1();
//					}
//					s2 = true;
//				}

				MessageInfo messageInfo = (MessageInfo) messageInfoCollection1.get_Item(i);

				MapiMessage message1 = pst.extractMessage(messageInfo);
				MapiConversionOptions d = MapiConversionOptions.getUnicodeFormat();
				MailConversionOptions de = new MailConversionOptions();
				MailMessage mess = message1.toMailMessage(de);
				MapiMessage message = MapiMessage.fromMailMessage(mess, d);
				if (mm.chckbxMigrateOrBackup.isSelected()) {
					mess.getAttachments().clear();
				}
				if (main_multiplefile.datefilter.isSelected()) {
					datevalidflag = mm.checkdate(message, mess);
				}
				if (message1.getMessageClass().equals("IPM.Contact")) {
					MailMessage mapi = new MailMessage();
					try {
						MapiContact con = (MapiContact) message1.toMapiMessageItem();
						try {
							mapi.setSubject(con.getSubject());
						} catch (Exception e) {

							mapi.setSubject("");
						}
						try {
							mapi.setBody(con.getBody());
						} catch (Exception e) {

							mapi.setBody("");
						}

						try {
							message1.setSenderEmailAddress(mess.getFrom().toString());
							mapi.setFrom(mess.getFrom());
						} catch (Exception e) {
						}

						con.save(temppathm + File.separator + countr + "contacttemp.vcf", ContactSaveFormat.VCard);
						File file = new File(temppathm + File.separator + countr + "contacttemp.vcf");
						mapi.addAttachment(new Attachment(temppathm + File.separator + countr + "contacttemp.vcf"));
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
										clientforimap_output.appendMessage(path_new, mapi);
										ConvertPSTOST_gmail.count_destination++;
									}
								} else {
									clientforimap_output.appendMessage(path_new, mapi);
									ConvertPSTOST_gmail.count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforimap_output.appendMessage(path_new, mapi);
									ConvertPSTOST_gmail.count_destination++;
								}
							} else {
								clientforimap_output.appendMessage(path_new, mapi);
								ConvertPSTOST_gmail.count_destination++;
							}
						}
						countr++;
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mm.namingconventionmapi(message));
					} catch (Exception e) {
						System.out.println(e.getMessage() + "  ===> 270 Contact ");
						if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage().contains("Operation has been canceled")
								|| e.getMessage().contains("Operation failed.")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
								|| e.getMessage().contains("Software caused connection abort: recv failed")
								|| e.getMessage().contains("Network is unreachable: connect")
								|| e.getMessage().contains("Object has been disposed.")
								|| e.getMessage().contains("Connection or outbound has closed")
								|| e.getMessage().contains("Connection reset by peer: socket write error")) {
							mm.Progressbar.setVisible(false);
							i--;
						}
						connectionHandle();
						mf.logger.warning(e.getMessage() + "Contact" + " " + countr + System.lineSeparator());
						continue;
					}

				} else if (message1.getMessageClass().equals("IPM.Appointment")
						|| message1.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message1.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
					MailMessage mapi = new MailMessage();
					try {

						MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();

						try {
							mapi.setSubject(cal.getSubject());
						} catch (Exception e) {

							mapi.setSubject("");
						}
						try {
							mapi.setBody(cal.getBody());
						} catch (Exception e) {

							mapi.setBody("");
						}

						try {
							message1.setSenderEmailAddress(cal.getOrganizer().getDisplayName());
							mapi.setFrom(new MailAddress(cal.getOrganizer().getDisplayName()));
						} catch (Exception e) {
							try {
								message1.setSenderEmailAddress(message.getSenderEmailAddress());
								mapi.setFrom(new MailAddress(message.getSenderEmailAddress()));
							} catch (Exception e1) {
								message1.setSenderEmailAddress(mess.getFrom().toString() + "@gmail.com");
								mapi.setFrom(new MailAddress(mess.getFrom().toString() + "@gmail.com"));
							}
						}

						cal.save(temppathm + File.separator + countr + "caltemp.ics", AppointmentSaveFormat.Ics);
						File file = new File(temppathm + File.separator + countr + "caltemp.ics");

						mapi.addAttachment(new Attachment(temppathm + File.separator + countr + "caltemp.ics"));
						file.delete();

						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = mm.duplicacymapiCal(cal);
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listdupliccal.contains(input)) {
								listdupliccal.add(input);

								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										clientforimap_output.appendMessage(iconnforimap_output, path_new, mapi);
										ConvertPSTOST_gmail.count_destination++;
									}
								} else {
									clientforimap_output.appendMessage(iconnforimap_output, path_new, mapi);
									ConvertPSTOST_gmail.count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforimap_output.appendMessage(iconnforimap_output, path_new, mapi);
									ConvertPSTOST_gmail.count_destination++;
								}
							} else {
								clientforimap_output.appendMessage(iconnforimap_output, path_new, mapi);
								ConvertPSTOST_gmail.count_destination++;
							}
						}
						countr++;
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mm.namingconventionmapi(message));
					} catch (Exception e) {
						System.out.println(e.getMessage() + "  ===> 361 Calendar ");
						if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage().contains("Operation has been canceled")
								|| e.getMessage().contains("Operation failed.")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
								|| e.getMessage().contains("Software caused connection abort: recv failed")
								|| e.getMessage().contains("Network is unreachable: connect")
								|| e.getMessage().contains("Object has been disposed.")
								|| e.getMessage().contains("Connection or outbound has closed")
								|| e.getMessage().contains("Connection reset by peer: socket write error")) {
							mm.Progressbar.setVisible(false);

							i--;
						}
						connectionHandle();

						mf.logger.warning(e.getMessage() + "Calendar" + " " + countr + System.lineSeparator());
						continue;
					}

				} else if (message1.getMessageClass().equals("IPM.Task")) {

					String s = filetype;
					try {

						filetype = "MSG";
						MapiTask task = null;
						if (message1.getMessageClass().equals("IPM.Task")) {
							task = (MapiTask) message1.toMapiMessageItem();
						}

						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = "";
							if (messageInfo.getMessageClass().equals("IPM.Task")) {
								input = mm.duplicacymapiTask(task);
							}
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listduplictask.contains(input)) {
								System.out.println("Not a duplicate message");
								listduplictask.add(input);

								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										clientforimap_output.appendMessage(iconnforimap_output, path_new, mess);
										ConvertPSTOST_gmail.count_destination++;
									}
								} else {
									clientforimap_output.appendMessage(iconnforimap_output, path_new, mess);
									ConvertPSTOST_gmail.count_destination++;
								}

							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforimap_output.appendMessage(iconnforimap_output, path_new, mess);
									ConvertPSTOST_gmail.count_destination++;
								}
							} else {
								clientforimap_output.appendMessage(iconnforimap_output, path_new, mess);
								ConvertPSTOST_gmail.count_destination++;
							}
						}
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mm.namingconventionmapi(message));
					} catch (Exception e) {
						System.out.println(e.getMessage() + "  ===> 430 Task ");
						if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage().contains("Operation has been canceled")
								|| e.getMessage().contains("Operation failed.")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
								|| e.getMessage().contains("Software caused connection abort: recv failed")
								|| e.getMessage().contains("Network is unreachable: connect")
								|| e.getMessage().contains("Object has been disposed.")
								|| e.getMessage().contains("Connection or outbound has closed")
								|| e.getMessage().contains("Connection reset by peer: socket write error")) {
							mm.Progressbar.setVisible(false);

							i--;
						}
						connectionHandle();
						mf.logger.warning(e.getMessage() + "Task" + " " + countr + System.lineSeparator());
						continue;
					} finally {
						filetype = s;
					}

				} else {
					try {
						String messageid = mailimap(mess, path_new);

						if (!messageid.equals("")) {

							if (((message.getFlags()
									& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
								clientforimap_output.changeMessageFlags(iconnforimap_output, messageid,
										ImapMessageFlags.isRead());
							} else {
								clientforimap_output.removeMessageFlags(iconnforimap_output, messageid,
										ImapMessageFlags.isRead());
							}
						}
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mm.namingconventionmapi(message));
					} catch (Exception e) {
						StringWriter sw = new StringWriter();
						e.printStackTrace(new PrintWriter(sw));
						String exceptionAsString = sw.toString();

						System.out.println(e.getMessage() + "  ===> 472 Mails ");
						if (exceptionAsString.contains("Message too large")) {
							File f = new File((System.getProperty("user.home") + File.separator + "Desktop")
									+ File.separator + calendertime + File.separator + "Attachment" + File.separator
									+ mm.namingconventionmapi(message));
							f.mkdirs();
							mf.logger.info(
									"Message size was greater than allowed size so attachment has been deleted and saved in "
											+ f.getAbsolutePath());
							for (MapiAttachment attachment : message.getAttachments()) {
								attachment.save(f.getAbsolutePath() + File.separator + main_multiplefile
										.getRidOfIllegalFileNameCharacters(attachment.getLongFileName()));
							}
							try {
								mess.getAttachments().clear();

								String messageid = mailimap(mess, path);
								if (!messageid.equalsIgnoreCase("")) {
									if (((message.getFlags()
											& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
										clientforimap_output.changeMessageFlags(iconnforimap_output, messageid,
												ImapMessageFlags.isRead());

									} else {

										clientforimap_output.removeMessageFlags(iconnforimap_output, messageid,
												ImapMessageFlags.isRead());

									}
								}
							} catch (Exception e1) {

							}

						} else if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage().contains("Operation has been canceled")
								|| e.getMessage().contains("Operation failed.")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
								|| e.getMessage().contains("Software caused connection abort: recv failed")
								|| e.getMessage().contains("Network is unreachable: connect")
								|| e.getMessage().contains("Object has been disposed.")
								|| e.getMessage().contains("Connection or outbound has closed")
								|| e.getMessage().contains("Connection reset by peer: socket write error")) {
							mm.Progressbar.setVisible(false);

							i--;
						}
						connectionHandle();

						e.printStackTrace();
						StringWriter sw1 = new StringWriter();
						PrintWriter pw = new PrintWriter(sw1);
						e.printStackTrace(pw);
						mf.logger.warning(sw1.toString());
						continue;
					}

				}
				mm.lbl_progressreport.setText("Total message Saved Count " + ConvertPSTOST_gmail.count_destination
						+ "  " + Folder + " Extracting messsage " + message.getSubject());

			} catch (Exception e) {
				System.out.println(e.getMessage() + "  ===> 534");
				if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
						|| e.getMessage().contains("ConnectFailure")
						|| e.getMessage().contains("Operation has been canceled")
						|| e.getMessage().contains("Operation failed.") || e.getMessage().contains("Rate limit hit")
						|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
						|| e.getMessage().contains("Software caused connection abort: recv failed")
						|| e.getMessage().contains("Network is unreachable: connect")
						|| e.getMessage().contains("Object has been disposed.")
						|| e.getMessage().contains("Connection or outbound has closed")
						|| e.getMessage().contains("Connection reset by peer: socket write error")) {
					mm.Progressbar.setVisible(false);

					i--;
				}
				connectionHandle();

				e.printStackTrace();
				StringWriter sw = new StringWriter();
				PrintWriter pw = new PrintWriter(sw);
				e.printStackTrace(pw);
				mf.logger.warning(sw.toString());

				continue;

			}

		}

		FolderInfoCollection folderInf = pst.getRootFolder().getSubFolders();

		for (int j = 0; j < folderInf.size(); j++) {
			try {

				if (mm.stop) {
					break;
				}
				boolean s22 = false;
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

						path = path_new + "/" + fol;

						try {
							if (clientforimap_output.existFolder(path)) {
								clientforimap_output.selectFolder(path);
							} else {
								clientforimap_output.createFolder(iconnforimap_output, path);
								clientforimap_output.selectFolder(iconnforimap_output, path);
							}
						} catch (Exception e3) {
							e3.printStackTrace();
						}

						listdupliccal.clear();
						listduplicacy.clear();
						listdupliccontact.clear();
						listduplictask.clear();
						MessageInfoCollection messageInfoCollection = folderInfo.getContents();
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
//								if ((ConvertPSTOST_gmail.count_destination % 1000) == 0) {
//									if (s22) {
//										connectionHandle1();
//									}
//									s22 = true;
//								}

								MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);
								MapiMessage message1 = pst.extractMessage(messageInfo);
								MapiConversionOptions d = MapiConversionOptions.getUnicodeFormat();
								MailConversionOptions de = new MailConversionOptions();
								MailMessage mess1 = message1.toMailMessage(de);
								MapiMessage message = MapiMessage.fromMailMessage(mess1, d);
								MailMessage mess = message.toMailMessage(de);
								if (mm.chckbxMigrateOrBackup.isSelected()) {
									mess.getAttachments().clear();
								}
								if (main_multiplefile.datefilter.isSelected()) {
									datevalidflag = mm.checkdate(message, mess);
								}
								if (message1.getMessageClass().equals("IPM.Contact")) {
									MailMessage mapi = new MailMessage();
									try {
										MapiContact con = (MapiContact) message1.toMapiMessageItem();
										try {
											mapi.setSubject(con.getSubject());
										} catch (Exception e) {

											mapi.setSubject("");
										}
										try {
											mapi.setBody(con.getBody());
										} catch (Exception e) {

											mapi.setBody("");
										}
										try {
											message1.setSenderEmailAddress(mess.getFrom().toString());
											mapi.setFrom(mess.getFrom());
										} catch (Exception e) {
										}
										con.save(temppathm + File.separator + i + mm.namingconventionmapi(message)
												+ ".vcf", ContactSaveFormat.VCard);
										File file = new File(temppathm + File.separator + i
												+ mm.namingconventionmapi(message) + ".vcf");
										mapi.addAttachment(new Attachment(temppathm + File.separator + i
												+ mm.namingconventionmapi(message) + ".vcf"));
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
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mapi);
														ConvertPSTOST_gmail.count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													ConvertPSTOST_gmail.count_destination++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													ConvertPSTOST_gmail.count_destination++;
												}
											} else {
												clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
												ConvertPSTOST_gmail.count_destination++;
											}
										}

									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										System.out.println(e.getMessage() + "  ===> 714");
										if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("Operation failed.")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage()
														.equalsIgnoreCase("The operation 'AppendMessage' terminated.")
												|| e.getMessage()
														.contains("Software caused connection abort: recv failed")
												|| e.getMessage().contains("Network is unreachable: connect")
												|| e.getMessage().contains("Object has been disposed.")
												|| e.getMessage().contains("Connection or outbound has closed")
												|| e.getMessage()
														.contains("Connection reset by peer: socket write error")) {
											mm.Progressbar.setVisible(false);
											i--;
										}
										connectionHandle();
										e.printStackTrace();
										StringWriter sw1 = new StringWriter();
										PrintWriter pw = new PrintWriter(sw1);
										e.printStackTrace(pw);
										mf.logger.warning(sw1.toString());
										continue;
									}
								} else if (message1.getMessageClass().equals("IPM.Appointment")
										|| message1.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message1.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
									MailMessage mapi = new MailMessage();
									try {

										MapiCalendar cal = null;
										try {
											cal = (MapiCalendar) message1.toMapiMessageItem();
										} catch (Exception e2) {

										}

										try {
											mapi.setSubject(cal.getSubject());
										} catch (Exception e) {

											mapi.setSubject("");
										}
										try {
											mapi.setBody(cal.getBody());
										} catch (Exception e) {

											mapi.setBody("");
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
												mapi.setFrom(new MailAddress(mess.getFrom().toString() + "@gmail.com"));
											}
										}
										cal.save(temppathm + File.separator + i + mm.namingconventionmapi(message)
												+ ".ics", AppointmentSaveFormat.Ics);
										File file = new File(temppathm + File.separator + i
												+ mm.namingconventionmapi(message) + ".ics");

										mapi.addAttachment(new Attachment(temppathm + File.separator + i
												+ mm.namingconventionmapi(message) + ".ics"));
										file.delete();
										if (mm.chckbxRemoveDuplicacy.isSelected()) {
											String input = mm.duplicacymapiCal(cal);
											input = input.replaceAll("\\s", "");
											input = input.trim();

											if (!listdupliccal.contains(input)) {
												listdupliccal.add(input);

												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mapi);
														ConvertPSTOST_gmail.count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													ConvertPSTOST_gmail.count_destination++;
												}

											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													ConvertPSTOST_gmail.count_destination++;
												}
											} else {
												clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
												ConvertPSTOST_gmail.count_destination++;
											}
										}

									} catch (OutOfMemoryError ep) {

										ep.printStackTrace();
										StringWriter sw1 = new StringWriter();
										PrintWriter pw = new PrintWriter(sw1);
										ep.printStackTrace(pw);
										mf.logger.warning(sw1.toString());
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										System.out.println(e.getMessage() + "  ===> 821");
										if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("Operation failed.")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage()
														.equalsIgnoreCase("The operation 'AppendMessage' terminated.")
												|| e.getMessage()
														.contains("Software caused connection abort: recv failed")
												|| e.getMessage().contains("Network is unreachable: connect")
												|| e.getMessage().contains("Object has been disposed.")
												|| e.getMessage().contains("Connection or outbound has closed")
												|| e.getMessage()
														.contains("Connection reset by peer: socket write error")) {
											mm.Progressbar.setVisible(false);

											i--;

										}
										connectionHandle();
										e.printStackTrace();
										StringWriter sw1 = new StringWriter();
										PrintWriter pw = new PrintWriter(sw1);
										e.printStackTrace(pw);
										mf.logger.warning(sw1.toString());

										mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + i
												+ mm.namingconventionmapi(message) + System.lineSeparator());
										continue;
									}

								} else if (message1.getMessageClass().equals("IPM.Task")) {
									String s = filetype;
									try {

										filetype = "MSG";
										MapiTask task = null;
										if (message1.getMessageClass().equals("IPM.Task")) {
											task = (MapiTask) message1.toMapiMessageItem();
										}

										if (mm.chckbxRemoveDuplicacy.isSelected()) {
											String input = "";
											if (message1.getMessageClass().equals("IPM.Task")) {
												input = mm.duplicacymapiTask(task);
											}
											input = input.replaceAll("\\s", "");
											input = input.trim();

											if (!listduplictask.contains(input)) {
												System.out.println("Not a duplicate message");
												listduplictask.add(input);
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, path,
																mess);
														ConvertPSTOST_gmail.count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path, mess);
													ConvertPSTOST_gmail.count_destination++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													clientforimap_output.appendMessage(iconnforimap_output, path, mess);
													ConvertPSTOST_gmail.count_destination++;
												}
											} else {
												clientforimap_output.appendMessage(iconnforimap_output, path, mess);
												ConvertPSTOST_gmail.count_destination++;
											}
										}

									} catch (OutOfMemoryError ep) {

										ep.printStackTrace();
										StringWriter sw1 = new StringWriter();
										PrintWriter pw = new PrintWriter(sw1);
										ep.printStackTrace(pw);
										mf.logger.warning(sw1.toString());
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										System.out.println(e.getMessage() + "  ===> 892");
										if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("Operation failed.")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage()
														.equalsIgnoreCase("The operation 'AppendMessage' terminated.")
												|| e.getMessage()
														.contains("Software caused connection abort: recv failed")
												|| e.getMessage().contains("Network is unreachable: connect")
												|| e.getMessage().contains("Object has been disposed.")
												|| e.getMessage().contains("Connection or outbound has closed")
												|| e.getMessage()
														.contains("Connection reset by peer: socket write error")) {
											mm.Progressbar.setVisible(false);

											i--;

										}
										connectionHandle();

										e.printStackTrace();
										StringWriter sw1 = new StringWriter();
										PrintWriter pw = new PrintWriter(sw1);
										e.printStackTrace(pw);
										mf.logger.warning(sw1.toString());

										mf.logger.warning("Exception : " + e.getMessage() + "Task" + " " + i
												+ mm.namingconventionmapi(message) + System.lineSeparator());
										continue;
									} finally {
										filetype = s;
									}

								} else {
									try {
										String messageid = mailimap(mess, path);
										if (messageid.equals("")) {
											if (((message.getFlags()
													& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
												clientforimap_output.changeMessageFlags(iconnforimap_output, messageid,
														ImapMessageFlags.isRead());
											} else {
												clientforimap_output.removeMessageFlags(iconnforimap_output, messageid,
														ImapMessageFlags.isRead());
											}
										}
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										StringWriter sw = new StringWriter();
										e.printStackTrace(new PrintWriter(sw));
										String exceptionAsString = sw.toString();

										System.out.println(e.getMessage() + "  ===> 940");
										if (exceptionAsString.contains("Message too large")) {
											File f = new File(
													(System.getProperty("user.home") + File.separator + "Desktop")
															+ File.separator + calendertime + File.separator
															+ "Attachment" + File.separator
															+ mm.namingconventionmapi(message));
											f.mkdirs();
											e.printStackTrace();
											StringWriter sw1 = new StringWriter();
											PrintWriter pw = new PrintWriter(sw1);
											e.printStackTrace(pw);
											mf.logger.warning(sw1.toString());

											mf.logger.info(
													"Message size was greater than allowed size so attachment has been deleted and saved in "
															+ f.getAbsolutePath());
											for (MapiAttachment attachment : message.getAttachments()) {

												attachment.save(f.getAbsolutePath() + File.separator
														+ main_multiplefile.getRidOfIllegalFileNameCharacters(
																attachment.getLongFileName()));
											}
											try {
												mess.getAttachments().clear();
												System.out.println(mess.getAttachments().size());
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

											}
										} else if (e.getMessage()
												.equalsIgnoreCase("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("Operation failed.")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage()
														.equalsIgnoreCase("The operation 'AppendMessage' terminated.")
												|| e.getMessage()
														.contains("Software caused connection abort: recv failed")
												|| e.getMessage().contains("Network is unreachable: connect")
												|| e.getMessage().contains("Object has been disposed.")
												|| e.getMessage().contains("Connection or outbound has closed")
												|| e.getMessage()
														.contains("Connection reset by peer: socket write error")) {
											mm.Progressbar.setVisible(false);
											i--;
										}
										connectionHandle();

										e.printStackTrace();
										StringWriter sw1 = new StringWriter();
										PrintWriter pw = new PrintWriter(sw1);
										e.printStackTrace(pw);
										mf.logger.warning(sw1.toString());

										mf.logger.warning("Exception : " + e.getMessage() + "Message " + " " + i
												+ mm.namingconventionmapi(message) + System.lineSeparator());
										continue;
									}

								}
								mm.lbl_progressreport
										.setText("  Total Message Saved Count  " + ConvertPSTOST_gmail.count_destination
												+ "  " + Folder + "   Extracting messsage " + message.getSubject());

							} catch (Exception e) {
								System.out.println(e.getMessage() + "  ===> 1009");
								if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
										|| e.getMessage().contains("ConnectFailure")
										|| e.getMessage().contains("Operation has been canceled")
										|| e.getMessage().contains("Operation failed.")
										|| e.getMessage().contains("Rate limit hit")
										|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
										|| e.getMessage().contains("Software caused connection abort: recv failed")
										|| e.getMessage().contains("Network is unreachable: connect")
										|| e.getMessage().contains("Object has been disposed.")
										|| e.getMessage().contains("Connection or outbound has closed")
										|| e.getMessage().contains("Connection reset by peer: socket write error")) {
									mm.Progressbar.setVisible(false);
									i--;
								}
								connectionHandle();
								continue;
							}

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
				if (filetype.equalsIgnoreCase("GoDaddy email")) {
					Folder = Folder.replaceAll("[^a-zA-Z0-9]", "");

				}

				sfolder = sfolder + File.separator + Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}

					if (pstfolderlist.get(l).equalsIgnoreCase(sfolder)) {
						path11 = path1 + "/" + Folder;
						mm.lbl_progressreport.setText("Getting : " + Folder);
						try {
							clientforimap_output.createFolder(iconnforimap_output, path11);
							clientforimap_output.selectFolder(iconnforimap_output, path11);
						} catch (Exception e2) {
						}
						listdupliccal.clear();
						listduplicacy.clear();
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
//									if ((count_destination % 500) == 0) {
//										if (s22) {
//											connectionHandle1();
//										}
//										s22 = true;
//									}

									MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);
									MapiMessage message1 = pst.extractMessage(messageInfo);
									MapiConversionOptions d = MapiConversionOptions.getUnicodeFormat();
									MailConversionOptions de = new MailConversionOptions();
									MailMessage mess1 = message1.toMailMessage(de);
									MapiMessage message = MapiMessage.fromMailMessage(mess1, d);
									MailMessage mess = message.toMailMessage(de);
									if (mm.chckbxMigrateOrBackup.isSelected()) {
										mess.getAttachments().clear();
									}

									if (main_multiplefile.datefilter.isSelected()) {
										datevalidflag = mm.checkdate(message, mess);
									}
									if (message1.getMessageClass().equals("IPM.Contact")) {
										MailMessage mapi = new MailMessage();
										try {
											MapiContact con = (MapiContact) message1.toMapiMessageItem();
											try {
												mapi.setSubject(con.getSubject());
											} catch (Exception e) {

												mapi.setSubject("");
											}
											try {
												mapi.setBody(con.getBody());
											} catch (Exception e) {

												mapi.setBody("");
											}
											try {
												message1.setSenderEmailAddress(mess.getFrom().toString());
												mapi.setFrom(mess.getFrom());
											} catch (Exception e) {
											}
											con.save(temppathm + File.separator + i + mm.namingconventionmapi(message)
													+ ".vcf", ContactSaveFormat.VCard);
											File file = new File(temppathm + File.separator + i
													+ mm.namingconventionmapi(message) + ".vcf");
											mapi.addAttachment(new Attachment(temppathm + File.separator + i
													+ mm.namingconventionmapi(message) + ".vcf"));
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
																	path11, mapi);
															ConvertPSTOST_gmail.count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, path11,
																mapi);
														ConvertPSTOST_gmail.count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, path11,
																mapi);
														ConvertPSTOST_gmail.count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path11,
															mapi);
													ConvertPSTOST_gmail.count_destination++;
												}
											}

										} catch (Exception e) {
											System.out.println(e.getMessage() + "  ===> 1200");
											if (e.getMessage()
													.equalsIgnoreCase("The operation 'FetchMessage' terminated.")
													|| e.getMessage().contains("ConnectFailure")
													|| e.getMessage().contains("Operation has been canceled")
													|| e.getMessage().contains("Operation failed.")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage()
															.contains("The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Network is unreachable: connect")
													|| e.getMessage().contains("Object has been disposed.")
													|| e.getMessage().contains("Connection or outbound has closed")
													|| e.getMessage()
															.contains("Connection reset by peer: socket write error")) {
												mm.Progressbar.setVisible(false);

												i--;

											}
											e.printStackTrace();
											StringWriter sw1 = new StringWriter();
											PrintWriter pw = new PrintWriter(sw1);
											e.printStackTrace(pw);
											mf.logger.warning(sw1.toString());
											connectionHandle();

											mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
													+ System.lineSeparator());
											continue;
										}

									} else if (message1.getMessageClass().equals("IPM.Appointment")
											|| message1.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message1.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
										MailMessage mapi = new MailMessage();
										try {

											MapiCalendar cal = null;
											try {
												cal = (MapiCalendar) message1.toMapiMessageItem();
											} catch (Exception e1) {

											}

											try {
												mapi.setSubject(cal.getSubject());
											} catch (Exception e) {

												mapi.setSubject("");
											}
											try {
												mapi.setBody(cal.getBody());
											} catch (Exception e) {

												mapi.setBody("");
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

											cal.save(temppathm + File.separator + i + mm.namingconventionmapi(message)
													+ ".ics", AppointmentSaveFormat.Ics);
											File file = new File(temppathm + File.separator + i
													+ mm.namingconventionmapi(message) + ".ics");

											mapi.addAttachment(new Attachment(temppathm + File.separator + i
													+ mm.namingconventionmapi(message) + ".ics"));
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
																	path11, mapi);
															ConvertPSTOST_gmail.count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, path11,
																mapi);
														ConvertPSTOST_gmail.count_destination++;
													}

												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, path11,
																mapi);
														ConvertPSTOST_gmail.count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path11,
															mapi);
													ConvertPSTOST_gmail.count_destination++;
												}
											}
										} catch (Exception e) {
											System.out.println(e.getMessage() + "  ===> 1309");
											if (e.getMessage()
													.equalsIgnoreCase("The operation 'FetchMessage' terminated.")
													|| e.getMessage().contains("ConnectFailure")
													|| e.getMessage().contains("Operation has been canceled")
													|| e.getMessage().contains("Operation failed.")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage()
															.contains("The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Network is unreachable: connect")
													|| e.getMessage().contains("Object has been disposed.")
													|| e.getMessage().contains("Connection or outbound has closed")
													|| e.getMessage()
															.contains("Connection reset by peer: socket write error")) {
												mm.Progressbar.setVisible(false);
												i--;

											}

											e.printStackTrace();
											StringWriter sw1 = new StringWriter();
											PrintWriter pw = new PrintWriter(sw1);
											e.printStackTrace(pw);
											mf.logger.warning(sw1.toString());
											connectionHandle();
											mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " "
													+ System.lineSeparator());
											continue;
										}

									} else if (message1.getMessageClass().equals("IPM.Task")) {
										try {
											MapiTask task = (MapiTask) message1.toMapiMessageItem();

											if (mm.chckbxRemoveDuplicacy.isSelected()) {
												String input = "";
												if (messageInfo.getMessageClass().equals("IPM.Task")) {
													input = mm.duplicacymapiTask(task);
												}
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listduplictask.contains(input)) {
													listduplictask.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag)
															clientforimap_output.appendMessage(iconnforimap_output,
																	path11, mess);
														ConvertPSTOST_gmail.count_destination++;
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, path11,
																mess);
														ConvertPSTOST_gmail.count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag)
														clientforimap_output.appendMessage(iconnforimap_output, path11,
																mess);
													ConvertPSTOST_gmail.count_destination++;
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path11,
															mess);
													ConvertPSTOST_gmail.count_destination++;
												}
											}
										} catch (Exception e) {
											System.out.println(e.getMessage() + "  ===> 1370");
											if (e.getMessage()
													.equalsIgnoreCase("The operation 'FetchMessage' terminated.")
													|| e.getMessage().contains("ConnectFailure")
													|| e.getMessage().contains("Operation has been canceled")
													|| e.getMessage().contains("Operation failed.")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage()
															.contains("The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Network is unreachable: connect")
													|| e.getMessage().contains("Object has been disposed.")
													|| e.getMessage().contains("Connection or outbound has closed")
													|| e.getMessage()
															.contains("Connection reset by peer: socket write error")) {
												mm.Progressbar.setVisible(false);
												i--;
											}

											e.printStackTrace();
											StringWriter sw1 = new StringWriter();
											PrintWriter pw = new PrintWriter(sw1);
											e.printStackTrace(pw);
											mf.logger.warning(sw1.toString());
											connectionHandle();
											mf.logger.warning("Exception : " + e.getMessage() + "Task" + " "
													+ System.lineSeparator());
											continue;
										}

									} else {
										try {
											String messageid = mailimap(mess, path11);
											if (messageid.equals("")) {

												if (((message.getFlags()
														& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
													clientforimap_output.changeMessageFlags(iconnforimap_output,
															messageid, ImapMessageFlags.isRead());
												} else {
													clientforimap_output.removeMessageFlags(iconnforimap_output,
															messageid, ImapMessageFlags.isRead());
												}
											}
										} catch (Exception e) {
											
											System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
											StringWriter sw = new StringWriter();
											e.printStackTrace(new PrintWriter(sw));
											String exceptionAsString = sw.toString();
											System.out.println(e.getMessage() + "  ===> 1407");
											if (exceptionAsString.contains("Message too large")) {
												File f1 = new File(
														(System.getProperty("user.home") + File.separator + "Desktop")
																+ File.separator + calendertime + File.separator
																+ "Attachment" + File.separator
																+ mm.namingconventionmapi(message));
												f1.mkdirs();
												mf.logger.info(
														"Message size was greater than allowed size so attachment has been deleted and saved in "
																+ f1.getAbsolutePath());
												for (MapiAttachment attachment : message.getAttachments()) {

													attachment.save(f1.getAbsolutePath() + File.separator
															+ main_multiplefile.getRidOfIllegalFileNameCharacters(
																	attachment.getLongFileName()));

												}
												try {
													mess.getAttachments().clear();
													String messageid = mailimap(mess, path11);
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
												}

											} else if (e.getMessage()
													.equalsIgnoreCase("The operation 'FetchMessage' terminated.")
													|| e.getMessage().contains("ConnectFailure")
													|| e.getMessage().contains("Operation has been canceled")
													|| e.getMessage().contains("Operation failed.")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage()
															.contains("The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Network is unreachable: connect")
													|| e.getMessage().contains("Object has been disposed.")
													|| e.getMessage().contains("Connection or outbound has closed")
													|| e.getMessage()
															.contains("Connection reset by peer: socket write error")) {
												mm.Progressbar.setVisible(false);

												i--;

											}

											e.printStackTrace();
											StringWriter sw1 = new StringWriter();
											PrintWriter pw = new PrintWriter(sw1);
											e.printStackTrace(pw);
											mf.logger.warning(sw1.toString());
											connectionHandle();
											mf.logger.warning("Exception : " + e.getMessage() + "Message" + " "
													+ System.lineSeparator());
											continue;
										}

									}
									mm.lbl_progressreport.setText("Total Message Saved Count  " + count_destination
											+ "  " + Folder + "   Extracting messsage " + message.getSubject());

								} catch (Exception e) {
									System.out.println(e.getMessage() + "  ===> 1472");
									if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
											|| e.getMessage().contains("ConnectFailure")
											|| e.getMessage().contains("Operation has been canceled")
											|| e.getMessage().contains("Operation failed.")
											|| e.getMessage().contains("Rate limit hit")
											|| e.getMessage()
													.equalsIgnoreCase("The operation 'AppendMessage' terminated.")
											|| e.getMessage().contains("Software caused connection abort: recv failed")
											|| e.getMessage().contains("Network is unreachable: connect")
											|| e.getMessage().contains("Object has been disposed.")
											|| e.getMessage().contains("Connection or outbound has closed")
											|| e.getMessage()
													.contains("Connection reset by peer: socket write error")) {
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
					getsubfolderforpstost_gmail(folderf, sfolder, path11);
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
		
		
		try {
		
		if (mm.chckbxRemoveDuplicacy.isSelected()) {

			String input = mm.duplicacymail(message);

			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
						foldermessagecount++;
						ConvertPSTOST_gmail.count_destination++;
					}
				} else {
					Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
					foldermessagecount++;
					ConvertPSTOST_gmail.count_destination++;
				}
			}
		} else {
			if (main_multiplefile.datefilter.isSelected()) {
				if (datevalidflag) {
					Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
					foldermessagecount++;
					ConvertPSTOST_gmail.count_destination++;
				}
			} else {
				Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
				ConvertPSTOST_gmail.count_destination++;
				System.out.println("message save : " + count_destination );

			}
		}
		}catch (Exception e) {

			System.out.println(" convvertPSTOST_gmail methord calling &&&&&&&&& 1648");
			e.printStackTrace();
			StringWriter sw = new StringWriter();
			e.printStackTrace(new PrintWriter(sw));//
			String exceptionAsString = sw.toString();
			if (exceptionAsString.contains("Message too large") || exceptionAsString.contains("TOOBIG")) {
				File f = new File(
						(System.getProperty("user.home") + File.separator + "Desktop") + File.separator + calendertime
								+ File.separator + "Attachment" + File.separator + mf.namingconventionmail(message));
				f.mkdirs();
				mf.logger.info("Message size was greater than allowed size so attachment has been deleted and saved in "
						+ f.getAbsolutePath());

				MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
				MapiMessage message1 = MapiMessage.fromMailMessage(message, d);

				try {
					message1.save(
							f.getAbsolutePath() + File.separator
									+ Main_Frame.getRidOfIllegalFileNameCharacters(message.getSubject()) + ".msg",
							SaveOptions.getDefaultMsg());
				} catch (Exception e1) {
				}

				for (MapiAttachment attachment : message1.getAttachments()) {

					attachment.save(f.getAbsolutePath() + File.separator
							+ getRidOfIllegalFileNameCharacters(attachment.getLongFileName()));

				}

				message1.getAttachments().clear();
				message1.getAttachments().removeAll(message1.getAttachments());

				MailConversionOptions d1 = new MailConversionOptions();
				message = message1.toMailMessage(d1);

				Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
				count_destination++;
			} else {
				throw e;
			}
		}
		return Messageid;
	}

	public void connectionHandle() {
		System.out.println("Connection Lost We are trying to Connect using of Connection Class.");
		mm.lbl_progressreport.setText("INTERNET Connection  LOST ");
		mm.label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));

		while (true) {
			try {
				mm.lbl_progressreport.setText("Connecting to Server Please Wait");
				if (filetype.equalsIgnoreCase("GMAIL")) {
					if (main_multiplefile.modern_Authentication.isSelected()) {
						String token = GetToken.refreshToken_Gmail_Output();
						if (token != null) {
							System.out.println("this is token ");
							clientforimap_output.dispose();
							clientforimap_output = GetToken.loginGmail_output(token);

						}
					} else {
						System.out.println("this is token output ");
						clientforimap_output.dispose();
						clientforimap_output = connectiontogmail_output();
					}

				} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {
					clientforimap_output.dispose();
					clientforimap_output = connectiontoyahoo_output(clientforimap_output);

				} else if (filetype.equalsIgnoreCase("AOL")) {
					clientforimap_output.dispose();
					clientforimap_output = connectiontoaol_output(clientforimap_output);
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
		clientforimap_output = new ImapClient("imap.gmail.com", 993, username_p3, password_p3);
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
					if (main_multiplefile.modern_Authentication.isSelected()) {
						String token = GetToken.refreshToken_Gmail_Output();
						if (token != null) {
							clientforimap_output.dispose();
							clientforimap_output = GetToken.loginGmail_output(token);
						}
					} else {
						connectiontogmail_output();
					}

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
	
	
	private static String getRidOfIllegalFileNameCharacters(String strName) {
		String strLegalName = strName;
		if (strLegalName.length() >= 80) {
			strLegalName = strLegalName.substring(0, 80);
		}
		strLegalName = strName.replace(":", " ").replace("\\", "").replace("?", "").replace("/", "").replace("|", "")
				.replace("*", "").replace("<", "").replace(">", "").replace("\t", "").replace("\"", "")
				.replace(",", "");

		if (strLegalName.length() >= 80) {
			strLegalName = strLegalName.substring(0, 80);
		}
		return strLegalName;
	}

}
