package email.code;

import java.io.File;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.ImageIcon;

import com.aspose.email.Appointment;
import com.aspose.email.AppointmentLoadOptions;
import com.aspose.email.AppointmentSaveFormat;
import com.aspose.email.Contact;
import com.aspose.email.EWSClient;
import com.aspose.email.EmailClient;
import com.aspose.email.ExchangeFolderInfo;
import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.IConnection;
import com.aspose.email.IEWSClient;
import com.aspose.email.ImapClient;
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
import com.aspose.email.OAuthNetworkCredential;
import com.aspose.email.PersonalStorage;
import com.aspose.email.system.NetworkCredential;

public class ConvertPSTOST_365 implements Runnable {

	private List<String> listdupliccal = new ArrayList<String>();
	private List<String> listduplictask = new ArrayList<String>();
	private List<String> listdupliccontact = new ArrayList<String>();
	List<String> listduplicacy = new ArrayList<String>();
	String parentfolder;
	PersonalStorage ost;
	int splitcount = 0;
	String splitpath = "";
	long foldermessagecount;
	static Date fromdate;
	static Date todate;
	String path3 = "";
	ArrayList<Date> fromList = new ArrayList<Date>();
	ArrayList<Date> toList = new ArrayList<Date>();
	String from;
	String to;
	String first = null, middle = null, last = null;
	private Main_Frame mf;
	private String path = "";
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
	boolean check = true;
	IEWSClient clientforexchange_output;
	static ImapClient clientforimap_output;
	static IConnection iconnforimap_output;
	String path1 = "";
	long maxsize = 0;
	int pstindex = 0;
	String Folderuri;
	String fa = "";
	String fname = "";
	static String mailboxUri = "https://outlook.office365.com/EWS/Exchange.asmx";
	String username_p3;
	String password_p3;
	static String calendar = "";

	public ConvertPSTOST_365(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm, String username_p3, String password_p3, String Folderuri,
			IEWSClient clientforexchange_output, String fa, String fname) {
		this.mf = mf;
		this.mm = mm;
		this.filetype = filetype;
		this.filepath = filepath;
		this.destination_path = destination_path;
		ConvertPSTOST_365.count_destination = count_destination;
		this.pstfolderlist = pstfolderlist;
		this.fromList = fromList;
		this.toList = toList;
		this.temppathm = temppathm;
		this.username_p3 = username_p3;
		this.password_p3 = password_p3;
		this.Folderuri = Folderuri;
		this.fa = fa;
		this.fname = fname;
		this.clientforexchange_output = clientforexchange_output;
	}

	@Override
	public void run() {
		convertPSTOST_365(mf, filetype, destination_path, count_destination, filepath, mm, pstfolderlist, fromList,
				toList, temppathm, Folderuri, clientforexchange_output, fa, fname);
		main_multiplefile.count_destination = ConvertPSTOST_365.count_destination;
		main_multiplefile.fa = fa;
		main_multiplefile.fname = fname;
		System.out.println(ConvertPSTOST_365.count_destination);
		Folderuri = fa;
	}

	@SuppressWarnings("deprecation")
	private void convertPSTOST_365(Main_Frame mf, String filetype, String destination_path2, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm, String Folderuri, IEWSClient clientforexchange_output, String fa,
			String fname) {

		ConvertPSTOST_365.count_destination = 0;
		Folderuri = clientforexchange_output.createFolder(Folderuri, fname).getUri();
		pst = PersonalStorage.fromFile(filepath);
		foldermessagecount = 0;
		FolderInfo folderInfo2 = pst.getRootFolder();
		Folder = folderInfo2.getDisplayName();
		Folder = Folder.replace(",", "").replace(".", "");
		Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
		Folder = Folder.replaceAll("[\\[\\]]", "");
		Folder = Folder.trim();
		if (Folder.equalsIgnoreCase("")) {
			Folder = "Root Folder";
		}
		path = Folder;
		path1 = Folder;

		//
		String name1 = "";
		if (mm.chckbxCustomFolderName.isSelected()) {
			String customerfolder = mm.textField_customfolder.getText().replace("//s", "");
			name1 = customerfolder + "_" + fname + "_" + "Calendar";
		} else {
			name1 = mm.calendertime + "_" + fname + "_" + "Calendar";
		}

		ExchangeFolderInfo subfolderInfo1[] = new ExchangeFolderInfo[] { null };
		if (!clientforexchange_output.folderExists(clientforexchange_output.getMailboxInfo().getCalendarUri(), name1,
				subfolderInfo1)) {
			calendar = clientforexchange_output.createFolder(clientforexchange_output.getMailboxInfo().getCalendarUri(),
					name1, null, "IPF.Appointment").getUri();
		}
		//

		MessageInfoCollection messageInfoCollection1 = folderInfo2.getContents();
		Folderuri = clientforexchange_output.createFolder(Folderuri, path).getUri();
		parentfolder = Folderuri;
		int messagesize1;
		listduplicacy.clear();
		listdupliccal.clear();
		listdupliccontact.clear();
		listduplictask.clear();
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
		String mailfolder = "";
		for (int i = 0; i < messagesize1; i++) {

			if (mm.stop) {
				break;
			}
			if ((i % 100) == 0) {
				System.gc();
			}
			try {
				MessageInfo messageInfo = (MessageInfo) messageInfoCollection1.get_Item(i);

				MapiMessage message1 = pst.extractMessage(messageInfo);
				MapiConversionOptions d = MapiConversionOptions.getUnicodeFormat();
				MailConversionOptions de = new MailConversionOptions();
				MailMessage mess1 = message1.toMailMessage(de);
				if (mm.chckbxMigrateOrBackup.isSelected()) {
					mess1.getAttachments().clear();
				}
				MapiMessage message = MapiMessage.fromMailMessage(mess1, d);
				if (main_multiplefile.datefilter.isSelected()) {
					datevalidflag = mm.checkdate(message, mess1);
				}
				if (message.getMessageClass().equals("IPM.Contact")) {

					try {
						ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
						if (!clientforexchange_output.folderExists(
								clientforexchange_output.getMailboxInfo().getContactsUri(),
								mm.calendertime + "_" + fname + "/" + Folder, subfolderInfo)) {
							mailfolder = clientforexchange_output
									.createFolder(clientforexchange_output.getMailboxInfo().getContactsUri(),
											mm.calendertime + "_" + fname + "/" + Folder, null, "IPF.Contact")
									.getUri();
						}

						MapiContact con = (MapiContact) message.toMapiMessageItem();
						Contact conn = Contact.to_Contact(con);
						if (mm.chckbxRemoveDuplicacy.isSelected()) {

							String input = mm.duplicacymapiContact(con);
							input = input.replaceAll("\\s", "");
							input = input.trim();
							if (!listdupliccontact.contains(input)) {
								listdupliccontact.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										try {
											clientforexchange_output.createContact(mailfolder, con);
											ConvertPSTOST_365.count_destination++;
											foldermessagecount++;
										} catch (Exception e) {
											clientforexchange_output.createContact(mailfolder, conn);
											ConvertPSTOST_365.count_destination++;
											foldermessagecount++;
										}
									}
								} else {
									try {
										clientforexchange_output.createContact(mailfolder, con);
										ConvertPSTOST_365.count_destination++;
										foldermessagecount++;
									} catch (Exception e) {
										clientforexchange_output.createContact(mailfolder, conn);
										ConvertPSTOST_365.count_destination++;
										foldermessagecount++;
									}
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									try {
										clientforexchange_output.createContact(mailfolder, con);
										ConvertPSTOST_365.count_destination++;
										foldermessagecount++;
									} catch (Exception e) {
										clientforexchange_output.createContact(mailfolder, conn);
										ConvertPSTOST_365.count_destination++;
										foldermessagecount++;
									}
								}
							} else {
								try {
									clientforexchange_output.createContact(mailfolder, con);
									ConvertPSTOST_365.count_destination++;
									foldermessagecount++;
								} catch (Exception e) {
									clientforexchange_output.createContact(mailfolder, conn);
									ConvertPSTOST_365.count_destination++;
									foldermessagecount++;
								}
							}
						}
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mm.namingconventionmapi(message));
					} catch (Exception e) {
						if(e.getMessage().contains("ERROR")
								|| e.getMessage().contains("ERROR_CORRUPT_DATA")
								|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
								) {
							continue;
						}
						else if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
								|| e.getMessage().contains(
										"No connection could be made because the target machine actively refused it.")
								|| e.getMessage().contains("Bad response")
								
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage()
										.contains("An existing connection was forcibly closed by the remote host.")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("Operation has been canceled")
								|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
								|| e.getMessage().contains("Cannot access a disposed object.")
								|| e.getMessage().contains(
										"The support for the specified socket type does not exist in this address family.")) {
							mm.Progressbar.setVisible(false);

							i--;
						}
						connectionHandle(e.getMessage());
						mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
								+ mm.namingconventionmapi(message) + System.lineSeparator());
						continue;
					}

				} else if (message.getMessageClass().equals("IPM.Appointment")
						|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")){

					try {

						MapiCalendar cal = null;
						Appointment calDoc = null;
						File file = null;
						try {
							cal = (MapiCalendar) message.toMapiMessageItem();
							cal.save(temppathm + File.separator + mm.namingconventionmapi(message) + ".ics",
									AppointmentSaveFormat.Ics);
							file = new File(temppathm + File.separator + mm.namingconventionmapi(message) + ".ics");
							AppointmentLoadOptions optiona = new AppointmentLoadOptions();
							optiona.getIgnoreSmtpAddressCheck();
							calDoc = Appointment.load(
									temppathm + File.separator + mm.namingconventionmapi(message) + ".ics", optiona);
						} catch (Exception e) {
						}
						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = mm.duplicacymapiCal(cal);
							input = input.replaceAll("\\s", "");
							input = input.trim();
							if (!listdupliccal.contains(input)) {
								listdupliccal.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										try {
											clientforexchange_output.appendMessage(calendar, mess1);
										} catch (Exception e) {
											clientforexchange_output.createAppointment(calDoc, calendar);
										}
										ConvertPSTOST_365.count_destination++;
									}
								} else {
									try {
										clientforexchange_output.appendMessage(calendar, mess1);
									} catch (Exception e) {
										clientforexchange_output.createAppointment(calDoc, calendar);
									}
									ConvertPSTOST_365.count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									try {
										clientforexchange_output.appendMessage(calendar, mess1);
									} catch (Exception e) {
										clientforexchange_output.createAppointment(calDoc, calendar);
									}
									ConvertPSTOST_365.count_destination++;
								}
							} else {
								try {
									clientforexchange_output.appendMessage(calendar, mess1);
								} catch (Exception e) {
									clientforexchange_output.createAppointment(calDoc, calendar);
								}
								ConvertPSTOST_365.count_destination++;
							}
						}
						file.delete();
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mm.namingconventionmapi(message));
					} catch (Exception e) {
						if(e.getMessage().contains("ERROR")
								|| e.getMessage().contains("ERROR_CORRUPT_DATA")
								|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
								) {
							continue;
						}
						else if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
								|| e.getMessage().contains(
										"No connection could be made because the target machine actively refused it.")
								|| e.getMessage().contains("Bad response")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage()
										.contains("An existing connection was forcibly closed by the remote host.")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("Operation has been canceled")
								|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
								|| e.getMessage().contains("Cannot access a disposed object.")
								|| e.getMessage().contains(
										"The support for the specified socket type does not exist in this address family.")) {
							mm.Progressbar.setVisible(false);

							i--;
						}
						connectionHandle(e.getMessage());
						mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + i
								+ mm.namingconventionmapi(message) + System.lineSeparator());
						continue;
					}

				} else if (message.getMessageClass().equals("IPM.Task")) {
					try {
						MapiTask task = (MapiTask) message.toMapiMessageItem();
						MailConversionOptions options = new MailConversionOptions();
						options.setConvertAsTnef(true);
						String taskuri = clientforexchange_output.getMailboxInfo().getTasksUri();
						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = "";
							if (message.getMessageClass().equals("IPM.Task")) {
								input = mm.duplicacymapiTask(task);
							}
							input = input.replaceAll("\\s", "");
							input = input.trim();
							if (!listduplictask.contains(input)) {
								listduplictask.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										clientforexchange_output.createTask(taskuri, task);
										ConvertPSTOST_365.count_destination++;
										foldermessagecount++;
									}
								} else {
									clientforexchange_output.createTask(taskuri, task);
									ConvertPSTOST_365.count_destination++;
									foldermessagecount++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforexchange_output.createTask(taskuri, task);
									ConvertPSTOST_365.count_destination++;
									foldermessagecount++;
								}
							} else {
								clientforexchange_output.createTask(taskuri, task);
								ConvertPSTOST_365.count_destination++;
								foldermessagecount++;
							}
						}
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mm.namingconventionmapi(message));
					} catch (Exception e) {
						if(e.getMessage().contains("ERROR")
								|| e.getMessage().contains("ERROR_CORRUPT_DATA")
								|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
								) {
							continue;
						}
						else if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
								|| e.getMessage().contains(
										"No connection could be made because the target machine actively refused it.")
								|| e.getMessage().contains("Bad response")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage()
										.contains("An existing connection was forcibly closed by the remote host.")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("Operation has been canceled")
								|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
								|| e.getMessage().contains("Cannot access a disposed object.")
								|| e.getMessage().contains(
										"The support for the specified socket type does not exist in this address family.")) {
							mm.Progressbar.setVisible(false);

							i--;
						}
						connectionHandle(e.getMessage());
						mf.logger.warning("Exception : " + e.getMessage() + "Task" + " " + i
								+ mm.namingconventionmapi(message) + System.lineSeparator());
						continue;
					}

				} else {
					try {
						String Messageid = null;
						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = mm.duplicacymapi(message);
							if (!listduplicacy.contains(input)) {
								listduplicacy.add(input);

								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										Messageid = clientforexchange_output.appendMessage(Folderuri, mess1);
										ConvertPSTOST_365.count_destination++;
										foldermessagecount++;
									}
								} else {
									Messageid = clientforexchange_output.appendMessage(Folderuri, mess1);
									ConvertPSTOST_365.count_destination++;
									foldermessagecount++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									Messageid = clientforexchange_output.appendMessage(Folderuri, mess1);
									ConvertPSTOST_365.count_destination++;
									foldermessagecount++;
								}
							} else {
								Messageid = clientforexchange_output.appendMessage(Folderuri, mess1);
								ConvertPSTOST_365.count_destination++;
								foldermessagecount++;
							}
						}
						if (Messageid != null) {
							if (((message.getFlags()
									& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
								clientforexchange_output.setReadFlag(Messageid, true);

							} else {
								clientforexchange_output.setReadFlag(Messageid, false);
							}
						}
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mm.namingconventionmapi(message));
					} catch (Exception e) {
						System.out.println(e.getMessage());
						StringWriter sw = new StringWriter();
						e.printStackTrace(new PrintWriter(sw));
						String exceptionAsString = sw.toString();
						int number = message.getAttachments().size();
						if (exceptionAsString.contains("Message too large")
								|| exceptionAsString.contains("The message exceeds the maximum supported size.")
								|| number > 10) {
							File f11 = new File((System.getProperty("user.home") + File.separator + "Desktop")
									+ File.separator + fname + File.separator + "Attachment" + File.separator
									+ mm.namingconventionmapi(message));
							f11.mkdirs();
							mf.logger.info(
									"Message size was greater than allowed size so attachment has been deleted and saved in "
											+ f11.getAbsolutePath());
							for (MapiAttachment attachment : message.getAttachments()) {

								attachment.save(f11.getAbsolutePath() + File.separator + main_multiplefile
										.getRidOfIllegalFileNameCharacters(attachment.getLongFileName()));

							}
							try {
								mess1.getAttachments().clear();
								//

								String Messageid = null;
								if (mm.chckbxRemoveDuplicacy.isSelected()) {
									String input = mm.duplicacymapi(message);
									if (!listduplicacy.contains(input)) {
										listduplicacy.add(input);

										if (main_multiplefile.datefilter.isSelected()) {
											if (datevalidflag) {
												Messageid = clientforexchange_output.appendMessage(Folderuri, mess1);
												ConvertPSTOST_365.count_destination++;
												foldermessagecount++;
											}
										} else {
											Messageid = clientforexchange_output.appendMessage(Folderuri, mess1);
											ConvertPSTOST_365.count_destination++;
											foldermessagecount++;
										}
									}
								} else {
									if (main_multiplefile.datefilter.isSelected()) {
										if (datevalidflag) {
											Messageid = clientforexchange_output.appendMessage(Folderuri, mess1);
											ConvertPSTOST_365.count_destination++;
											foldermessagecount++;
										}
									} else {
										Messageid = clientforexchange_output.appendMessage(Folderuri, mess1);
										ConvertPSTOST_365.count_destination++;
										foldermessagecount++;
									}
								}
								if (Messageid != null) {
									if (((message.getFlags()
											& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
										clientforexchange_output.setReadFlag(Messageid, true);

									} else {
										clientforexchange_output.setReadFlag(Messageid, false);
									}
								}
							} catch (Exception e1) {
							}

						}
						else	if(e.getMessage().contains("ERROR")
								|| e.getMessage().contains("ERROR_CORRUPT_DATA")
								|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
								) {
							continue;
						}
						else if (e.getMessage().equalsIgnoreCase("The operation 'FetchMessage' terminated.")
								|| e.getMessage().contains(
										"No connection could be made because the target machine actively refused it.")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage()
										.contains("An existing connection was forcibly closed by the remote host.")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("Bad response")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("Operation has been canceled")
								|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
								|| e.getMessage().contains("Cannot access a disposed object.")
								|| e.getMessage().contains(
										"The support for the specified socket type does not exist in this address family.")) {
							mm.Progressbar.setVisible(false);

							i--;
						}
						connectionHandle(e.getMessage());
						mf.logger.warning("Exception : " + e.getMessage() + "Message " + " " + i
								+ mm.namingconventionmapi(message) + System.lineSeparator());
						continue;
					}

				}

				mm.lbl_progressreport.setText("  Total Message Saved Count  " + ConvertPSTOST_365.count_destination
						+ "  " + Folder + "   Extracting messsage " + message.getSubject());
				System.out.println("  Total Message Saved Count  " + ConvertPSTOST_365.count_destination + "  " + Folder
						+ "   Extracting messsage " + message.getSubject());
			} catch (Exception e) {
				continue;
			}

		}
		
		for(int i=0;i<pstfolderlist.size();i++) {
			System.out.println(pstfolderlist.get(i)+" folder name");
		}

		FolderInfoCollection folderInf = pst.getRootFolder().getSubFolders();

		for (int j = 0; j < folderInf.size(); j++) {
			try {

				if (mm.stop) {
					break;
				}
				String subfolder = Folderuri;
				FolderInfo folderInfo = folderInf.get_Item(j);
				String Folder = folderInfo.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();

				path = path1 + File.separator + Folder;
               
				
				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}
					if (pstfolderlist.get(l).equalsIgnoreCase(path)) {
                        
						System.out.println("path matched "+path);
						mm.lbl_progressreport.setText(" Getting Folder " + Folder);

						subfolder = clientforexchange_output.createFolder(subfolder, Folder).getUri();

						MessageInfoCollection messageInfoCollection = folderInfo.getContents();
						listduplicacy.clear();
						listdupliccal.clear();
						listdupliccontact.clear();
						listduplictask.clear();
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
							if (mm.stop) {
								break;
							}
							if ((i % 100) == 0) {
								System.gc();
							}
							try {
								MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);
								MapiMessage message1 = pst.extractMessage(messageInfo);
								MapiConversionOptions d = MapiConversionOptions.getUnicodeFormat();
								MailConversionOptions de = new MailConversionOptions();
								MailMessage mess = message1.toMailMessage(de);
								if (mm.chckbxMigrateOrBackup.isSelected()) {
									mess.getAttachments().clear();
								}
								MapiMessage message = MapiMessage.fromMailMessage(mess, d);

								if (main_multiplefile.datefilter.isSelected()) {
									datevalidflag = mm.checkdate(message, mess);
								}
								if (message.getMessageClass().equals("IPM.Contact")) {
									try {
										ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
										if (!clientforexchange_output.folderExists(
												clientforexchange_output.getMailboxInfo().getContactsUri(),
												mm.calendertime + "_" + fname + "/" + Folder, subfolderInfo)) {
											mailfolder = clientforexchange_output.createFolder(
													clientforexchange_output.getMailboxInfo().getContactsUri(),
													mm.calendertime + "_" + fname + "/" + Folder, null, "IPF.Contact")
													.getUri();
										}

										MapiContact con = (MapiContact) message.toMapiMessageItem();
										Contact conn = Contact.to_Contact(con);
										if (mm.chckbxRemoveDuplicacy.isSelected()) {

											String input = mm.duplicacymapiContact(con);
											input = input.replaceAll("\\s", "");
											input = input.trim();

											if (!listdupliccontact.contains(input)) {
												listdupliccontact.add(input);
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														try {
															clientforexchange_output.createContact(mailfolder, con);
															ConvertPSTOST_365.count_destination++;
															foldermessagecount++;
														} catch (Exception e) {
															clientforexchange_output.createContact(mailfolder, conn);
															ConvertPSTOST_365.count_destination++;
															foldermessagecount++;
														}
													}
												} else {
													try {
														clientforexchange_output.createContact(mailfolder, con);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													} catch (Exception e) {
														clientforexchange_output.createContact(mailfolder, conn);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													}
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													try {
														clientforexchange_output.createContact(mailfolder, con);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													} catch (Exception e) {
														clientforexchange_output.createContact(mailfolder, conn);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													}
												}
											} else {
												try {
													clientforexchange_output.createContact(mailfolder, con);
													ConvertPSTOST_365.count_destination++;
													foldermessagecount++;
												} catch (Exception e) {
													clientforexchange_output.createContact(mailfolder, conn);
													ConvertPSTOST_365.count_destination++;
													foldermessagecount++;
												}
											}
										}
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										if(e.getMessage().contains("ERROR")
												|| e.getMessage().contains("ERROR_CORRUPT_DATA")
												|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
												) {
											continue;
										}
										else	if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Bad response")
												|| e.getMessage().contains(
														"An existing connection was forcibly closed by the remote host.")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
												|| e.getMessage().contains("Cannot access a disposed object.")
												|| e.getMessage().contains(
														"The support for the specified socket type does not exist in this address family.")) {
											mm.Progressbar.setVisible(false);
											i--;
										}

										connectionHandle(e.getMessage());
										mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
												+ mm.namingconventionmapi(message) + System.lineSeparator());
										continue;
									}
								} else if (message.getMessageClass().equals("IPM.Appointment")
										|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
									try {

										MapiCalendar cal = null;
										Appointment calDoc = null;
										File file = null;
										try {
											cal = (MapiCalendar) message.toMapiMessageItem();
											cal.save(temppathm + File.separator + mm.namingconventionmapi(message)
													+ ".ics", AppointmentSaveFormat.Ics);
											file = new File(temppathm + File.separator
													+ mm.namingconventionmapi(message) + ".ics");
											AppointmentLoadOptions optiona = new AppointmentLoadOptions();
											optiona.getIgnoreSmtpAddressCheck();
											calDoc = Appointment.load(temppathm + File.separator
													+ mm.namingconventionmapi(message) + ".ics", optiona);
										} catch (Exception e) {
										}

										if (mm.chckbxRemoveDuplicacy.isSelected()) {
											String input = mm.duplicacymapiCal(cal);
											input = input.replaceAll("\\s", "");
											input = input.trim();
											if (!listdupliccal.contains(input)) {
												listdupliccal.add(input);
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														try {
															clientforexchange_output.appendMessage(calendar, mess);
														} catch (Exception e) {
															clientforexchange_output.createAppointment(calDoc,
																	calendar);
														}
														ConvertPSTOST_365.count_destination++;
													}
												} else {
													try {
														clientforexchange_output.appendMessage(calendar, mess);
													} catch (Exception e) {
														clientforexchange_output.createAppointment(calDoc, calendar);
													}
													ConvertPSTOST_365.count_destination++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													try {
														clientforexchange_output.appendMessage(calendar, mess);
													} catch (Exception e) {
														clientforexchange_output.createAppointment(calDoc, calendar);
													}
													ConvertPSTOST_365.count_destination++;
												}
											} else {
												try {
													clientforexchange_output.appendMessage(calendar, mess);
												} catch (Exception e) {
													clientforexchange_output.createAppointment(calDoc, calendar);
												}
												ConvertPSTOST_365.count_destination++;
											}
										}
										file.delete();
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										if(e.getMessage().contains("ERROR")
												|| e.getMessage().contains("ERROR_CORRUPT_DATA")
												|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
												) {
											continue;
										}
										else if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Bad response")
												|| e.getMessage().contains(
														"An existing connection was forcibly closed by the remote host.")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
												|| e.getMessage().contains("Cannot access a disposed object.")
												|| e.getMessage().contains(
														"The support for the specified socket type does not exist in this address family.")) {
											mm.Progressbar.setVisible(false);
											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + i
												+ mm.namingconventionmapi(message) + System.lineSeparator());
										continue;
									}

								} else if (message.getMessageClass().equals("IPM.Task")) {
									try {
										MapiTask task = (MapiTask) message.toMapiMessageItem();
										MailConversionOptions options = new MailConversionOptions();
										options.setConvertAsTnef(true);
										String taskuri = clientforexchange_output.getMailboxInfo().getTasksUri();
										if (mm.chckbxRemoveDuplicacy.isSelected()) {
											String input = "";
											if (message.getMessageClass().equals("IPM.Task")) {
												input = mm.duplicacymapiTask(task);
											}
											input = input.replaceAll("\\s", "");
											input = input.trim();
											if (!listduplictask.contains(input)) {
												listduplictask.add(input);
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforexchange_output.createTask(taskuri, task);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													}
												} else {
													clientforexchange_output.createTask(taskuri, task);
													ConvertPSTOST_365.count_destination++;
													foldermessagecount++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													clientforexchange_output.createTask(taskuri, task);
													ConvertPSTOST_365.count_destination++;
													foldermessagecount++;
												}
											} else {
												clientforexchange_output.createTask(taskuri, task);
												ConvertPSTOST_365.count_destination++;
												foldermessagecount++;
											}
										}
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										if(e.getMessage().contains("ERROR")
												|| e.getMessage().contains("ERROR_CORRUPT_DATA")
												|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
												) {
											continue;
										}
										else if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("Bad response")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains(
														"An existing connection was forcibly closed by the remote host.")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
												|| e.getMessage().contains("Cannot access a disposed object.")
												|| e.getMessage().contains(
														"The support for the specified socket type does not exist in this address family.")) {
											mm.Progressbar.setVisible(false);

											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning("Exception : " + e.getMessage() + "Task" + " " + i
												+ mm.namingconventionmapi(message) + System.lineSeparator());
										continue;
									}
								} else {
									try {
										String Messageid = null;
										if (mm.chckbxRemoveDuplicacy.isSelected()) {
											String input = mm.duplicacymapi(message);
											if (!listduplicacy.contains(input)) {
												listduplicacy.add(input);

												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														Messageid = clientforexchange_output.appendMessage(subfolder,
																mess);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													}
												} else {
													Messageid = clientforexchange_output.appendMessage(subfolder, mess);
													ConvertPSTOST_365.count_destination++;
													foldermessagecount++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													Messageid = clientforexchange_output.appendMessage(subfolder, mess);
													ConvertPSTOST_365.count_destination++;
													foldermessagecount++;
												}
											} else {
												Messageid = clientforexchange_output.appendMessage(subfolder, mess);
												ConvertPSTOST_365.count_destination++;
												foldermessagecount++;
											}
										}
										if (Messageid != null) {
											if (((message.getFlags()
													& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
												clientforexchange_output.setReadFlag(Messageid, true);

											} else {
												clientforexchange_output.setReadFlag(Messageid, false);
											}
										}
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										System.out.println(e.getMessage());
										StringWriter sw = new StringWriter();
										e.printStackTrace(new PrintWriter(sw));
										String exceptionAsString = sw.toString();
										int number = message.getAttachments().size();
										if (exceptionAsString.contains("Message too large")
												|| exceptionAsString
														.contains("The message exceeds the maximum supported size.")
												|| number > 10) {
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
												//
												String Messageid = null;
												if (mm.chckbxRemoveDuplicacy.isSelected()) {
													String input = mm.duplicacymapi(message);
													if (!listduplicacy.contains(input)) {
														listduplicacy.add(input);

														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																Messageid = clientforexchange_output
																		.appendMessage(subfolder, mess);
																ConvertPSTOST_365.count_destination++;
																foldermessagecount++;
															}
														} else {
															Messageid = clientforexchange_output
																	.appendMessage(subfolder, mess);
															ConvertPSTOST_365.count_destination++;
															foldermessagecount++;
														}
													}
												} else {
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															Messageid = clientforexchange_output
																	.appendMessage(subfolder, mess);
															ConvertPSTOST_365.count_destination++;
															foldermessagecount++;
														}
													} else {
														Messageid = clientforexchange_output.appendMessage(subfolder,
																mess);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													}
												}
												if (Messageid != null) {
													if (((message.getFlags()
															& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
														clientforexchange_output.setReadFlag(Messageid, true);

													} else {
														clientforexchange_output.setReadFlag(Messageid, false);
													}
												}
												//

											} catch (Exception e1) {
											}

										}
										else	if(e.getMessage().contains("ERROR")
												|| e.getMessage().contains("ERROR_CORRUPT_DATA")
												|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
												) {
											continue;
										}
										else if (e.getMessage()
												.equalsIgnoreCase("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains(
														"An existing connection was forcibly closed by the remote host.")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Bad response")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
												|| e.getMessage().contains("Cannot access a disposed object.")
												|| e.getMessage().contains(
														"The support for the specified socket type does not exist in this address family.")
												|| e.getMessage().contains(
														"An existing connection was forcibly closed by the remote host.")) {
											mm.Progressbar.setVisible(false);

											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning("Exception : " + e.getMessage() + "Message " + " " + i
												+ mm.namingconventionmapi(message) + System.lineSeparator());
										continue;
									}

								}

								mm.lbl_progressreport
										.setText("  Total Message Saved Count  " + ConvertPSTOST_365.count_destination
												+ "  " + Folder + "   Extracting messsage " + message.getSubject());
								System.out.println("  Total Message Saved Count  " + ConvertPSTOST_365.count_destination
										+ "  " + Folder + "   Extracting messsage " + message.getSubject());
							} catch (Exception e) {
								continue;
							}

						}
					}
				}
				
				if (folderInfo.hasSubFolders()) {
					getsubfolderpstost_exchange(folderInfo, subfolder);

				}
			} catch (Exception e) {
				continue;
			}

		}

	}

	@SuppressWarnings("deprecation")
	void getsubfolderpstost_exchange(FolderInfo folderi, String p) {

		FolderInfoCollection subfolder1 = folderi.getSubFolders();
		for (int k = 0; k < subfolder1.size(); k++) {

			try {
				if (mm.stop) {
					break;
				}
				FolderInfo folderInfo = subfolder1.get_Item(k);
				String Folder = folderInfo.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();

				String mailfolder = "";
				path = path + File.separator + Folder;

				String subfolder = p;
				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}
					Map<String, String> dest_folder_path = new HashMap<>();
					
					if (pstfolderlist.get(l).equalsIgnoreCase(path)) {
						mm.lbl_progressreport.setText("Getting : " + Folder);

						try {
							if(clientforexchange_output.folderExists(p, Folder)) {
								subfolder = dest_folder_path.get(Folder);
							}else {
								subfolder = clientforexchange_output.createFolder(p, Folder).getUri();
								dest_folder_path.put( Folder,p);
							}
							
						
						} catch (Exception e2) {
							// TODO Auto-generated catch block
							e2.printStackTrace();
//							subfolder = clientforexchange_output.createFolder(p, Folder+"-"+l).getUri();
//						continue;
						}
						
						
						listduplicacy.clear();
						listdupliccal.clear();
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
								MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);
								MapiMessage message1 = pst.extractMessage(messageInfo);
								MapiConversionOptions d = MapiConversionOptions.getUnicodeFormat();
								MailConversionOptions de = new MailConversionOptions();
								MailMessage mess = message1.toMailMessage(de);
								if (mm.chckbxMigrateOrBackup.isSelected()) {
									mess.getAttachments().clear();
								}
								MapiMessage message = MapiMessage.fromMailMessage(mess, d);
								if (main_multiplefile.datefilter.isSelected()) {
									datevalidflag = mm.checkdate(message, mess);
								}
								if (message.getMessageClass().equals("IPM.Contact")) {
									try {
										ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
										if (!clientforexchange_output.folderExists(
												clientforexchange_output.getMailboxInfo().getContactsUri(),
												mm.calendertime + "_" + fname + "/" + Folder, subfolderInfo)) {
											mailfolder = clientforexchange_output.createFolder(
													clientforexchange_output.getMailboxInfo().getContactsUri(),
													mm.calendertime + "_" + fname + "/" + Folder, null, "IPF.Contact")
													.getUri();
										}
										MapiContact con = (MapiContact) message.toMapiMessageItem();
										Contact conn = Contact.to_Contact(con);
										if (mm.chckbxRemoveDuplicacy.isSelected()) {
											String input = mm.duplicacymapiContact(con);
											input = input.replaceAll("\\s", "");
											input = input.trim();

											if (!listdupliccontact.contains(input)) {
												listdupliccontact.add(input);
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														try {
															clientforexchange_output.createContact(mailfolder, con);
															ConvertPSTOST_365.count_destination++;
															foldermessagecount++;
														} catch (Exception e) {
															clientforexchange_output.createContact(mailfolder, conn);
															ConvertPSTOST_365.count_destination++;
															foldermessagecount++;
														}

													}
												} else {
													try {
														clientforexchange_output.createContact(mailfolder, con);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													} catch (Exception e) {
														clientforexchange_output.createContact(mailfolder, conn);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													}

												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													try {
														clientforexchange_output.createContact(mailfolder, con);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													} catch (Exception e) {
														clientforexchange_output.createContact(mailfolder, conn);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													}
												}
											} else {
												try {
													clientforexchange_output.createContact(mailfolder, con);
													ConvertPSTOST_365.count_destination++;
													foldermessagecount++;
												} catch (Exception e) {
													clientforexchange_output.createContact(mailfolder, conn);
													ConvertPSTOST_365.count_destination++;
													foldermessagecount++;
												}
											}
										}
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										if(e.getMessage().contains("ERROR")
												|| e.getMessage().contains("ERROR_CORRUPT_DATA")
												|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
												) {
											continue;
										}
										else	if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Bad response")
												|| e.getMessage().contains(
														"An existing connection was forcibly closed by the remote host.")
												|| e.getMessage().contains("Rate limit hit")

												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
												|| e.getMessage().contains("Cannot access a disposed object.")
												|| e.getMessage().contains(
														"The support for the specified socket type does not exist in this address family.")) {
											mm.Progressbar.setVisible(false);
											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
												+ mm.namingconventionmapi(message) + System.lineSeparator());
										continue;
									}

								} else if (message.getMessageClass().equals("IPM.Appointment")
										|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {

									try {

										MapiCalendar cal = null;
										Appointment calDoc = null;
										File file = null;
										try {
											cal = (MapiCalendar) message.toMapiMessageItem();
											cal.save(temppathm + File.separator + mm.namingconventionmapi(message)
													+ ".ics", AppointmentSaveFormat.Ics);
											file = new File(temppathm + File.separator
													+ mm.namingconventionmapi(message) + ".ics");
											AppointmentLoadOptions optiona = new AppointmentLoadOptions();
											optiona.getIgnoreSmtpAddressCheck();
											calDoc = Appointment.load(temppathm + File.separator
													+ mm.namingconventionmapi(message) + ".ics", optiona);
										} catch (Exception e) {
										}
										if (mm.chckbxRemoveDuplicacy.isSelected()) {
											String input = mm.duplicacymapiCal(cal);
											input = input.replaceAll("\\s", "");
											input = input.trim();
											if (!listdupliccal.contains(input)) {
												listdupliccal.add(input);
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														try {
															clientforexchange_output.appendMessage(calendar, mess);
														} catch (Exception e) {
															clientforexchange_output.createAppointment(calDoc,
																	calendar);
														}
														ConvertPSTOST_365.count_destination++;
													}
												} else {
													try {
														clientforexchange_output.appendMessage(calendar, mess);
													} catch (Exception e) {
														clientforexchange_output.createAppointment(calDoc, calendar);
													}
													ConvertPSTOST_365.count_destination++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													try {
														clientforexchange_output.appendMessage(calendar, mess);
													} catch (Exception e) {
														clientforexchange_output.createAppointment(calDoc, calendar);
													}
													ConvertPSTOST_365.count_destination++;
												}
											} else {
												try {
													clientforexchange_output.appendMessage(calendar, mess);
												} catch (Exception e) {
													clientforexchange_output.createAppointment(calDoc, calendar);
												}
												ConvertPSTOST_365.count_destination++;
											}
										}
										file.delete();
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										if(e.getMessage().contains("ERROR")
												|| e.getMessage().contains("ERROR_CORRUPT_DATA")
												|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
												) {
											continue;
										}
										else if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("Bad response")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains(
														"An existing connection was forcibly closed by the remote host.")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
												|| e.getMessage().contains("Cannot access a disposed object.")
												|| e.getMessage().contains(
														"The support for the specified socket type does not exist in this address family.")) {
											mm.Progressbar.setVisible(false);

											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + i
												+ mm.namingconventionmapi(message) + System.lineSeparator());
										continue;
									}

								} else if (message.getMessageClass().equals("IPM.Task")) {
									try {
										MapiTask task = (MapiTask) message.toMapiMessageItem();
										MailConversionOptions options = new MailConversionOptions();
										options.setConvertAsTnef(true);
										String taskuri = clientforexchange_output.getMailboxInfo().getTasksUri();
										if (mm.chckbxRemoveDuplicacy.isSelected()) {
											String input = "";
											if (message.getMessageClass().equals("IPM.Task")) {
												input = mm.duplicacymapiTask(task);
											}
											input = input.replaceAll("\\s", "");
											input = input.trim();
											if (!listduplictask.contains(input)) {
												listduplictask.add(input);
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforexchange_output.createTask(taskuri, task);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													}
												} else {
													clientforexchange_output.createTask(taskuri, task);
													ConvertPSTOST_365.count_destination++;
													foldermessagecount++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													clientforexchange_output.createTask(taskuri, task);
													ConvertPSTOST_365.count_destination++;
													foldermessagecount++;
												}
											} else {
												clientforexchange_output.createTask(taskuri, task);
												ConvertPSTOST_365.count_destination++;
												foldermessagecount++;
											}
										}
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										if(e.getMessage().contains("ERROR")
												|| e.getMessage().contains("ERROR_CORRUPT_DATA")
												|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
												) {
											continue;
										}
										else if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Bad response")
												|| e.getMessage().contains(
														"An existing connection was forcibly closed by the remote host.")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
												|| e.getMessage().contains("Cannot access a disposed object.")
												|| e.getMessage().contains(
														"The support for the specified socket type does not exist in this address family.")) {
											mm.Progressbar.setVisible(false);
											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning("Exception : " + e.getMessage() + "Task" + " " + i
												+ mm.namingconventionmapi(message) + System.lineSeparator());
										continue;

									}

								} else {
									try {
										String Messageid = null;
										if (mm.chckbxRemoveDuplicacy.isSelected()) {
											String input = mm.duplicacymapi(message);
											if (!listduplicacy.contains(input)) {
												listduplicacy.add(input);

												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														Messageid = clientforexchange_output.appendMessage(subfolder,
																mess);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													}
												} else {
													Messageid = clientforexchange_output.appendMessage(subfolder, mess);
													ConvertPSTOST_365.count_destination++;
													foldermessagecount++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													Messageid = clientforexchange_output.appendMessage(subfolder, mess);
													ConvertPSTOST_365.count_destination++;
													foldermessagecount++;
												}
											} else {
												Messageid = clientforexchange_output.appendMessage(subfolder, mess);
												ConvertPSTOST_365.count_destination++;
												foldermessagecount++;
											}
										}
										if (Messageid != null) {
											if (((message.getFlags()
													& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
												clientforexchange_output.setReadFlag(Messageid, true);

											} else {
												clientforexchange_output.setReadFlag(Messageid, false);
											}
										}
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										System.out.println(e.getMessage());
										StringWriter sw = new StringWriter();
										e.printStackTrace(new PrintWriter(sw));
										String exceptionAsString = sw.toString();
										int number = message.getAttachments().size();

										if (exceptionAsString.contains("Message too large")
												|| exceptionAsString
														.contains("The message exceeds the maximum supported size.")
												|| number > 10) {
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
												//
												String Messageid = null;
												if (mm.chckbxRemoveDuplicacy.isSelected()) {
													String input = mm.duplicacymapi(message);
													if (!listduplicacy.contains(input)) {
														listduplicacy.add(input);

														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																Messageid = clientforexchange_output
																		.appendMessage(subfolder, mess);
																ConvertPSTOST_365.count_destination++;
																foldermessagecount++;
															}
														} else {
															Messageid = clientforexchange_output
																	.appendMessage(subfolder, mess);
															ConvertPSTOST_365.count_destination++;
															foldermessagecount++;
														}
													}
												} else {
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															Messageid = clientforexchange_output
																	.appendMessage(subfolder, mess);
															ConvertPSTOST_365.count_destination++;
															foldermessagecount++;
														}
													} else {
														Messageid = clientforexchange_output.appendMessage(subfolder,
																mess);
														ConvertPSTOST_365.count_destination++;
														foldermessagecount++;
													}
												}
												if (Messageid != null) {
													if (((message.getFlags()
															& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
														clientforexchange_output.setReadFlag(Messageid, true);

													} else {
														clientforexchange_output.setReadFlag(Messageid, false);
													}
												}
												//

											} catch (Exception e1) {
											}

										}
										else	if(e.getMessage().contains("ERROR")
												|| e.getMessage().contains("ERROR_CORRUPT_DATA")
												|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
												) {
											continue;
										}
										else if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Bad response")
												|| e.getMessage().contains(
														"An existing connection was forcibly closed by the remote host.")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
												|| e.getMessage().contains("Cannot access a disposed object.")
												|| e.getMessage().contains(
														"The support for the specified socket type does not exist in this address family.")) {
											mm.Progressbar.setVisible(false);
											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning("Exception : " + e.getMessage() + "Message " + " " + i
												+ mm.namingconventionmapi(message) + System.lineSeparator());

										continue;
									}

								}
								mm.lbl_progressreport
										.setText("  Total Message Saved Count  " + ConvertPSTOST_365.count_destination
												+ "  " + Folder + "   Extracting messsage " + message.getSubject());
								System.out.println("  Total Message Saved Count  " + ConvertPSTOST_365.count_destination
										+ "  " + Folder + "   Extracting messsage " + message.getSubject());
							} catch (Exception e) {
								if(e.getMessage().contains("ERROR")
										|| e.getMessage().contains("ERROR_CORRUPT_DATA")
										|| e.getMessage().contains("ERROR_ITEM_SAVE_PROPERTY")
										) {
									continue;
								}
								else if (e.getMessage().equalsIgnoreCase("The operation 'FetchMessage' terminated.")
										|| e.getMessage().contains(
												"No connection could be made because the target machine actively refused it.")
										|| e.getMessage().contains("ConnectFailure")
										|| e.getMessage().contains(
									
												"An existing connection was forcibly closed by the remote host.")
										|| e.getMessage().contains("Bad response")
										|| e.getMessage().contains("Operation failed")
										|| e.getMessage().contains("Rate limit hit")
										|| e.getMessage().contains("Operation has been canceled")
										|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
										|| e.getMessage().contains("Cannot access a disposed object.")
										|| e.getMessage().contains(
												"The support for the specified socket type does not exist in this address family.")) {
									mm.Progressbar.setVisible(false);
									i--;
								}
								connectionHandle(e.getMessage());
								continue;
							}

						}
					}
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderpstost_exchange(folderInfo, subfolder);
				}

				path = mm.removefolder(path);

			} catch (Exception e) {

				if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
						|| e.getMessage()
								.contains("No connection could be made because the target machine actively refused it.")
						|| e.getMessage().contains("ConnectFailure") || e.getMessage().contains("Operation failed")
						|| e.getMessage().contains("Bad response")
						|| e.getMessage().contains("Rate limit hit")
						|| e.getMessage().contains("An existing connection was forcibly closed by the remote host.")
						|| e.getMessage().contains("Operation has been canceled")
						|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
						|| e.getMessage().contains("Cannot access a disposed object.") || e.getMessage().contains(
								"The support for the specified socket type does not exist in this address family.")) {
					mm.Progressbar.setVisible(false);

				}
				connectionHandle(e.getMessage());
				continue;
			}

		}

	}

	String mailexchange(MailMessage message, IEWSClient clientforexchange_output, String Folderuri,
			MapiMessage message2) throws Exception {
		String Messageid = "";
		if (mm.chckbxRemoveDuplicacy.isSelected()) {

			String input = mm.duplicacymapi(message2);
			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						Messageid = clientforexchange_output.appendMessage(Folderuri, message);
						ConvertPSTOST_365.count_destination++;
					}
				} else {
					Messageid = clientforexchange_output.appendMessage(Folderuri, message);
					ConvertPSTOST_365.count_destination++;
				}
			}
		} else {
			if (main_multiplefile.datefilter.isSelected()) {
				if (datevalidflag) {
					Messageid = clientforexchange_output.appendMessage(Folderuri, message);
					ConvertPSTOST_365.count_destination++;
				}
			} else {
				Messageid = clientforexchange_output.appendMessage(Folderuri, message);
				ConvertPSTOST_365.count_destination++;
			}
		}
		return Messageid;
	}

	public void connectionHandle(String gotMessage) {
		mm.lbl_progressreport.setText("INTERNET Connection  LOST ");
		mm.label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));

		try {
			System.out.println("Connection Lost");
			mm.lbl_progressreport.setText("Connecting to Server Please Wait");
			if (filetype.equalsIgnoreCase("OFFICE 365")) {
				clientforexchange_output.dispose();
				
//				if(gotMessage.contains("Bad response")) {
//					ConnectionToOffice.refresh();
//					clientforexchange_output = conntiontooffice365_output(clientforexchange_output);
//				}else {
					clientforexchange_output = conntiontooffice365_output(clientforexchange_output);
//				}
				
				
			} else if (filetype.equalsIgnoreCase("Hotmail")) {
				clientforexchange_output.dispose();
				clientforexchange_output = conntiontohotmail_output(clientforexchange_output);
			}
			mm.label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));
			mm.lbl_progressreport.setText("Connection extablished Retriving Messasge");
		} catch (Exception e) {
			mm.lbl_progressreport.setText("INTERNET Connection  LOST ");
		}
		mm.Progressbar.setVisible(true);

	}

	@SuppressWarnings("deprecation")
	public IEWSClient conntiontooffice365_output(IEWSClient clientforexchange_output) throws Exception {
		while (true) {
			try {
				if (main_multiplefile.modern_Authentication.isSelected()) {
					clientforexchange_output=	ConnectionToOffice.conntiontooffice365_output1();
//					String token = Refresh_Token.refreshinput();
//					NetworkCredential credentials = new OAuthNetworkCredential(token);
//					EWSClient.useSAAJAPI(true);
//					clientforexchange_output = EWSClient.getEWSClient("https://outlook.office365.com/ews/exchange.asmx",
//							credentials);
//					clientforexchange_output.setTimeout(5 * 60 * 1000);
//					EmailClient.setSocketsLayerVersion2(true);
//					EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
				} else {
					clientforexchange_output=	main_multiplefile.conntiontooffice365_output();
//					clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3, password_p3);
//					clientforexchange_output.setTimeout(5 * 60 * 1000);
				}
				System.out.println("Connection Done : ");
				break;
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return clientforexchange_output;
	}

	// Hotmail

	@SuppressWarnings("deprecation")
	public IEWSClient conntiontohotmail_output(IEWSClient clientforexchange_output) throws Exception {
		while (true) {
			try {
				clientforexchange_output = EWSClient.getEWSClient("https://outlook.live.com/EWS/Exchange.asmx",
						username_p3, password_p3);
				clientforexchange_output.setTimeout(5 * 60 * 1000);
				EmailClient.setSocketsLayerVersion2(true);
				EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
				break;
			} catch (Exception e) {
			}
		}

		return clientforexchange_output;
	}

}
