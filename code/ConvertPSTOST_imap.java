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
import com.aspose.email.SecurityOptions;

public class ConvertPSTOST_imap implements Runnable {

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
	static Date fromdate;
	static Date todate;
	String path3 = "";
	ArrayList<Date> fromList = new ArrayList<Date>();
	ArrayList<Date> toList = new ArrayList<Date>();
	String from;
	String to;
	String sepreter = "";
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
	String domain_p3;
	int portnofiletype;
	String fname = "";

	public ConvertPSTOST_imap(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm, String username_p3, String password_p3, String Folderuri,
			ImapClient clientforimap_output, IConnection iconnforimap_output, String path, String domain_p3,
			int portnofiletype, String fa, String fname, String match) {

		this.mf = mf;
		this.mm = mm;
		this.filetype = filetype;
		this.filepath = filepath;
		this.destination_path = destination_path;
		ConvertPSTOST_imap.count_destination = count_destination;
		this.pstfolderlist = pstfolderlist;
		this.fromList = fromList;
		this.toList = toList;
		this.temppathm = temppathm;
		this.username_p3 = username_p3;
		this.password_p3 = password_p3;
		this.domain_p3 = domain_p3;
		this.Folderuri = Folderuri;
		this.clientforimap_output = clientforimap_output;
		this.iconnforimap_output = iconnforimap_output;
		this.path = path;
		this.fa = fa;
		this.fname = fname;
		this.match = match;
		this.portnofiletype = portnofiletype;
	}

	@Override
	public void run() {
		mm.label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));
		System.out.println(path + "  >>>path");
		match = path;
		System.out.println(match + "  >>> match ");
		sepreter = clientforimap_output.getDelimiter();
		path = path + sepreter + main_multiplefile.fname;

		clientforimap_output.createFolder(iconnforimap_output, path);
		clientforimap_output.selectFolder(iconnforimap_output, path);
		convertPSTOST_imap(mf, filetype, destination_path, count_destination, filepath, mm, pstfolderlist, fromList,
				toList, temppathm, Folderuri, clientforimap_output, iconnforimap_output, path, domain_p3,
				portnofiletype, match);
		main_multiplefile.count_destination = ConvertPSTOST_imap.count_destination;
		main_multiplefile.match = match;
		main_multiplefile.fa = fa;
		main_multiplefile.fname = fname;
		System.out.println(ConvertPSTOST_imap.count_destination);
	}

	private void convertPSTOST_imap(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm, String folderuri, ImapClient clientforimap_output,
			IConnection iconnforimap_output, String path, String domain_p3, int portnofiletype, String match) {
		System.out.println("Starting......");

		match = path;
		ConvertPSTOST_imap.count_destination = 0;
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
		String path1 = Folder;
		String sepreter = clientforimap_output.getDelimiter();
		path = path + sepreter + Folder;
		clientforimap_output.createFolder(iconnforimap_output, path);
		clientforimap_output.selectFolder(iconnforimap_output, path);
		listduplicacy.clear();
		listdupliccal.clear();
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
				if ((ConvertPSTOST_imap.count_destination % 1000) == 0) {
					if (s2) {
						mm.connectionHandle1();
					}
					s2 = true;
				}

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
				MailMessage mess = message.toMailMessage(options);
				if (message.getMessageClass().equals("IPM.Contact")) {
					MailMessage mapi = new MailMessage();
					try {
						MapiContact con = (MapiContact) message.toMapiMessageItem();
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

						con.save(temppathm + File.separator + countr + "contacttemp.vcf", ContactSaveFormat.VCard);
						File file = new File(temppathm + File.separator + countr + "contacttemp.vcf");
						mapi.addAttachment(new Attachment(temppathm + File.separator + countr + "contacttemp.vcf"));
						file.delete();

						if (mm.chckbxRemoveDuplicacy.isSelected()) {

							String input = mm.duplicacymapiContact(con);
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listdupliccontact.contains(input)) {
								listdupliccontact.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
										ConvertPSTOST_imap.count_destination++;
									}
								} else {
									clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
									ConvertPSTOST_imap.count_destination++;
								}

							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
									ConvertPSTOST_imap.count_destination++;
								}
							} else {
								clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
								ConvertPSTOST_imap.count_destination++;
							}
						}

						countr++;
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mm.namingconventionmapi(message));
					} catch (Exception e) {
						if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("Operation has been canceled")
								|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
								|| e.getMessage().contains("Software caused connection abort: recv failed")
								|| e.getMessage().contains("Timeout")
								|| e.getMessage().contains("Network is unreachable: connect")) {
							mm.Progressbar.setVisible(false);
							i--;
						}
						connectionHandle(e.getMessage());
						mf.logger.warning(e.getMessage() + "Contact" + " " + countr + System.lineSeparator());
						continue;
					}

				} else if (message.getMessageClass().equals("IPM.Appointment")
						|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
					MailMessage mapi = new MailMessage();
					try {

						MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

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
										clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
										ConvertPSTOST_imap.count_destination++;
									}
								} else {
									clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
									ConvertPSTOST_imap.count_destination++;
								}

							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
									ConvertPSTOST_imap.count_destination++;
								}
							} else {
								clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
								ConvertPSTOST_imap.count_destination++;
							}
						}
						countr++;
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mm.namingconventionmapi(message));
					} catch (Exception e) {
						if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("Operation has been canceled")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
								|| e.getMessage().contains("Software caused connection abort: recv failed")
								|| e.getMessage().contains("Timeout")
								|| e.getMessage().contains("Network is unreachable: connect")) {
							mm.Progressbar.setVisible(false);
							i--;
						}
						connectionHandle(e.getMessage());
						mf.logger.warning(e.getMessage() + "Calendar" + " " + countr + System.lineSeparator());
						continue;
					}
				} else if (message.getMessageClass().equals("IPM.Task")) {
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
										clientforimap_output.appendMessage(iconnforimap_output, path, mess);
										ConvertPSTOST_imap.count_destination++;
									}
								} else {
									clientforimap_output.appendMessage(iconnforimap_output, path, mess);
									ConvertPSTOST_imap.count_destination++;
								}

							}
						} else {

							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforimap_output.appendMessage(iconnforimap_output, path, mess);
									ConvertPSTOST_imap.count_destination++;
								}
							} else {
								clientforimap_output.appendMessage(iconnforimap_output, path, mess);
								ConvertPSTOST_imap.count_destination++;
							}
						}
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mm.namingconventionmapi(message));
					} catch (Exception e) {
						if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("Operation has been canceled")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
								|| e.getMessage().contains("Software caused connection abort: recv failed")
								|| e.getMessage().contains("Timeout")
								|| e.getMessage().contains("Network is unreachable: connect")) {
							mm.Progressbar.setVisible(false);
							i--;
						}
						connectionHandle(e.getMessage());
						mf.logger.warning(e.getMessage() + "Task" + " " + countr + System.lineSeparator());
						continue;
					} finally {
						filetype = s;
					}
				} else {
					try {
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
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mm.namingconventionmapi(message));
					} catch (Exception e) {
						if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("Operation has been canceled")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
								|| e.getMessage().contains("Software caused connection abort: recv failed")
								|| e.getMessage().contains("Timeout")
								|| e.getMessage().contains("Network is unreachable: connect")) {
							mm.Progressbar.setVisible(false);
							i--;
						}
						connectionHandle(e.getMessage());
						mf.logger.warning(
								e.getMessage() + "Message" + " " + message.getDeliveryTime() + System.lineSeparator());
						continue;
					}

				}
				mm.lbl_progressreport.setText("Total message Saved Count " + ConvertPSTOST_imap.count_destination + "  "
						+ Folder + " Extracting messsage " + message.getSubject());
			} catch (Exception e) {
				if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
						|| e.getMessage().contains("ConnectFailure") || e.getMessage().contains("Rate limit hit")
						|| e.getMessage().contains("Operation failed")
						|| e.getMessage().contains("Operation has been canceled")
						|| e.getMessage().contains("Operation failed")
						|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
						|| e.getMessage().contains("Software caused connection abort: recv failed")
						|| e.getMessage().contains("Timeout")
						|| e.getMessage().contains("Network is unreachable: connect")) {
					mm.Progressbar.setVisible(false);
					i--;
				}
				connectionHandle(e.getMessage());
				mf.logger.warning(e.getMessage() + System.lineSeparator());

				continue;

			}

		}

		FolderInfoCollection folderInf = pst.getRootFolder().getSubFolders();

		for (int j = 0; j < folderInf.size(); j++) {
			try {

				if (mm.stop) {
					break;
				}
				FolderInfo folderInfo = folderInf.get_Item(j);

				Folder = folderInfo.getDisplayName();
				Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();
				Folder = path1 + File.separator + Folder;

				String sfolder = Folder;
				for (int l = 0; l < pstfolderlist.size(); l++) {

					if (mm.stop) {
						break;
					}
					if (pstfolderlist.get(l).equalsIgnoreCase(Folder)) {
						sepreter = clientforimap_output.getDelimiter();
						path = path + sepreter + folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");

						mm.lbl_progressreport.setText(" Getting Folder " + Folder);
						if (clientforimap_output.existFolder(path)) {
							clientforimap_output.selectFolder(path);
						} else {
							clientforimap_output.createFolder(iconnforimap_output, path);
							clientforimap_output.selectFolder(iconnforimap_output, path);
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
						boolean s22 = false;
						for (int i = 0; i < messagesize; i++) {
							try {

								if (mm.stop) {
									break;
								}
								if ((i % 100) == 0) {
									System.gc();
								}
								if ((ConvertPSTOST_imap.count_destination % 1000) == 0) {
									if (s22) {
										mm.connectionHandle1();
									}
									s22 = true;
								}
								MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);
								MapiMessage message1 = pst.extractMessage(messageInfo);
								MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
								MailConversionOptions de = new MailConversionOptions();
								MailMessage mess = message1.toMailMessage(de);
								if (mm.chckbxMigrateOrBackup.isSelected()) {
									mess.getAttachments().clear();
								}
								MapiMessage message = MapiMessage.fromMailMessage(mess, d);

								Date Receiveddate = message.getDeliveryTime();
								if (main_multiplefile.datefilter.isSelected()) {
									datevalidflag = mm.checkdate(message, mess);
								}
								if (message.getMessageClass().equals("IPM.Contact")) {
									MailMessage mapi = new MailMessage();
									try {
										MapiContact con = (MapiContact) message.toMapiMessageItem();
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
										con.save(temppathm + File.separator + countr + "contacttemp.vcf",
												ContactSaveFormat.VCard);
										File file = new File(temppathm + File.separator + countr + "contacttemp.vcf");
										mapi.addAttachment(new Attachment(
												temppathm + File.separator + countr + "contacttemp.vcf"));
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
														ConvertPSTOST_imap.count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													ConvertPSTOST_imap.count_destination++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													ConvertPSTOST_imap.count_destination++;
												}
											} else {
												clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
												ConvertPSTOST_imap.count_destination++;
											}
										}
										countr++;
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
												|| e.getMessage()
														.contains("Software caused connection abort: recv failed")
												|| e.getMessage().contains("Timeout")
												|| e.getMessage().contains("Network is unreachable: connect")) {
											mm.Progressbar.setVisible(false);
											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning(
												e.getMessage() + "Contact" + " " + countr + System.lineSeparator());
										continue;
									}
								} else if (message.getMessageClass().equals("IPM.Appointment")
										|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
									MailMessage mapi = new MailMessage();
									try {

										MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

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
										cal.save(temppathm + File.separator + countr + "caltemp.ics",
												AppointmentSaveFormat.Ics);
										File file = new File(temppathm + File.separator + countr + "caltemp.ics");

										mapi.addAttachment(
												new Attachment(temppathm + File.separator + countr + "caltemp.ics"));
										file.delete();
										Receiveddate = cal.getStartDate();
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
														ConvertPSTOST_imap.count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													ConvertPSTOST_imap.count_destination++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													ConvertPSTOST_imap.count_destination++;
												}
											} else {
												clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
												ConvertPSTOST_imap.count_destination++;
											}
										}
										countr++;
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
												|| e.getMessage()
														.contains("Software caused connection abort: recv failed")
												|| e.getMessage().contains("Timeout")
												|| e.getMessage().contains("Network is unreachable: connect")) {
											mm.Progressbar.setVisible(false);
											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning(
												e.getMessage() + "Calendar" + " " + countr + System.lineSeparator());
										continue;
									}

								} else if (message.getMessageClass().equals("IPM.Task")) {

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
														ConvertPSTOST_imap.count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path, mess);
													ConvertPSTOST_imap.count_destination++;
												}
											}
										} else {

											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													clientforimap_output.appendMessage(iconnforimap_output, path, mess);
													ConvertPSTOST_imap.count_destination++;
												}
											} else {
												clientforimap_output.appendMessage(iconnforimap_output, path, mess);
												ConvertPSTOST_imap.count_destination++;
											}
										}
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {

										if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
												|| e.getMessage()
														.contains("Software caused connection abort: recv failed")
												|| e.getMessage().contains("Timeout")
												|| e.getMessage().contains("Network is unreachable: connect")) {
											mm.Progressbar.setVisible(false);

											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning(
												e.getMessage() + "Task" + " " + countr + System.lineSeparator());
										continue;
									} finally {
										filetype = s;
									}

								} else {
									try {

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
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mm.namingconventionmapi(message));
									} catch (Exception e) {
										if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
												|| e.getMessage()
														.contains("Software caused connection abort: recv failed")
												|| e.getMessage().contains("Timeout")
												|| e.getMessage().contains("Network is unreachable: connect")) {
											mm.Progressbar.setVisible(false);
											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning(e.getMessage() + "Message" + " " + message.getDeliveryTime()
												+ System.lineSeparator());
										continue;
									}
								}
								mm.lbl_progressreport
										.setText("Total message Saved Count " + ConvertPSTOST_imap.count_destination
												+ "  " + Folder + " Extracting messsage " + message.getSubject());
							} catch (Exception e) {
								if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
										|| e.getMessage().contains("ConnectFailure")
										|| e.getMessage().contains("Operation failed")
										|| e.getMessage().contains("Rate limit hit")
										|| e.getMessage().contains("Operation has been canceled")
										|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
										|| e.getMessage().contains("Software caused connection abort: recv failed")
										|| e.getMessage().contains("Timeout")
										|| e.getMessage().contains("Network is unreachable: connect")) {
									mm.Progressbar.setVisible(false);
//									i--;
								}
								connectionHandle(e.getMessage());
								mf.logger.warning(e.getMessage() + System.lineSeparator());
								continue;

							}

						}
					}
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderforpstost_imap(folderInfo, sfolder);
				}
				path = path.replace("." + folderInfo.getDisplayName(), "");
			} catch (Exception e) {
				if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
						|| e.getMessage().contains("ConnectFailure") || e.getMessage().contains("Operation failed")
						|| e.getMessage().contains("Rate limit hit")
						|| e.getMessage().contains("Operation has been canceled")
						|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
						|| e.getMessage().contains("Software caused connection abort: recv failed")
						|| e.getMessage().contains("Timeout")
						|| e.getMessage().contains("Network is unreachable: connect")) {
					mm.Progressbar.setVisible(false);
				}
				connectionHandle(e.getMessage());
				mf.logger.warning(e.getMessage() + System.lineSeparator());

				continue;

			}

		}

	}

	private void getsubfolderforpstost_imap(FolderInfo f, String sfolder) {

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
				sfolder = sfolder + File.separator + Folder;
				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}
					if (pstfolderlist.get(l).equalsIgnoreCase(sfolder)) {

						mm.lbl_progressreport.setText("Getting Folder " + Folder);
						sepreter = clientforimap_output.getDelimiter();
						System.out.println(sepreter + "  sepreter >>");
						String path1 = path + sepreter + sfolder.replace(File.separator, sepreter);
						System.out.println(path);
						if (clientforimap_output.existFolder(path1)) {
							clientforimap_output.selectFolder(path1);
						} else {
							clientforimap_output.createFolder(iconnforimap_output, path1);
							clientforimap_output.selectFolder(iconnforimap_output, path1);
						}
						mf.listduplicacy.clear();
						int countr = 1;
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
									if ((ConvertPSTOST_imap.count_destination % 1000) == 0) {
										if (s22) {
											mm.connectionHandle1();
										}
										s22 = true;
									}

									MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

									MapiMessage message1 = pst.extractMessage(messageInfo);
									MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
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
										MailMessage mapi = new MailMessage();
										try {
											MapiContact con = (MapiContact) message.toMapiMessageItem();
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
											con.save(temppathm + File.separator + mm.namingconventionmapi(message)
													+ ".vcf", ContactSaveFormat.VCard);
											File file = new File(temppathm + File.separator
													+ mm.namingconventionmapi(message) + ".vcf");
											mapi.addAttachment(new Attachment(temppathm + File.separator
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
																	path1, mapi);
															ConvertPSTOST_imap.count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, path1,
																mapi);
														ConvertPSTOST_imap.count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, path1,
																mapi);
														ConvertPSTOST_imap.count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path1,
															mapi);
													ConvertPSTOST_imap.count_destination++;
												}
											}
											countr++;
										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mm.namingconventionmapi(message));
										} catch (Exception e) {
											if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
													|| e.getMessage().contains("ConnectFailure")
													|| e.getMessage().contains("Operation failed")
													|| e.getMessage().contains("Operation has been canceled")
													|| e.getMessage()
															.contains("The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Timeout")
													|| e.getMessage().contains("Network is unreachable: connect")) {
												mm.Progressbar.setVisible(false);

												i--;
											}
											connectionHandle(e.getMessage());
											mf.logger.warning(
													e.getMessage() + "Contact" + " " + countr + System.lineSeparator());
											continue;
										}

									} else if (message.getMessageClass().equals("IPM.Appointment")
											|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
										MailMessage mapi = new MailMessage();
										try {

											MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

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
											cal.save(temppathm + File.separator + mm.namingconventionmapi(message)
													+ ".ics", AppointmentSaveFormat.Ics);
											File file = new File(temppathm + File.separator
													+ mm.namingconventionmapi(message) + ".ics");

											mapi.addAttachment(new Attachment(temppathm + File.separator
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
																	path1, mapi);
															ConvertPSTOST_imap.count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, path1,
																mapi);
														ConvertPSTOST_imap.count_destination++;
													}

												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, path1,
																mapi);
														ConvertPSTOST_imap.count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path1,
															mapi);
													ConvertPSTOST_imap.count_destination++;
												}
											}
											countr++;
										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mm.namingconventionmapi(message));
										} catch (Exception e) {
											if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
													|| e.getMessage().contains("ConnectFailure")
													|| e.getMessage().contains("Operation failed.")
													|| e.getMessage().contains("Operation has been canceled")
													|| e.getMessage()
															.contains("The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Timeout")
													|| e.getMessage().contains("Network is unreachable: connect")) {
												mm.Progressbar.setVisible(false);

												i--;
											}
											connectionHandle(e.getMessage());
											mf.logger.warning(e.getMessage() + "Calendar" + " " + countr
													+ System.lineSeparator());
											continue;
										}

									} else if (message.getMessageClass().equals("IPM.Task")) {

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
															clientforimap_output.appendMessage(iconnforimap_output,
																	path1, mess);
															ConvertPSTOST_imap.count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, path1,
																mess);
														ConvertPSTOST_imap.count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, path1,
																mess);
														ConvertPSTOST_imap.count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path1,
															mess);
													ConvertPSTOST_imap.count_destination++;
												}
											}

										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mm.namingconventionmapi(message));
										} catch (Exception e) {
											if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
													|| e.getMessage().contains("ConnectFailure")
													|| e.getMessage().contains("Operation failed")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage().contains("Operation has been canceled")
													|| e.getMessage()
															.contains("The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Timeout")
													|| e.getMessage().contains("Network is unreachable: connect")) {
												mm.Progressbar.setVisible(false);
												i--;
											}
											connectionHandle(e.getMessage());
											mf.logger.warning(
													e.getMessage() + "Task" + " " + countr + System.lineSeparator());
											continue;
										} finally {
											filetype = s;
										}

									} else {
										try {
											String messageid = mailimap(mess, path1);
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
										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mm.namingconventionmapi(message));
										} catch (Exception e) {

											if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
													|| e.getMessage().contains("ConnectFailure")
													|| e.getMessage().contains("Operation failed")
													|| e.getMessage().contains("Rate limit hit")
													|| e.getMessage().contains("Operation has been canceled")
													|| e.getMessage()
															.contains("The operation 'AppendMessage' terminated.")
													|| e.getMessage()
															.contains("Software caused connection abort: recv failed")
													|| e.getMessage().contains("Timeout")
													|| e.getMessage().contains("Network is unreachable: connect")) {
												mm.Progressbar.setVisible(false);

												i--;
											}
											connectionHandle(e.getMessage());
											mf.logger.warning(e.getMessage() + "Message" + " "
													+ message.getDeliveryTime() + System.lineSeparator());
											continue;
										}

									}
									mm.lbl_progressreport
											.setText("Total message Saved Count " + ConvertPSTOST_imap.count_destination
													+ "  " + Folder + " Extracting messsage " + message.getSubject());

								} catch (Exception e) {
									if (e.getMessage().equalsIgnoreCase("The operation 'FetchMessage' terminated.")
											|| e.getMessage().contains("ConnectFailure")
											|| e.getMessage().contains("Operation failed")
											|| e.getMessage().contains("Rate limit hit")
											|| e.getMessage().contains("Operation has been canceled")
											|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
											|| e.getMessage().contains("Software caused connection abort: recv failed")
											|| e.getMessage().contains("Timeout")
											|| e.getMessage().contains("Network is unreachable: connect")) {
										mm.Progressbar.setVisible(false);

//										i--;
									}
									connectionHandle(e.getMessage());
									mf.logger.warning(e.getMessage() + System.lineSeparator());

									continue;
								}
							}
						}
					}
				}
				if (folderf.hasSubFolders()) {

					getsubfolderforpstost_imap(folderf, sfolder);
				}

				path = path.replace("." + Folder, "");
				sfolder = mm.removefolder(sfolder);
			} catch (Exception e) {

				if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
						|| e.getMessage().contains("Operation failed") || e.getMessage().contains("Rate limit hit")
						|| e.getMessage().contains("Operation has been canceled")
						|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
						|| e.getMessage().contains("Software caused connection abort: recv failed")
						|| e.getMessage().contains("Timeout")
						|| e.getMessage().contains("Network is unreachable: connect")) {
					mm.Progressbar.setVisible(false);

				}
				connectionHandle(e.getMessage());
			}

		}

	}

	void ExceptionHandler(Exception e) {

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
							ConvertPSTOST_imap.count_destination++;
						}
					} else {
						Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
						foldermessagecount++;
						ConvertPSTOST_imap.count_destination++;
					}
				}
			} else {
				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
						foldermessagecount++;
						ConvertPSTOST_imap.count_destination++;
					}
				} else {
					Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
					ConvertPSTOST_imap.count_destination++;

				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			StringWriter sw = new StringWriter();
			e.printStackTrace(new PrintWriter(sw));//
			String exceptionAsString = sw.toString();
			if (exceptionAsString.contains("Message too large") || exceptionAsString.contains("TOOBIG")) {
				File f = new File((System.getProperty("user.home") + File.separator + "Desktop") + File.separator
						+ mm.calendertime + File.separator + "Attachment" + File.separator
						+ mf.namingconventionmail(message));
				f.mkdirs();
				mf.logger.info("Message size was greater than allowed size so attachment has been deleted and saved in "
						+ f.getAbsolutePath());

				MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
				MapiMessage message1 = MapiMessage.fromMailMessage(message, d);

				for (MapiAttachment attachment : message1.getAttachments()) {

					attachment.save(f.getAbsolutePath() + File.separator
							+ main_multiplefile.getRidOfIllegalFileNameCharacters(attachment.getLongFileName()));

				}

				message1.getAttachments().clear();
				message1.getAttachments().removeAll(message1.getAttachments());

				MailConversionOptions d1 = new MailConversionOptions();
				message = message1.toMailMessage(d1);

				Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
				ConvertPSTOST_imap.count_destination++;
			}
		}
		return Messageid;
	}

	public void connectionHandle(String gotMessage) {
		mm.lbl_progressreport.setText("INTERNET Connection  LOST ");

		mm.label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));

		while (true) {
			try {
				mm.lbl_progressreport.setText("Connecting to Server Please Wait");
				if (filetype.equalsIgnoreCase("IMAP")) {
					connectiontoimap_output();
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

	@SuppressWarnings("deprecation")
	public ImapClient connectiontoimap_output() throws Exception {
		clientforimap_output = new ImapClient(mm.domain_p3, mm.portnofiletype, mm.username_p3, mm.password_p3);
		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);
		clientforimap_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}
}
