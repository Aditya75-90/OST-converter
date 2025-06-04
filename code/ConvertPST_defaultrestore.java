package email.code;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.aspose.email.Appointment;
import com.aspose.email.AppointmentLoadOptions;
import com.aspose.email.AppointmentSaveFormat;
import com.aspose.email.BodyContentType;
import com.aspose.email.ContactSaveFormat;
import com.aspose.email.ExchangeFolderInfo;
import com.aspose.email.FileFormatVersion;
import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.IConnection;
import com.aspose.email.IEWSClient;
import com.aspose.email.ImapClient;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiMessageFlags;
import com.aspose.email.MapiTask;
import com.aspose.email.MessageInfo;
import com.aspose.email.MessageInfoCollection;
import com.aspose.email.PersonalStorage;

public class ConvertPST_defaultrestore implements Runnable {
	PersonalStorage ost;
	int splitcount = 0;
	String splitpath = "";
	private List<String> listdupliccal = new ArrayList<String>();
	private List<String> listduplictask = new ArrayList<String>();
	private List<String> listdupliccontact = new ArrayList<String>();
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
	private String path = "";
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
	static IEWSClient clientforexchange_output;
	static ImapClient clientforimap_output;
	static IConnection iconnforimap_output;
	String path1 = "";
	long maxsize = 0;
	int pstindex = 0;

	public ConvertPST_defaultrestore(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm) {
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
	}

	@Override
	public void run() {
		convertPST_defaultrestore(mf, filetype, destination_path, count_destination, filepath, mm, pstfolderlist,
				fromList, toList, temppathm);
		main_multiplefile.count_destination = count_destination;
	}

	@SuppressWarnings("deprecation")
	private void convertPST_defaultrestore(Main_Frame mf, String filetype, String destination_path,
			long count_destination, String filepath, main_multiplefile mm, List<String> pstfolderlist,
			ArrayList<Date> fromList, ArrayList<Date> toList, String temppathm) {

		String path2 = "";
		FolderInfo folderInfo2 = pst.getRootFolder();

		Folder = folderInfo2.getDisplayName();
		Folder = Folder.replace(",", "").replace(".", "");
		Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
		Folder = Folder.replaceAll("[\\[\\]]", "");
		Folder = Folder.trim();
		if (Folder.equalsIgnoreCase("")) {
			Folder = "Root Folder";
		}
		path = path + File.separator + Folder;
		path2 = Folder;

		String defaultfolder = "";
		if (mm.chckbxCustomFolderName.isSelected()) {
			defaultfolder = clientforexchange_output.createFolder(mm.textField_customfolder.getText() + mm.calendertime)
					.getUri();
		} else {
			defaultfolder = clientforexchange_output.createFolder(mm.fname + mm.calendertime).getUri();
		}

		if (folderInfo2.getContentCount() > 0) {
			String mailbox = clientforexchange_output.createFolder(defaultfolder, Folder).getUri();

			MessageInfoCollection messageInfoCollection1 = folderInfo2.getContents();
			messageaddpst(messageInfoCollection1, mailbox, Folder);
		}

		FolderInfoCollection folderInf = pst.getRootFolder().getSubFolders();

		for (int j = 0; j < folderInf.size(); j++) {
			try {

				FolderInfo folderInfo = folderInf.get_Item(j);
				String Folder = folderInfo.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();

				if (mm.stop) {
					break;
				}

				path = path + File.separator + Folder;
				String path3 = Folder;
				path3 = path3.replaceAll("[\\[\\]]", "");
				Folder = path3;
				String mailfolder = "";
				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}
					String path1 = pstfolderlist.get(l).replace(path2 + File.separator, "");

					if (path1.equalsIgnoreCase(path3)) {
						mm.lbl_progressreport.setText(" Getting Folder " + Folder);

						listdupliccal.clear();
						mf.listduplicacy.clear();
						listdupliccontact.clear();
						listduplictask.clear();

						if (Folder.contains("Inbox")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getInboxUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();

								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Deleted Item")) {

							if (folderInfo.getContentCount() > 0) {

								mailfolder = clientforexchange_output.getMailboxInfo().getDeletedItemsUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Calendar")) {

							if (folderInfo.getContentCount() > 0) {

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
								ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
								if (!clientforexchange_output.folderExists(
										clientforexchange_output.getMailboxInfo().getCalendarUri(),
										mm.calendertime + "/" + Folder, subfolderInfo)) {
									mailfolder = clientforexchange_output
											.createFolder(clientforexchange_output.getMailboxInfo().getCalendarUri(),
													mm.calendertime + "/" + Folder, null, "IPF.Appointment")
											.getUri();
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
										de.setConvertAsTnef(true);
										MailMessage mess = message1.toMailMessage(de);
										MapiMessage message = MapiMessage.fromMailMessage(mess, d);
										if (main_multiplefile.datefilter.isSelected()) {
											datevalidflag = checkdate(message);
											System.out.println(datevalidflag);
										}
										if (message.getMessageClass().equals("IPM.Appointment")
												|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

											try {

												MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();
												cal.save(temppathm + File.separator + mf.namingconventionmapi(message)
														+ ".ics", AppointmentSaveFormat.Ics);
												File file = new File(temppathm + File.separator
														+ mf.namingconventionmapi(message) + ".ics");
												AppointmentLoadOptions optiona = new AppointmentLoadOptions();
												optiona.getIgnoreSmtpAddressCheck();
												Appointment calDoc = Appointment.load(temppathm + File.separator
														+ mf.namingconventionmapi(message) + ".ics", optiona);

												if (mm.chckbxRemoveDuplicacy.isSelected()) {

													String input = mm.duplicacymapiCal(cal);
													input = input.replaceAll("\\s", "");
													input = input.trim();

													if (!listdupliccal.contains(input)) {
														System.out.println("Not a duplicate message");
														listdupliccal.add(input);

														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																clientforexchange_output.createAppointment(calDoc,
																		mailfolder);

																count_destination++;
															}
														} else {
															clientforexchange_output.createAppointment(calDoc,
																	mailfolder);

															count_destination++;
														}
													}
												} else {
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforexchange_output.createAppointment(calDoc,
																	mailfolder);
															count_destination++;
														}
													} else {
														clientforexchange_output.createAppointment(calDoc, mailfolder);
														count_destination++;
													}

												}
												file.delete();
											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception : " + e.getMessage() + "  "
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										}

										mm.lbl_progressreport
												.setText("  Total Message Saved Count  " + count_destination + "  "
														+ Folder + "   Extracting messsage " + message.getSubject());

									} catch (Exception e) {
										continue;
									}

								}

							}
						} else if (Folder.contains("Tasks") || Folder.contains("ToDo")) {

							if (folderInfo.getContentCount() > 0) {

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
										de.setConvertAsTnef(true);
										MailMessage mess = message1.toMailMessage(de);
										MapiMessage message = MapiMessage.fromMailMessage(mess, d);
										if (main_multiplefile.datefilter.isSelected()) {
											datevalidflag = checkdate(message);
										}
										if (messageInfo.getMessageClass().equals("IPM.StickyNote")
												|| messageInfo.getMessageClass().equals("IPM.Task")) {
											try {

												MapiTask task = null;
												if (messageInfo.getMessageClass().equals("IPM.Task")) {
													task = (MapiTask) message.toMapiMessageItem();
												}
												mailfolder = clientforexchange_output.getMailboxInfo().getTasksUri();
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
																clientforexchange_output.appendMessage(mailfolder,
																		mess);
																count_destination++;
															}
														} else {
															clientforexchange_output.appendMessage(mailfolder, mess);
															count_destination++;
														}

													}
												} else {

													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforexchange_output.appendMessage(mailfolder, mess);
															count_destination++;
														}
													} else {
														clientforexchange_output.appendMessage(mailfolder, mess);
														count_destination++;

													}
												}
											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception : " + e.getMessage() + "Task" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										}

										mm.lbl_progressreport
												.setText("  Total Message Saved Count  " + count_destination + "  "
														+ Folder + "   Extracting messsage " + message.getSubject());

									} catch (Exception e) {
										continue;
									}

								}
							}
						} else if (Folder.contains("Contacts")) {

							if (folderInfo.getContentCount() > 0) {

								ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
								if (!clientforexchange_output.folderExists(
										clientforexchange_output.getMailboxInfo().getContactsUri(),
										mm.calendertime + "/" + Folder, subfolderInfo)) {
									mailfolder = clientforexchange_output
											.createFolder(clientforexchange_output.getMailboxInfo().getContactsUri(),
													mm.calendertime + "/" + Folder, null, "IPF.Contact")
											.getUri();
								}

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
										de.setConvertAsTnef(true);
										MailMessage mess = message1.toMailMessage(de);
										MapiMessage message = MapiMessage.fromMailMessage(mess, d);
										if (main_multiplefile.datefilter.isSelected()) {
											datevalidflag = checkdate(message);
										}
										if (message.getMessageClass().equals("IPM.Contact")) {
											try {

												MapiContact con = (MapiContact) message.toMapiMessageItem();
												ByteArrayOutputStream bos = new ByteArrayOutputStream();
												con.save(bos, ContactSaveFormat.VCard);
												ByteArrayInputStream inStream = new ByteArrayInputStream(
														bos.toByteArray());
												MapiContact mapicontact = MapiContact.fromVCard(inStream);
												if (mm.chckbxRemoveDuplicacy.isSelected()) {

													String input = mm.duplicacymapiContact(con);

													if (!listdupliccontact.contains(input)) {
														System.out.println("Not a duplicate message");
														listdupliccontact.add(input);

														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																clientforexchange_output.createContact(mailfolder,
																		mapicontact);
																count_destination++;
															}
														} else {
															clientforexchange_output.createContact(mailfolder,
																	mapicontact);
															count_destination++;
														}

													}
												} else {
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforexchange_output.createContact(mailfolder,
																	mapicontact);
															count_destination++;
														}
													} else {
														clientforexchange_output.createContact(mailfolder, mapicontact);
														count_destination++;
													}
												}
											} catch (Error e) {
												mf.logger.warning("ERROR : " + e.getMessage() + System.lineSeparator());
											} catch (Exception e) {
												mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										}

										mm.lbl_progressreport
												.setText("  Total Message Saved Count  " + count_destination + "  "
														+ Folder + "   Extracting messsage " + message.getSubject());

									} catch (Exception e) {
										continue;
									}

								}
							}
						} else if (Folder.contains("Outbox")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getOutboxUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Draft")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getDraftsUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Junk E-Mail")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getJunkeMailsUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Notes")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getNotesUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Journal")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getJournalUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Sent")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getSentItemsUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else {
							mailfolder = clientforexchange_output.createFolder(defaultfolder, Folder).getUri();
							if (folderInfo.getContentCount() > 0) {
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();

								mf.listduplicacy.clear();
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
								String scm = "";

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
										MapiMessage message = MapiMessage.fromMailMessage(mess, d);
										if (main_multiplefile.datefilter.isSelected()) {
											datevalidflag = checkdate(message);
										}
										if (message.getMessageClass().equals("IPM.Contact")) {
											try {
												ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };

												if (!clientforexchange_output.folderExists(
														clientforexchange_output.getMailboxInfo().getContactsUri(),
														mm.calendertime + "/" + Folder, subfolderInfo)) {
													scm = clientforexchange_output
															.createFolder(
																	clientforexchange_output.getMailboxInfo()
																			.getContactsUri(),
																	mm.calendertime + "/" + Folder, null, "IPF.Contact")
															.getUri();
												}

												MapiContact con = (MapiContact) message.toMapiMessageItem();

												if (mm.chckbxRemoveDuplicacy.isSelected()) {

													String input = mm.duplicacymapiContact(con);

													if (!listdupliccontact.contains(input)) {
														System.out.println("Not a duplicate message");
														listdupliccontact.add(input);
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																clientforexchange_output.createContact(scm, con);
																count_destination++;
															}
														} else {
															clientforexchange_output.createContact(scm, con);
															count_destination++;
														}
													}
												} else {
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforexchange_output.createContact(scm, con);
															count_destination++;
														}
													} else {
														clientforexchange_output.createContact(scm, con);
														count_destination++;
													}
												}
											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												continue;
											}

										} else if (message.getMessageClass().equals("IPM.Appointment")
												|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

											try {

												MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

												ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
												if (!clientforexchange_output.folderExists(
														clientforexchange_output.getMailboxInfo().getCalendarUri(),
														mm.calendertime + "/" + Folder, subfolderInfo)) {
													scm = clientforexchange_output.createFolder(
															clientforexchange_output.getMailboxInfo().getCalendarUri(),
															mm.calendertime + "/" + Folder, null, "IPF.Appointment")
															.getUri();
												}

												cal.save(temppathm + File.separator + mf.namingconventionmapi(message)
														+ ".ics", AppointmentSaveFormat.Ics);
												File file = new File(temppathm + File.separator
														+ mf.namingconventionmapi(message) + ".ics");
												AppointmentLoadOptions optiona = new AppointmentLoadOptions();
												optiona.getIgnoreSmtpAddressCheck();
												Appointment calDoc = Appointment.load(temppathm + File.separator
														+ mf.namingconventionmapi(message) + ".ics", optiona);

												if (mm.chckbxRemoveDuplicacy.isSelected()) {
													String input = mm.duplicacymapiCal(cal);
													input = input.replaceAll("\\s", "");
													input = input.trim();
													if (!listdupliccal.contains(input)) {
														listdupliccal.add(input);
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																clientforexchange_output.createAppointment(calDoc, scm);
																count_destination++;
															}
														} else {
															clientforexchange_output.createAppointment(calDoc, scm);
															count_destination++;
														}
													}
												} else {
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforexchange_output.createAppointment(calDoc, scm);
															count_destination++;
														}
													} else {
														clientforexchange_output.createAppointment(calDoc, scm);
														count_destination++;
													}
												}

												file.delete();
											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										} else if (message.getMessageClass().equals("IPM.StickyNote")
												|| message.getMessageClass().equals("IPM.Task")) {
											try {
												MapiTask task = (MapiTask) message.toMapiMessageItem();
												MailConversionOptions options = new MailConversionOptions();
												options.setConvertAsTnef(true);
												String taskuri = clientforexchange_output.getMailboxInfo()
														.getTasksUri();
												if (mm.chckbxRemoveDuplicacy.isSelected()) {
													String input = mm.duplicacymapiTask(task);
													input = input.replaceAll("\\s", "");
													input = input.trim();
													if (!listdupliccal.contains(input)) {
														listdupliccal.add(input);
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																clientforexchange_output.createTask(taskuri, task);
																count_destination++;
															}
														} else {
															clientforexchange_output.createTask(taskuri, task);
															count_destination++;
														}
													}
												} else {
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforexchange_output.createTask(taskuri, task);
															count_destination++;
														}

													} else {
														clientforexchange_output.createTask(taskuri, task);
														count_destination++;
													}
												}
											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception : " + e.getMessage() + "Task" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										} else {

											try {
												String messageid = mailexchange(mess, clientforexchange_output,
														mailfolder);
												if (!messageid.equalsIgnoreCase("")) {
													if (((message.getFlags()
															& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
														clientforexchange_output.setReadFlag(messageid, true);

													} else {
														clientforexchange_output.setReadFlag(messageid, false);
													}
												}

											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										}

										mm.lbl_progressreport
												.setText("  Total Message Saved Count  " + count_destination + "  "
														+ Folder + "   Extracting messsage " + message.getSubject());
									} catch (Exception e) {
										e.printStackTrace();
										continue;
									}

								}
							}
						}

					}
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderforPSTOST_defaultrestore(folderInfo, path2, path3, mailfolder);
				}

			} catch (Exception e) {
				continue;
			}

		}

	}

	private void getsubfolderforPSTOST_defaultrestore(FolderInfo f, String path2, String path3, String mailfolder) {

		FolderInfoCollection subfolder = f.getSubFolders();
		String scm = "";
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

				path = path1 + File.separator + Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}
					String path31 = pstfolderlist.get(l).replace(path2 + File.separator, "");
					path = path.replaceAll("[\\[\\]]", "");
					Folder = path;
					if (path31.equalsIgnoreCase(path)) {

						String subfolder1 = clientforexchange_output.createFolder(mailfolder, Folder).getUri();
						mm.lbl_progressreport.setText(" Getting Folder " + Folder);
						if (folderf.getContainerClass().contains("IPF.Appointment")) {

							if (folderf.getContentCount() > 0) {

								MessageInfoCollection messageInfoCollection = null;
								try {
									messageInfoCollection = folderf.getContents();
								} catch (Exception e1) {

									e1.printStackTrace();
								}

								if (!(messageInfoCollection == null)) {

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
											de.setConvertAsTnef(true);
											MailMessage mess = message1.toMailMessage(de);
											if (mm.chckbxMigrateOrBackup.isSelected()) {
												mess.getAttachments().clear();
											}
											MapiMessage message = MapiMessage.fromMailMessage(mess, d);
											if (main_multiplefile.datefilter.isSelected()) {
												datevalidflag = checkdate(message);
											}
											if (message.getMessageClass().equals("IPM.Appointment")
													|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

												try {
													MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();
													ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] {
															null };
													if (!clientforexchange_output.folderExists(
															clientforexchange_output.getMailboxInfo().getCalendarUri(),
															mm.calendertime + "/" + Folder, subfolderInfo)) {
														scm = clientforexchange_output.createFolder(
																clientforexchange_output.getMailboxInfo()
																		.getCalendarUri(),
																mm.calendertime + "/" + Folder, null, "IPF.Appointment")
																.getUri();
													}
													cal.save(
															temppathm + File.separator
																	+ mf.namingconventionmapi(message) + ".ics",
															AppointmentSaveFormat.Ics);
													File file = new File(temppathm + File.separator
															+ mf.namingconventionmapi(message) + ".ics");
													AppointmentLoadOptions optiona = new AppointmentLoadOptions();
													optiona.getIgnoreSmtpAddressCheck();
													Appointment calDoc = Appointment.load(
															temppathm + File.separator
																	+ mf.namingconventionmapi(message) + ".ics",
															optiona);

													if (mm.chckbxRemoveDuplicacy.isSelected()) {
														String input = mm.duplicacymapiCal(cal);
														input = input.replaceAll("\\s", "");
														input = input.trim();
														if (!listdupliccal.contains(input)) {
															listdupliccal.add(input);
															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	clientforexchange_output.createAppointment(calDoc,
																			scm);
																	count_destination++;
																}
															} else {
																clientforexchange_output.createAppointment(calDoc, scm);
																count_destination++;
															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																clientforexchange_output.createAppointment(calDoc, scm);
																count_destination++;
															}
														} else {
															clientforexchange_output.createAppointment(calDoc, scm);
															count_destination++;
														}
													}
													file.delete();
												} catch (OutOfMemoryError ep) {
													mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
															+ mf.namingconventionmapi(message));
												} catch (Exception e) {
													mf.logger.warning("Exception : " + e.getMessage() + "  "
															+ mf.namingconventionmapi(message)
															+ System.lineSeparator());
													e.printStackTrace();
													continue;
												}

											}

											mm.lbl_progressreport.setText(
													"  Total Message Saved Count  " + count_destination + "  " + Folder
															+ "   Extracting messsage " + message.getSubject());

										} catch (Exception e) {
											continue;
										}

									}

								}
							} else if (folderf.getContainerClass().contains("IPF.Task")
									|| folderf.getContainerClass().contains("IPF.StickyNote")) {

								if (folderf.getContentCount() > 0) {

									MessageInfoCollection messageInfoCollection = folderf.getContents();
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
											de.setConvertAsTnef(true);
											MailMessage mess = message1.toMailMessage(de);

											if (mm.chckbxMigrateOrBackup.isSelected()) {
												mess.getAttachments().clear();
											}
											MapiMessage message = MapiMessage.fromMailMessage(mess, d);
											if (main_multiplefile.datefilter.isSelected()) {
												datevalidflag = checkdate(message);
											}
											String taskuri = clientforexchange_output.getMailboxInfo().getTasksUri();

											Date Receiveddate = message.getDeliveryTime();
											if (messageInfo.getMessageClass().equals("IPM.StickyNote")
													|| messageInfo.getMessageClass().equals("IPM.Task")) {
												try {

													MapiTask task = null;
													Boolean checktask = false;
													if (messageInfo.getMessageClass().equals("IPM.Task")) {
														task = (MapiTask) message.toMapiMessageItem();
													}

													if (mm.chckbxRemoveDuplicacy.isSelected()) {
														String input = "";
														if (messageInfo.getMessageClass().equals("IPM.Task")) {
															input = mm.duplicacymapiTask(task);
														}

														if (!listduplictask.contains(input)) {
															System.out.println("Not a duplicate message");
															listduplictask.add(input);

															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	psttask(task, message, taskuri, checktask,
																			messageInfo);
																	count_destination++;
																}

															} else {
																psttask(task, message, taskuri, checktask, messageInfo);
																count_destination++;
															}
														}
													} else {

														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																psttask(task, message, taskuri, checktask, messageInfo);
																count_destination++;
															}
														} else {
															psttask(task, message, taskuri, checktask, messageInfo);
															count_destination++;
														}
													}
												} catch (OutOfMemoryError ep) {
													mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
															+ mf.namingconventionmapi(message));
												} catch (Exception e) {
													mf.logger.warning("Exception : " + e.getMessage() + "Task" + " " + i
															+ mf.namingconventionmapi(message)
															+ System.lineSeparator());
													e.printStackTrace();
													continue;
												}
											}
											mm.lbl_progressreport.setText(
													"  Total Message Saved Count  " + count_destination + "  " + Folder
															+ "   Extracting messsage " + message.getSubject());

										} catch (Exception e) {
											continue;
										}

									}
								}
							} else if (folderf.getContainerClass().contains("IPF.Contact")) {

								System.out.println(Folder + "Folder Name");
								if (folderf.getContentCount() > 0) {

									ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
									if (!clientforexchange_output.folderExists(
											clientforexchange_output.getMailboxInfo().getContactsUri(),
											mm.calendertime + "/" + Folder, subfolderInfo)) {
										scm = clientforexchange_output.createFolder(
												clientforexchange_output.getMailboxInfo().getContactsUri(),
												mm.calendertime + "/" + Folder, null, "IPF.Contact").getUri();
									}

									MessageInfoCollection messageInfoCollection = folderf.getContents();

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
											de.setConvertAsTnef(true);
											MailMessage mess = message1.toMailMessage(de);
											if (mm.chckbxMigrateOrBackup.isSelected()) {
												mess.getAttachments().clear();
											}
											MapiMessage message = MapiMessage.fromMailMessage(mess, d);
											if (main_multiplefile.datefilter.isSelected()) {
												datevalidflag = checkdate(message);
											}
											if (message.getMessageClass().equals("IPM.Contact")) {
												try {

													MapiContact con = (MapiContact) message.toMapiMessageItem();
													ByteArrayOutputStream bos = new ByteArrayOutputStream();
													con.save(bos, ContactSaveFormat.VCard);
													ByteArrayInputStream inStream = new ByteArrayInputStream(
															bos.toByteArray());
													MapiContact mapicontact = MapiContact.fromVCard(inStream);
													if (mm.chckbxRemoveDuplicacy.isSelected()) {
														String input = mm.duplicacymapiContact(con);
														if (!listdupliccontact.contains(input)) {
															listdupliccontact.add(input);
															clientforexchange_output.createContact(scm, mapicontact);
															count_destination++;
														}
													} else {
														clientforexchange_output.createContact(scm, mapicontact);
														count_destination++;
													}
												} catch (Error e) {
													mf.logger.warning(
															"ERROR : " + e.getMessage() + System.lineSeparator());
												} catch (Exception e) {
													mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
															+ i + mf.namingconventionmapi(message)
															+ System.lineSeparator());
													e.printStackTrace();
													continue;
												}

											}

											mm.lbl_progressreport.setText(
													"  Total Message Saved Count  " + count_destination + "  " + Folder
															+ "   Extracting messsage " + message.getSubject());

										} catch (Exception e) {
											continue;
										}

									}
								}
							} else {

								if (folderf.getContentCount() > 0) {
									MessageInfoCollection messageInfoCollection = folderf.getContents();

									messageaddpst(messageInfoCollection, subfolder1, Folder);
								}
							}
							listdupliccal.clear();
							mf.listduplicacy.clear();
							listdupliccontact.clear();
							listduplictask.clear();
						}
						if (folderf.hasSubFolders()) {
							getsubfolderforPSTOST_defaultrestore(folderf, path2, path, subfolder1);
						}

					}
				}

				path = mm.removefolder(path);
			} catch (Exception e) {
				continue;
			}

		}

	}

	@SuppressWarnings("deprecation")
	private void psttask(MapiTask task, MapiMessage message, String taskuri, Boolean checktask,
			MessageInfo messageInfo) {

		if (mm.chckbxRemoveDuplicacy.isSelected()) {

			String input = "";
			if (messageInfo.getMessageClass().equals("IPM.Task")) {
				input = mm.duplicacymapiTask(task);
			}

			if (!listduplictask.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplictask.add(input);
				if (checktask) {
					if (main_multiplefile.datefilter.isSelected()) {
						if (datevalidflag) {
							clientforexchange_output.createTask(taskuri, task);
							count_destination++;
						}
					} else {
						clientforexchange_output.createTask(taskuri, task);
						count_destination++;
					}
				} else {
					if (main_multiplefile.datefilter.isSelected()) {
						if (datevalidflag) {
							clientforexchange_output.createTask(taskuri, task);
							count_destination++;
						}
					} else {
						clientforexchange_output.createTask(taskuri, task);
						count_destination++;
					}
				}

			}
		} else {

			if (checktask) {
				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						clientforexchange_output.createTask(taskuri, task);
						count_destination++;
					}
				} else {
					clientforexchange_output.createTask(taskuri, task);
					count_destination++;

				}
			} else {
				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						clientforexchange_output.createTask(taskuri, task);
						count_destination++;
					}
				} else {
					clientforexchange_output.createTask(taskuri, task);
					count_destination++;

				}
			}

		}

	}

	@SuppressWarnings("deprecation")
	private void messageaddpst(MessageInfoCollection messageInfoCollection, String mailbox, String Folder) {

		String scm = "";
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
				de.setConvertAsTnef(true);
				MailMessage mess = message1.toMailMessage(de);
				if (mm.chckbxMigrateOrBackup.isSelected()) {
					mess.getAttachments().clear();
				}
				MapiMessage message = MapiMessage.fromMailMessage(mess, d);
				if (main_multiplefile.datefilter.isSelected()) {
					datevalidflag = checkdate(message);
					System.out.println(datevalidflag);
				}
				if (message.getMessageClass().equals("IPM.Contact")) {
					try {

						ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };

						if (!clientforexchange_output.folderExists(
								clientforexchange_output.getMailboxInfo().getContactsUri(),
								mm.calendertime + "/" + Folder, subfolderInfo)) {
							scm = clientforexchange_output
									.createFolder(clientforexchange_output.getMailboxInfo().getContactsUri(),
											mm.calendertime + "/" + Folder, null, "IPF.Contact")
									.getUri();
						}

						MapiContact con = (MapiContact) message.toMapiMessageItem();
						ByteArrayOutputStream bos = new ByteArrayOutputStream();
						con.save(bos, ContactSaveFormat.VCard);
						ByteArrayInputStream inStream = new ByteArrayInputStream(bos.toByteArray());
						MapiContact mapicontact = MapiContact.fromVCard(inStream);

						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = mm.duplicacymapiContact(mapicontact);
							if (!listdupliccontact.contains(input)) {
								listdupliccontact.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										clientforexchange_output.createContact(scm, mapicontact);
										count_destination++;
									}
								} else {
									clientforexchange_output.createContact(scm, mapicontact);
									count_destination++;
								}

							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforexchange_output.createContact(scm, mapicontact);
									count_destination++;
								}
							} else {
								clientforexchange_output.createContact(scm, mapicontact);
								count_destination++;
							}

						}

					} catch (Error e) {
						mf.logger.warning("ERROR : " + e.getMessage() + System.lineSeparator());
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " " + i
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else if (message.getMessageClass().equals("IPM.Appointment")
						|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

					try {

						MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

						ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
						if (!clientforexchange_output.folderExists(
								clientforexchange_output.getMailboxInfo().getCalendarUri(),
								mm.calendertime + "/" + Folder, subfolderInfo)) {
							scm = clientforexchange_output
									.createFolder(clientforexchange_output.getMailboxInfo().getCalendarUri(),
											mm.calendertime + "/" + Folder, null, "IPF.Appointment")
									.getUri();
						}

						cal.save(temppathm + File.separator + mf.namingconventionmapi(message) + ".ics",
								AppointmentSaveFormat.Ics);
						File file = new File(temppathm + File.separator + mf.namingconventionmapi(message) + ".ics");
						AppointmentLoadOptions optiona = new AppointmentLoadOptions();
						optiona.getIgnoreSmtpAddressCheck();
						Appointment calDoc = Appointment
								.load(temppathm + File.separator + mf.namingconventionmapi(message) + ".ics", optiona);

						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = mm.duplicacymapiCal(cal);
							input = input.replaceAll("\\s", "");
							input = input.trim();
							if (!listdupliccal.contains(input)) {
								listdupliccal.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										clientforexchange_output.createAppointment(calDoc, scm);
										count_destination++;
									}
								} else {
									clientforexchange_output.createAppointment(calDoc, scm);
									count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforexchange_output.createAppointment(calDoc, scm);
									count_destination++;
								}
							} else {
								clientforexchange_output.createAppointment(calDoc, scm);
								count_destination++;
							}
						}

						file.delete();

					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "  " + mf.namingconventionmapi(message)
								+ System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else if (messageInfo.getMessageClass().equals("IPM.StickyNote")
						|| messageInfo.getMessageClass().equals("IPM.Task")) {
					try {

						Boolean checktask = false;
						MapiTask task = null;
						if (messageInfo.getMessageClass().equals("IPM.Task")) {
							task = (MapiTask) message.toMapiMessageItem();
						}
						String taskuri = clientforexchange_output.getMailboxInfo().getTasksUri();
						psttask(task, message, taskuri, checktask, messageInfo);

					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "Task" + " " + i
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else {

					try {

						mailexchange(mess, clientforexchange_output, mailbox);

					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + i
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				}

				mm.lbl_progressreport.setText("  Total Message Saved Count  " + count_destination + "  " + Folder
						+ "   Extracting messsage " + message.getSubject());

			} catch (Exception e) {
				continue;
			}

		}

	}

	String mailexchange(MailMessage message, IEWSClient clientforexchange_output1, String Folderuri) throws Exception {
		String Messageid = "";
		if (mm.chckbxRemoveDuplicacy.isSelected()) {

			String input = mm.duplicacymail(message);

			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						Messageid = clientforexchange_output1.appendMessage(Folderuri, message);
						count_destination++;
					}
				} else {
					Messageid = clientforexchange_output1.appendMessage(Folderuri, message);
					count_destination++;
				}
			}
		} else {
			if (main_multiplefile.datefilter.isSelected()) {
				if (datevalidflag) {
					Messageid = clientforexchange_output1.appendMessage(Folderuri, message);
					count_destination++;
				}
			} else {
				Messageid = clientforexchange_output1.appendMessage(Folderuri, message);
				count_destination++;
			}
		}
		return Messageid;
	}

	void pstcontact(MapiContact contact, Date Receiveddate, MapiMessage message, FolderInfo info, Boolean checkcon) {

		if (mm.chckbxRemoveDuplicacy.isSelected()) {

			String input = mm.duplicacymapiContact(contact);
			input = input.replaceAll("\\s", "");
			input = input.trim();

			if (!listdupliccontact.contains(input)) {
				System.out.println("Not a duplicate message");
				listdupliccontact.add(input);
				if (checkcon) {
					if (main_multiplefile.datefilter.isSelected()) {
						if (datevalidflag) {
							info.addMessage(message);
							count_destination++;
						}
					} else {
						info.addMessage(message);
						count_destination++;
					}

				} else {

					if (main_multiplefile.datefilter.isSelected()) {
						if (datevalidflag) {
							info.addMessage(message);
							count_destination++;
						}
					} else {
						info.addMessage(message);
						count_destination++;
					}
				}

			}
		} else {
			if (checkcon) {
				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						info.addMessage(message);
						count_destination++;
					}
				} else {
					info.addMessage(message);
					count_destination++;
				}
			} else {
				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						info.addMessage(message);
						count_destination++;
					}
				} else {
					info.addMessage(message);
					count_destination++;
				}
			}

		}

	}

	boolean checkdate(MapiMessage message1) {
		for (int k = 0; k < fromList.size(); k++) {
			fromdate = fromList.get(k);
			todate = toList.get(k);
			System.out.println(" Date : " + message1.getDeliveryTime());
			if (message1.getDeliveryTime() != null) {
				if (message1.getDeliveryTime().after(fromdate) && message1.getDeliveryTime().before(todate)) {
					datevalidflag = true;
					break;

				} else {
					datevalidflag = false;
				}
			}
		}

		return datevalidflag;
	}
}
