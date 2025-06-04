package email.code;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.aspose.email.AppointmentSaveFormat;
import com.aspose.email.Attachment;
import com.aspose.email.ContactSaveFormat;
import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiContactNamePropertySet;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiTask;
import com.aspose.email.MboxrdStorageWriter;
import com.aspose.email.MessageInfo;
import com.aspose.email.MessageInfoCollection;
import com.aspose.email.PersonalStorage;

public class ConvertPSTOST_mbox implements Runnable {
	long foldermessagecount;
	List<String> listdupliccal = new ArrayList<String>();
	List<String> listduplictask = new ArrayList<String>();
	List<String> listdupliccontact = new ArrayList<String>();
	List<String> listduplicacy = new ArrayList<String>();
	static Date fromdate;
	static Date todate;
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

	public ConvertPSTOST_mbox(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm) {
		this.mf = mf;
		this.mm = mm;
		this.filetype = filetype;
		this.filepath = filepath;
		this.destination_path = destination_path;
		ConvertPSTOST_mbox.count_destination = count_destination;
		this.pstfolderlist = pstfolderlist;
		this.fromList = fromList;
		this.toList = toList;
		this.temppathm = temppathm;
	}

	@SuppressWarnings("resource")
	@Override
	public void run() {

		if (filetype.equalsIgnoreCase("Thunderbird")) {

			new MboxrdStorageWriter(destination_path.replace(mm.fname + ".sbd", "") + mm.fname, false);
		}

		convertPSTOST_mbox(mf, filetype, destination_path, count_destination, filepath, mm, pstfolderlist, fromList,
				toList);
		main_multiplefile.count_destination = ConvertPSTOST_mbox.count_destination;
	}

	@SuppressWarnings("resource")
	private void convertPSTOST_mbox(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList) {

		pst = PersonalStorage.fromFile(filepath);
		MailConversionOptions options = new MailConversionOptions();
		FolderInfo folderInfo2 = pst.getRootFolder();
		ConvertPSTOST_mbox.count_destination = 0;
		Folder = folderInfo2.getDisplayName();
		Folder = Folder.replace(",", "").replace(".", "");
		Folder = Main_Frame.getRidOfIllegalFileNameCharacters(Folder);
		Folder = Folder.replaceAll("[\\[\\]]", "");
		Folder = Folder.trim();
		if (Folder.equalsIgnoreCase("")) {
			Folder = "Root Folder";
		}
		String path1 = "";
		String sop = "";
		path = path + File.separator + Folder;
		path1 = Folder;
		listdupliccal.clear();
		listduplictask.clear();
		listdupliccontact.clear();
		listduplicacy.clear();
		MessageInfoCollection messageInfoCollection1 = folderInfo2.getContents();
		if (filetype.equalsIgnoreCase("Thunderbird")) {
			sop = path;
			path = path + ".sbd";
		}

		new File(destination_path + File.separator + path).mkdirs();

		MboxrdStorageWriter wr1 = null;

		if (filetype.equalsIgnoreCase("Opera Mail")) {
			wr1 = new MboxrdStorageWriter(destination_path + File.separator + sop + File.separator
					+ Main_Frame.getRidOfIllegalFileNameCharacters(Folder) + ".mbs", false);
		} else if (filetype.equalsIgnoreCase("Thunderbird")) {
			wr1 = new MboxrdStorageWriter(destination_path + File.separator + sop, false);
		} else {
			wr1 = new MboxrdStorageWriter(destination_path + File.separator + sop + File.separator
					+ Main_Frame.getRidOfIllegalFileNameCharacters(Folder) + ".mbx", false);
		}
		int countr = 0;
		int messagesize1;
		if (main_multiplefile.demo) {
			if (messageInfoCollection1.size() <= All_Data.demo_count) {
				messagesize1 = messageInfoCollection1.size();
			} else {
				messagesize1 = All_Data.demo_count;
			}

		} else {
			messagesize1 = messageInfoCollection1.size();
		}
		for (int i = 0; i < messagesize1; i++) {
			try {

				if (mm.stop) {
					break;
				}
				if ((i % 100) == 0) {
					System.gc();

				}

				MessageInfo messageInfo = (MessageInfo) messageInfoCollection1.get_Item(i);

				MapiMessage message1 = pst.extractMessage(messageInfo);
				MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
				MailConversionOptions de = new MailConversionOptions();
				MailMessage mess1 = message1.toMailMessage(de);
				MapiMessage message = MapiMessage.fromMailMessage(mess1, d);
				
				if (mm.chckbxMigrateOrBackup.isSelected()) {
					message1.getAttachments().clear();
					message.getAttachments().clear();
					mess1.getAttachments().clear();
				}
				
				
				if (main_multiplefile.datefilter.isSelected()) {
					datevalidflag = mm.checkdate(message1, mess1);
				}
				Date Receiveddate = message.getDeliveryTime();
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

						try {
							mapi.setDate(Receiveddate);
						} catch (Exception e) {
							mapi.setDate(null);
						}
						con.setBody(mess1.getBody());
						MapiContactNamePropertySet NamePropSet = con.getNameInfo();

						if (NamePropSet.getGivenName() != null) {
							first = NamePropSet.getGivenName();
						} else {
							first = "";
						}
						if (NamePropSet.getMiddleName() != null) {
							middle = NamePropSet.getMiddleName();
						} else {
							middle = "";
						}
						if (NamePropSet.getSurname() != null) {
							last = NamePropSet.getSurname();
						} else {
							last = "";
						}
						MapiContactNamePropertySet NameProp = new MapiContactNamePropertySet();
						NameProp.setDisplayName(first + " " + middle + " " + last);
						con.setNameInfo(NameProp);

						for (Attachment attachment : mess1.getAttachments()) {
							mapi.addAttachment(attachment);
						}

						con.save(temppathm + File.separator + i + mf.namingconventionmapi(message) + ".vcf",
								ContactSaveFormat.VCard);
						File file = new File(
								temppathm + File.separator + i + mf.namingconventionmapi(message) + ".vcf");
						mapi.addAttachment(new Attachment(
								temppathm + File.separator + i + mf.namingconventionmapi(message) + ".vcf"));
						file.delete();

						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = mm.duplicacymapiContact(con);
							input = input.replaceAll("\\s", "");
							input = input.trim();
							if (!listdupliccontact.contains(input)) {
								listdupliccontact.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										wr1.writeMessage(mapi);
										ConvertPSTOST_mbox.count_destination++;
									}
								} else {
									wr1.writeMessage(mapi);
									ConvertPSTOST_mbox.count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									wr1.writeMessage(mapi);
									ConvertPSTOST_mbox.count_destination++;
								}
							} else {
								wr1.writeMessage(mapi);
								ConvertPSTOST_mbox.count_destination++;
							}
						}
						countr++;
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning(
								"Exception : " + e.getMessage() + "Contact" + " " + countr + System.lineSeparator());
						return;
					}

				} else if (message.getMessageClass().equals("IPM.Appointment")
						|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {

					MailMessage mapi = new MailMessage();

					try {
						for (Attachment attachment : mess1.getAttachments()) {
							mapi.addAttachment(attachment);
						}
						MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();
						try {
							mapi.setSubject(cal.getSubject());
						} catch (Exception e) {

							mapi.setSubject("");
						}
						try {
							mapi.setHtmlBody(cal.getBodyHtml());
						} catch (Exception e) {

							mapi.setBody("");
						}
						try {
							mapi.setDate(message.getDeliveryTime());
						} catch (Exception e) {
							mapi.setDate(null);
						}
						cal.save(temppathm + File.separator + i + mf.namingconventionmapi(message) + ".ics",
								AppointmentSaveFormat.Ics);
						File file = new File(
								temppathm + File.separator + i + mf.namingconventionmapi(message) + ".ics");

						mapi.addAttachment(new Attachment(
								temppathm + File.separator + i + mf.namingconventionmapi(message) + ".ics"));
						file.delete();

						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = mm.duplicacymapiCal(cal);
							input = input.replaceAll("\\s", "");
							input = input.trim();
							if (!listdupliccal.contains(input)) {
								listdupliccal.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										mm.lbl_progressreport.setText("Extracting message " + message.getSubject());
										wr1.writeMessage(mapi);
										ConvertPSTOST_mbox.count_destination++;
										mm.progressBar_message_p3.setValue(100);
									}
								} else {
									mm.lbl_progressreport.setText("Extracting message " + message.getSubject());
									wr1.writeMessage(mapi);
									ConvertPSTOST_mbox.count_destination++;
									mm.progressBar_message_p3.setValue(100);
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									mm.lbl_progressreport.setText("Extracting message " + message.getSubject());
									wr1.writeMessage(mapi);
									ConvertPSTOST_mbox.count_destination++;
									mm.progressBar_message_p3.setValue(100);
								}
							} else {
								mm.lbl_progressreport.setText("Extracting message " + message.getSubject());
								wr1.writeMessage(mapi);
								ConvertPSTOST_mbox.count_destination++;
								mm.progressBar_message_p3.setValue(100);
							}
						}

					} catch (Error ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + System.lineSeparator());
						return;
					}
				} else if (message.getMessageClass().equals("IPM.Task")) {
					try {
						MailMessage mess = message.toMailMessage(options);
						MapiTask task = null;
						if (message.getMessageClass().equals("IPM.Task")) {
							task = (MapiTask) message.toMapiMessageItem();
						}
						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = mm.duplicacymapiTask(task);
							input = input.replaceAll("\\s", "");
							input = input.trim();
							if (!listduplictask.contains(input)) {
								listduplictask.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										wr1.writeMessage(mess);
										ConvertPSTOST_mbox.count_destination++;
									}
								} else {
									wr1.writeMessage(mess);
									ConvertPSTOST_mbox.count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									wr1.writeMessage(mess);
									ConvertPSTOST_mbox.count_destination++;
								}
							} else {
								wr1.writeMessage(mess);
								ConvertPSTOST_mbox.count_destination++;
							}
						}
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning(
								"Exception : " + e.getMessage() + "Task" + " " + countr + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else {
					try {
						MailMessage mess = message.toMailMessage(options);
						mailmbox(mess, wr1, message);

					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + i
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						continue;
					}

				}

				mm.lbl_progressreport.setText("  Total Message Saved Count  " + ConvertPSTOST_mbox.count_destination
						+ "  " + Folder + "   Extracting messsage " + message.getSubject());

			} catch (Exception e) {
				continue;
			}

		}
		wr1.dispose();

		FolderInfoCollection folderInf = pst.getRootFolder().getSubFolders();

		for (int j = 0; j < folderInf.size(); j++) {
			try {
				if (mm.stop) {
					break;
				}
				FolderInfo folderInfo = folderInf.get_Item(j);
				String Folder = folderInfo.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = Main_Frame.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();

				mm.lbl_progressreport.setText(" Getting Folder " + Folder);
				String sop1 = "";

				if (filetype.equalsIgnoreCase("Thunderbird")) {
					sop1 = path + File.separator + Folder;
					path = path + File.separator + Folder + ".sbd";

				} else {
					path = path + File.separator + Folder;
				}

				String path3 = path1 + File.separator + Folder;
				listdupliccal.clear();
				listduplictask.clear();
				listdupliccontact.clear();
				listduplicacy.clear();
				try {
					for (int l = 0; l < pstfolderlist.size(); l++) {
						if (mm.stop) {
							break;
						}
						if (pstfolderlist.get(l).equalsIgnoreCase(path3)) {

							new File(destination_path + File.separator + path).mkdirs();
							MboxrdStorageWriter wr = null;

							if (filetype.equalsIgnoreCase("Opera Mail")) {
								wr = new MboxrdStorageWriter(destination_path
										+ File.separator + path + File.separator + main_multiplefile
												.getRidOfIllegalFileNameCharacters(folderInfo.getDisplayName())
										+ ".mbs", false);
							} else if (filetype.equalsIgnoreCase("Thunderbird")) {
								wr = new MboxrdStorageWriter(destination_path + File.separator + sop1, false);
							} else {
								wr = new MboxrdStorageWriter(destination_path
										+ File.separator + path + File.separator + main_multiplefile
												.getRidOfIllegalFileNameCharacters(folderInfo.getDisplayName())
										+ ".mbx", false);
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
									MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
									MailConversionOptions de = new MailConversionOptions();
									MailMessage mess = message1.toMailMessage(de);
									MapiMessage message = MapiMessage.fromMailMessage(mess, d);

									if (mm.chckbxMigrateOrBackup.isSelected()) {
										message1.getAttachments().clear();
										message.getAttachments().clear();
										mess.getAttachments().clear();
									}
									if (main_multiplefile.datefilter.isSelected()) {
										datevalidflag = mm.checkdate(message1, mess);
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

											try {
												mapi.setDate(message.getDeliveryTime());
											} catch (Exception e) {
												mapi.setDate(null);
											}
											con.setBody(mess.getBody());
											MapiContactNamePropertySet NamePropSet = con.getNameInfo();

											if (NamePropSet.getGivenName() != null) {
												first = NamePropSet.getGivenName();
											} else {
												first = "";
											}
											if (NamePropSet.getMiddleName() != null) {
												middle = NamePropSet.getMiddleName();
											} else {
												middle = "";
											}
											if (NamePropSet.getSurname() != null) {
												last = NamePropSet.getSurname();
											} else {
												last = "";
											}
											MapiContactNamePropertySet NameProp = new MapiContactNamePropertySet();
											NameProp.setDisplayName(first + " " + middle + " " + last);
											con.setNameInfo(NameProp);

											for (Attachment attachment : mess.getAttachments()) {
												mapi.addAttachment(attachment);
											}

											con.save(temppathm + File.separator + i + mf.namingconventionmapi(message)
													+ ".vcf", ContactSaveFormat.VCard);
											File file = new File(temppathm + File.separator + i
													+ mf.namingconventionmapi(message) + ".vcf");
											mapi.addAttachment(new Attachment(temppathm + File.separator + i
													+ mf.namingconventionmapi(message) + ".vcf"));
											file.delete();

											if (mm.chckbxRemoveDuplicacy.isSelected()) {
												String input = mm.duplicacymapiContact(con);
												input = input.replaceAll("\\s", "");
												input = input.trim();
												if (!listdupliccontact.contains(input)) {
													listdupliccontact.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															mm.lbl_progressreport.setText(
																	"Extracting message " + message.getSubject());
															wr.writeMessage(mapi);
															ConvertPSTOST_mbox.count_destination++;
														}

													} else {
														mm.lbl_progressreport
																.setText("Extracting message " + message.getSubject());
														wr.writeMessage(mapi);

														ConvertPSTOST_mbox.count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														mm.lbl_progressreport
																.setText("Extracting message " + message.getSubject());
														wr.writeMessage(mapi);
														ConvertPSTOST_mbox.count_destination++;
													}
												} else {
													mm.lbl_progressreport
															.setText("Extracting message " + message.getSubject());
													wr.writeMessage(mapi);

													ConvertPSTOST_mbox.count_destination++;
												}
											}
											countr++;
										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mf.namingconventionmapi(message));
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " " + countr
													+ System.lineSeparator());
											return;
										}

									} else if (message.getMessageClass().equals("IPM.Appointment")
											|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
										MailMessage mapi = new MailMessage();
										try {

											for (Attachment attachment : mess.getAttachments()) {
												mapi.addAttachment(attachment);
											}
											MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

											try {
												mapi.setSubject(cal.getSubject());
											} catch (Exception e) {

												mapi.setSubject("");
											}
											try {
												mapi.setHtmlBody(cal.getBodyHtml());
											} catch (Exception e) {

												mapi.setBody("");
											}

											try {
												mapi.setDate(message.getDeliveryTime());
											} catch (Exception e) {
												mapi.setDate(null);
											}

											cal.save(temppathm + File.separator + i + mf.namingconventionmapi(message)
													+ ".ics", AppointmentSaveFormat.Ics);
											File file = new File(temppathm + File.separator + i
													+ mf.namingconventionmapi(message) + ".ics");

											mapi.addAttachment(new Attachment(temppathm + File.separator + i
													+ mf.namingconventionmapi(message) + ".ics"));
											file.delete();

											if (mm.chckbxRemoveDuplicacy.isSelected()) {
												String input = mm.duplicacymapiCal(cal);
												input = input.replaceAll("\\s", "");
												input = input.trim();
												if (!listdupliccal.contains(input)) {
													listdupliccal.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															mm.lbl_progressreport.setText(
																	"Extracting message " + message.getSubject());
															wr.writeMessage(mapi);
															ConvertPSTOST_mbox.count_destination++;
														}

													} else {
														mm.lbl_progressreport
																.setText("Extracting message " + message.getSubject());
														wr.writeMessage(mapi);
														ConvertPSTOST_mbox.count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														mm.lbl_progressreport
																.setText("Extracting message " + message.getSubject());
														wr.writeMessage(mapi);
														ConvertPSTOST_mbox.count_destination++;
													}
												} else {
													mm.lbl_progressreport
															.setText("Extracting message " + message.getSubject());
													wr.writeMessage(mapi);
													ConvertPSTOST_mbox.count_destination++;
												}
											}

											countr++;
										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mf.namingconventionmapi(message));
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " "
													+ countr + System.lineSeparator());
											e.printStackTrace();
											return;
										}

									} else if (message.getMessageClass().equals("IPM.Task")) {
										try {

											MailMessage mess1 = message.toMailMessage(options);
											MapiTask task = null;
											if (message.getMessageClass().equals("IPM.Task")) {
												task = (MapiTask) message.toMapiMessageItem();
											}

											if (mm.chckbxRemoveDuplicacy.isSelected()) {
												String input = mm.duplicacymapiTask(task);
												input = input.replaceAll("\\s", "");
												input = input.trim();
												if (!listduplictask.contains(input)) {
													listduplictask.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															wr.writeMessage(mess1);
															ConvertPSTOST_mbox.count_destination++;
														}
													} else {
														wr.writeMessage(mess1);
														ConvertPSTOST_mbox.count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														wr.writeMessage(mess1);
														ConvertPSTOST_mbox.count_destination++;
													}
												} else {
													wr.writeMessage(mess1);
													ConvertPSTOST_mbox.count_destination++;
												}
											}
										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mf.namingconventionmapi(message));
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + "Task" + " " + countr
													+ System.lineSeparator());
											e.printStackTrace();
											continue;
										}

									} else {
										try {
											MailMessage mess1 = message.toMailMessage(options);
											mailmbox(mess1, wr, message);

										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mf.namingconventionmapi(message));
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + i
													+ mf.namingconventionmapi(message) + System.lineSeparator());
											continue;
										}

									}

									mm.lbl_progressreport.setText(
											"  Total Message Saved Count  " + ConvertPSTOST_mbox.count_destination
													+ "  " + Folder + "   Extracting messsage " + message.getSubject());

								} catch (Exception e) {
									continue;
								}

							}
							wr.dispose();
						}
					}
				} catch (Exception e) {

					e.printStackTrace();
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderforPSTOST_Mbox(folderInfo, path3, sop1);
				}
				if (filetype.equalsIgnoreCase("Thunderbird")) {

					path = mm.removefolder(path);
					sop1 = mm.removefolder(sop1);

				} else {
					path = mm.removefolder(path);
				}

			} catch (Exception e) {
				continue;
			}

		}

	}

	@SuppressWarnings("resource")
	private void getsubfolderforPSTOST_Mbox(FolderInfo f, String path3, String sop1) {

		FolderInfoCollection subfolder = f.getSubFolders();
		MailConversionOptions options = new MailConversionOptions();
		for (int k = 0; k < subfolder.size(); k++) {
			try {
				FolderInfo folderf = subfolder.get_Item(k);
				if (mm.stop) {
					break;
				}

				String Folder = folderf.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();
				mm.lbl_progressreport.setText("Getting : " + Folder);

				String sop2 = "";
				if (filetype.equalsIgnoreCase("Thunderbird")) {
					sop2 = sop1 + ".sbd" + File.separator + Folder;
					path3 = path3 + File.separator + Folder;
					path = path + File.separator + Folder + ".sbd";

				} else {
					path = path + File.separator + Folder;
					path3 = path3 + File.separator + Folder;
				}

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}
					if (pstfolderlist.get(l).equalsIgnoreCase(path3)) {
						new File(destination_path + File.separator + path).mkdirs();
						MboxrdStorageWriter wr = null;

						if (filetype.equalsIgnoreCase("Opera Mail")) {
							wr = new MboxrdStorageWriter(destination_path + File.separator + path + File.separator
									+ main_multiplefile.getRidOfIllegalFileNameCharacters(folderf.getDisplayName())
									+ ".mbs", false);
						} else if (filetype.equalsIgnoreCase("Thunderbird")) {
							wr = new MboxrdStorageWriter(destination_path + File.separator + sop2, false);
						} else {
							wr = new MboxrdStorageWriter(destination_path + File.separator + path + File.separator
									+ main_multiplefile.getRidOfIllegalFileNameCharacters(folderf.getDisplayName())
									+ ".mbx", false);
						}

						MessageInfoCollection messageInfoCollection = null;
						try {
							messageInfoCollection = folderf.getContents();
						} catch (Exception e1) {

							e1.printStackTrace();
						}

						if (!(messageInfoCollection == null)) {
							listdupliccal.clear();
							listduplictask.clear();
							listdupliccontact.clear();
							listduplicacy.clear();
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
									MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

									if (mm.stop) {
										break;
									}
									if ((i % 100) == 0) {
										System.gc();
									}
									MapiMessage message1 = pst.extractMessage(messageInfo);
									MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
									MailConversionOptions de = new MailConversionOptions();
									MailMessage mess1 = message1.toMailMessage(de);
									MapiMessage message = MapiMessage.fromMailMessage(mess1, d);
									if (mm.chckbxMigrateOrBackup.isSelected()) {
										message1.getAttachments().clear();
										message.getAttachments().clear();
										mess1.getAttachments().clear();
									}
									if (main_multiplefile.datefilter.isSelected()) {
										datevalidflag = mm.checkdate(message1, mess1);
									}
									if (message.getMessageClass().equals("IPM.Contact")) {
										MailMessage mapi = new MailMessage();
										try {
											MapiContact con = (MapiContact) message.toMapiMessageItem();
											con.setBody(mess1.getBody());
											try {
												mapi.setSubject(con.getSubject());
											} catch (Exception e) {

												mapi.setSubject("");
											}
											try {
												mapi.setBody(mess1.getBody());
											} catch (Exception e) {
												mapi.setBody("");
											}
											try {
												mapi.setDate(message.getDeliveryTime());
											} catch (Exception e) {
												mapi.setDate(null);
											}

											MapiContactNamePropertySet NamePropSet = con.getNameInfo();

											if (NamePropSet.getGivenName() != null) {
												first = NamePropSet.getGivenName();
											} else {
												first = "";
											}
											if (NamePropSet.getMiddleName() != null) {
												middle = NamePropSet.getMiddleName();
											} else {
												middle = "";
											}
											if (NamePropSet.getSurname() != null) {
												last = NamePropSet.getSurname();
											} else {
												last = "";
											}
											MapiContactNamePropertySet NameProp = new MapiContactNamePropertySet();
											NameProp.setDisplayName(first + " " + middle + " " + last);
											con.setNameInfo(NameProp);
											for (Attachment attachment : mess1.getAttachments()) {
												mapi.addAttachment(attachment);
											}

											try {
												mapi.setDate(message.getDeliveryTime());
											} catch (Exception e) {
												mapi.setDate(null);
											}

											con.save(temppathm + File.separator + i + mf.namingconventionmapi(message)
													+ ".vcf", ContactSaveFormat.VCard);
											File file = new File(temppathm + File.separator + i
													+ mf.namingconventionmapi(message) + ".vcf");
											mapi.addAttachment(new Attachment(temppathm + File.separator + i
													+ mf.namingconventionmapi(message) + ".vcf"));
											file.delete();

											if (mm.chckbxRemoveDuplicacy.isSelected()) {
												String input = mm.duplicacymapiContact(con);
												input = input.replaceAll("\\s", "");
												input = input.trim();
												if (!listdupliccontact.contains(input)) {
													listdupliccontact.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															mm.lbl_progressreport.setText(
																	"Extracting message " + message.getSubject());
															wr.writeMessage(mapi);
															ConvertPSTOST_mbox.count_destination++;
															mm.progressBar_message_p3.setValue(100);
														}

													} else {
														mm.lbl_progressreport
																.setText("Extracting message " + message.getSubject());
														wr.writeMessage(mapi);
														ConvertPSTOST_mbox.count_destination++;
														mm.progressBar_message_p3.setValue(100);

													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														mm.lbl_progressreport
																.setText("Extracting message " + message.getSubject());
														wr.writeMessage(mapi);
														ConvertPSTOST_mbox.count_destination++;
														mm.progressBar_message_p3.setValue(100);
													}

												} else {
													mm.lbl_progressreport
															.setText("Extracting message " + message.getSubject());
													wr.writeMessage(mapi);
													ConvertPSTOST_mbox.count_destination++;
													mm.progressBar_message_p3.setValue(100);
												}
											}
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
													+ System.lineSeparator());
											return;
										}

									} else if (message.getMessageClass().equals("IPM.Appointment")
											|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")){

										MailMessage mapi = new MailMessage();
										try {
											for (Attachment attachment : mess1.getAttachments()) {
												mapi.addAttachment(attachment);
											}
											MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();
											try {
												mapi.setSubject(cal.getSubject());
											} catch (Exception e) {

												mapi.setSubject("");
											}
											try {
												mapi.setHtmlBody(cal.getBodyHtml());
											} catch (Exception e) {

												mapi.setBody("");
											}
											try {
												mapi.setDate(message.getDeliveryTime());
											} catch (Exception e) {
												mapi.setDate(null);
											}
											cal.save(temppathm + File.separator + i + mf.namingconventionmapi(message)
													+ ".ics", AppointmentSaveFormat.Ics);
											File file = new File(temppathm + File.separator + i
													+ mf.namingconventionmapi(message) + ".ics");

											mapi.addAttachment(new Attachment(temppathm + File.separator + i
													+ mf.namingconventionmapi(message) + ".ics"));
											file.delete();

											if (mm.chckbxRemoveDuplicacy.isSelected()) {
												String input = mm.duplicacymapiCal(cal);
												input = input.replaceAll("\\s", "");
												input = input.trim();
												if (!listdupliccal.contains(input)) {
													listdupliccal.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															mm.lbl_progressreport.setText(
																	"Extracting message " + message.getSubject());
															wr.writeMessage(mapi);
															ConvertPSTOST_mbox.count_destination++;
															mm.progressBar_message_p3.setValue(100);
														}
													} else {
														mm.lbl_progressreport
																.setText("Extracting message " + message.getSubject());
														wr.writeMessage(mapi);
														ConvertPSTOST_mbox.count_destination++;
														mm.progressBar_message_p3.setValue(100);
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														mm.lbl_progressreport
																.setText("Extracting message " + message.getSubject());
														wr.writeMessage(mapi);
														ConvertPSTOST_mbox.count_destination++;
														mm.progressBar_message_p3.setValue(100);
													}
												} else {
													mm.lbl_progressreport
															.setText("Extracting message " + message.getSubject());
													wr.writeMessage(mapi);
													ConvertPSTOST_mbox.count_destination++;
													mm.progressBar_message_p3.setValue(100);
												}
											}

										} catch (Error ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mf.namingconventionmapi(message));
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " "
													+ System.lineSeparator());
											return;
										}

									} else if (message.getMessageClass().equals("IPM.Task")) {
										try {
											MailMessage msg = message.toMailMessage(options);
											MapiTask task = (MapiTask) message.toMapiMessageItem();

											if (mm.chckbxRemoveDuplicacy.isSelected()) {
												String input = mm.duplicacymapiTask(task);
												input = input.replaceAll("\\s", "");
												input = input.trim();
												if (!listduplictask.contains(input)) {
													listduplictask.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															mm.lbl_progressreport.setText(
																	"Extracting message " + message.getSubject());
															wr.writeMessage(msg);
															ConvertPSTOST_mbox.count_destination++;
															mm.progressBar_message_p3.setValue(100);
														}
													} else {
														mm.lbl_progressreport
																.setText("Extracting message " + message.getSubject());
														wr.writeMessage(msg);
														ConvertPSTOST_mbox.count_destination++;
														mm.progressBar_message_p3.setValue(100);
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														mm.lbl_progressreport
																.setText("Extracting message " + message.getSubject());
														wr.writeMessage(msg);
														ConvertPSTOST_mbox.count_destination++;
														mm.progressBar_message_p3.setValue(100);
													}
												} else {
													mm.lbl_progressreport
															.setText("Extracting message " + message.getSubject());
													wr.writeMessage(msg);
													ConvertPSTOST_mbox.count_destination++;
													mm.progressBar_message_p3.setValue(100);
												}

											}
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + "Task" + " "
													+ System.lineSeparator());
											return;
										}

									} else {
										try {
											MailMessage mess = message.toMailMessage(options);
											mailmbox(mess, wr, message);
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + "Message" + " "
													+ System.lineSeparator());
											continue;
										}

									}

									mm.lbl_progressreport.setText(
											"  Total Message Saved Count  " + ConvertPSTOST_mbox.count_destination
													+ "  " + Folder + "   Extracting messsage " + message.getSubject());
								} catch (Exception e) {
									continue;
								}

							}
							wr.dispose();
						}
					}
				}
				if (folderf.hasSubFolders()) {
					getsubfolderforPSTOST_Mbox(folderf, path3, sop2);
				}

				if (filetype.equalsIgnoreCase("Thunderbird")) {

					path = mm.removefolder(path);
					sop2 = mm.removefolder(sop2);
					path3 = mm.removefolder(path3);
				} else {
					path = mm.removefolder(path);
					path3 = mm.removefolder(path3);

				}

			} catch (Exception e) {
				continue;
			}
		}

	}

	void mailmbox(MailMessage mess, MboxrdStorageWriter wr, MapiMessage message) {
		if (mm.chckbxRemoveDuplicacy.isSelected()) {
			String input = mm.duplicacymapi(message);
			if (!listduplicacy.contains(input)) {
				listduplicacy.add(input);
				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						wr.writeMessage(mess);
						ConvertPSTOST_mbox.count_destination++;
						foldermessagecount++;
					}
				} else {
					wr.writeMessage(mess);
					ConvertPSTOST_mbox.count_destination++;
					foldermessagecount++;
				}
			}
		} else {
			if (main_multiplefile.datefilter.isSelected()) {
				if (datevalidflag) {
					wr.writeMessage(mess);
					ConvertPSTOST_mbox.count_destination++;
					foldermessagecount++;
				}
			} else {
				wr.writeMessage(mess);
				ConvertPSTOST_mbox.count_destination++;
				foldermessagecount++;
			}
		}
	}
}
