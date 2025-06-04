package email.code;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.aspose.email.BodyContentType;
import com.aspose.email.FileFormatVersion;
import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiContactPersonalInfoPropertySet;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiMessageFlags;
import com.aspose.email.MapiTask;
import com.aspose.email.MessageInfo;
import com.aspose.email.MessageInfoCollection;
import com.aspose.email.PersonalStorage;
import com.aspose.email.StandardIpmFolder;

public class ConvertOST_PST implements Runnable {

	static boolean check = false;
	static String word = "Top of Personal Folders";
	static String word1 = "Top of Outlook data file";
	FolderInfo info1;
	long maxsize;
	int pstindex = 0;
	String pstfilename = "";
	int ijj = 1;
	int splitcount = 0;
	String splitpath = "";
	PersonalStorage ost;
	FolderInfo info = new FolderInfo();
	FolderInfo folderInfo;
	private List<String> listduplicacy = new ArrayList<String>();
	private List<String> listdupliccal = new ArrayList<String>();
	private List<String> listduplictask = new ArrayList<String>();
	private List<String> listdupliccontact = new ArrayList<String>();
	static Date fromdate;
	static Date todate;
	ArrayList<Date> fromList = new ArrayList<Date>();
	ArrayList<Date> toList = new ArrayList<Date>();
	String from;
	String to;
	String first = null, middle = null, last = null;
	private Main_Frame mf;
	private main_multiplefile mm;
	private String filetype = "";
	private String filepath = "";
	private String destination_path = "";
	private String path = "";
	static long count_destination;
	long count_AllMails;
	static PersonalStorage pst;
	String Folder;
	List<String> pstfolderlist;
	static boolean datevalidflag = false;
	static long foldermessagecount = 0;

	public ConvertOST_PST(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, long maxsize) {
		this.mf = mf;
		this.mm = mm;
		this.filetype = filetype;
		this.filepath = filepath;
		this.destination_path = destination_path;
		ConvertOST_PST.count_destination = count_destination;
		this.pstfolderlist = pstfolderlist;
		this.fromList = fromList;
		this.toList = toList;
		this.maxsize = maxsize;
	}

	@Override
	public void run() {

		convertOST_PST(mf, filetype, destination_path, count_destination, filepath, mm, pstfolderlist, fromList, toList,
				maxsize);
		main_multiplefile.count_destination = ConvertOST_PST.count_destination;
		System.out.println(count_AllMails);
	}

	private void convertOST_PST(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, long maxsize) {

		splitpath = destination_path + File.separator + mm.fname + splitcount + ".pst";
		ost = PersonalStorage.create(splitpath, FileFormatVersion.Unicode);
		ost.getStore().changeDisplayName(mm.fname);
		pst = PersonalStorage.fromFile(filepath);
		splitcount = 0;
		// String path2 = "";
		ConvertOST_PST.count_destination = 0;
		count_AllMails = 0;
		FolderInfo folderInfo2 = pst.getRootFolder();

		Folder = folderInfo2.getDisplayName();
		Folder = Folder.replace(",", "").replace(".", "");
		Folder = Main_Frame.getRidOfIllegalFileNameCharacters(Folder);
		Folder = Folder.replaceAll("[\\[\\]]", "");
		Folder = Folder.trim();
		if (Folder.equalsIgnoreCase("")) {
			Folder = "Root Folder";
		}

		path = Folder;
		String path2 = Folder;
		// path = path + File.separator + Folder;
//		path2 = Folder;

		MessageInfoCollection messageInfoCollection1 = folderInfo2.getContents();

		info = ost.getRootFolder();

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
		String sss;
		if (filetype.contains("OST")) {
			sss = ".ost";

		} else {
			sss = ".pst";
		}
		listduplicacy.clear();
		listdupliccal.clear();
		listdupliccontact.clear();
		listduplictask.clear();
		pstindex = 0;

		File filechk = new File(destination_path + File.separator + mm.fname + pstindex + sss);
		for (int i = 0; i < messagesize1; i++) {
			try {

				if (mm.stop) {
					break;
				}
				if ((i % 100) == 0) {
					System.gc();

				}

				long currentsize = filechk.length();
				if (mm.chckbx_splitpst.isSelected()) {
					if (currentsize > maxsize) {
						pstindex++;

						ost = PersonalStorage.create(destination_path + File.separator + mm.fname + pstindex + sss,
								FileFormatVersion.Unicode);
						ost.getStore().changeDisplayName(mm.fname);
						filechk = new File(destination_path + File.separator + mm.fname + pstindex + sss);

						info = ost.getRootFolder();

					}
				}

				MessageInfo messageInfo = (MessageInfo) messageInfoCollection1.get_Item(i);

				MapiMessage message1 = pst.extractMessage(messageInfo);
				MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
				MailConversionOptions de = new MailConversionOptions();
				de.setConvertAsTnef(true);
				MailMessage mess = message1.toMailMessage(de);
				MapiMessage message = MapiMessage.fromMailMessage(mess, d);
				if (mm.chckbxMigrateOrBackup.isSelected()) {
					message1.getAttachments().clear();
					message.getAttachments().clear();
					message1.getAttachments().clear();
					mess.getAttachments().clear();
				}

				int bct = message.getBodyType();
				if (bct == 0) {
					message.setBodyContent(message.getBodyHtml(), BodyContentType.Html);
				} else {
					message.setBodyContent(message.getBodyRtf(), BodyContentType.Rtf);
				}

				if (main_multiplefile.datefilter.isSelected()) {
					datevalidflag = mm.checkdate(message1, mess);
				}
				try {

					if (mm.chckbxRemoveDuplicacy.isSelected()) {

						String input = mm.duplicacymapi(message1);
						input = input.replaceAll("\\s", "");
						input = input.trim();

						if (!listduplicacy.contains(input)) {
							listduplicacy.add(input);

							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									ConvertOST_PST.count_destination++;

									if (((message1.getFlags()
											& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
										info.addMessage(message1);
									} else {
										message1.setMessageFlags(MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
										info.addMessage(message1);

									}

								}
							} else {
								ConvertOST_PST.count_destination++;

								if (((message1.getFlags()
										& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
									info.addMessage(message1);
								} else {
									message1.setMessageFlags(MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
									info.addMessage(message1);

								}

							}
						}
					} else {
						if (main_multiplefile.datefilter.isSelected()) {
							if (datevalidflag) {
								ConvertOST_PST.count_destination++;

								if (((message1.getFlags()
										& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
									info.addMessage(message1);
								} else {
									message1.setMessageFlags(MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
									info.addMessage(message1);

								}

							}
						} else {
							ConvertOST_PST.count_destination++;

							if (((message1.getFlags()
									& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
								info.addMessage(message1);
							} else {
								message1.setMessageFlags(MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
								info.addMessage(message1);

							}

						}

					}

				} catch (OutOfMemoryError ep) {
					mf.logger.info("Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
				} catch (Exception e) {
					mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + i
							+ mf.namingconventionmapi(message) + System.lineSeparator());
					e.printStackTrace();
					continue;
				}

				mm.lbl_progressreport.setText("  Total Message Saved Count  " + ConvertOST_PST.count_destination + "  "
						+ Folder + "   Extracting messsage " + message.getSubject());

			} catch (Exception e) {
				continue;
			}

		}
		FolderInfoCollection folderInf = pst.getRootFolder().getSubFolders();

		for (int j = 0; j < folderInf.size(); j++) {
			try {

				FolderInfo folderInfo = folderInf.get_Item(j);
				String Folder = folderInfo.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = Main_Frame.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();
				if (mm.stop) {
					break;
				}
//				path = path + File.separator + Folder;
//				String path3 = Folder;
				path = Folder;
				try {
					for (int l = 0; l < pstfolderlist.size(); l++) {
						if (mm.stop) {
							break;
						}
						String path3 = pstfolderlist.get(l).replace(path2 + File.separator, "");
//						String path1 = pstfolderlist.get(l);

						// String path1 = "";
//						String str = pstfolderlist.get(l);
//						if (pstfolderlist.get(l).contains(word) || pstfolderlist.get(l).contains(word1)) {
//							str = str.replace("\\", "_");
//							String msg[] = str.split("_");
//							String new_str = "";
//
//							for (int i1 = 0; i1 < msg.length; i1++) {
//								if (i1 == 0 && msg[i1].equals(word) || msg[i1].equals(word1)) {
//								} else {
//									check = true;
//									new_str += msg[i1] + "\\";
//								}
//							}
//
//							if (check) {
//								path1 = (String) new_str.subSequence(0, new_str.length() - 1);
//								check = false;
//							}
//						} else {
//							path1 = str;
//						}
//						System.out.println(path1 + "  path1");
//						System.out.println(path3 + "  path3");
						if (path3.equalsIgnoreCase(path)) {
							mm.lbl_progressreport.setText(" Getting Folder " + Folder);
							listduplicacy.clear();
							listdupliccal.clear();
							listdupliccontact.clear();
							listduplictask.clear();

							path = Folder;

							if (Folder.contains("Inbox")) {

								if (folderInfo.getContentCount() > 0) {
									info1 = ost.createPredefinedFolder(Folder, StandardIpmFolder.Inbox, true);
									MessageInfoCollection messageInfoCollection = folderInfo.getContents();

									messageaddOst(messageInfoCollection, info1, Folder);
								}
							} else if (Folder.contains("Deleted Item")) {

								if (folderInfo.getContentCount() > 0) {
									info1 = ost.getPredefinedFolder(StandardIpmFolder.DeletedItems);
									MessageInfoCollection messageInfoCollection = folderInfo.getContents();
									messageaddOst(messageInfoCollection, info1, Folder);
								}
							} else if (Folder.contains("Calendar")) {

								if (folderInfo.getContentCount() > 0) {
									info = ost.createPredefinedFolder(Folder, StandardIpmFolder.Appointments, true);
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
											de.setConvertAsTnef(true);
											MailMessage mess = message1.toMailMessage(de);
											MapiMessage message = MapiMessage.fromMailMessage(mess, d);
											int bct = message.getBodyType();
											if (bct == 0) {
												message.setBodyContent(message.getBodyHtml(), BodyContentType.Html);
											} else {
												message.setBodyContent(message.getBodyRtf(), BodyContentType.Rtf);
											}
											if (main_multiplefile.datefilter.isSelected()) {
												datevalidflag = mm.checkdate(message1, mess);
											}

											if (message1.getMessageClass().equals("IPM.Appointment") || message1
													.getMessageClass().equals("IPM.Schedule.Meeting.Request")) {

												try {

													MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();

													if (mm.chckbxRemoveDuplicacy.isSelected()) {
														String input = mm.duplicacymapiCal(cal);
														input = input.replaceAll("\\s", "");
														input = input.trim();
														if (!listdupliccal.contains(input)) {
															listdupliccal.add(input);
															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	info.addMessage(message1);
																	ConvertOST_PST.count_destination++;

																}
															} else {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														} else {
															info.addMessage(message1);
															ConvertOST_PST.count_destination++;

														}
													}
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
											} else {

												try {

													if (ost.getRootFolder().getSubFolder(path, true) != null) {

														info = ost.getRootFolder().getSubFolder(path, true);
													} else {

														info = ost.getRootFolder().addSubFolder(path, true);
													}

													if (mm.chckbxRemoveDuplicacy.isSelected()) {

														String input = mm.duplicacymapi(message1);
														input = input.replaceAll("\\s", "");
														input = input.trim();

														if (!listduplicacy.contains(input)) {
															listduplicacy.add(input);

															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	ConvertOST_PST.count_destination++;
																	count_AllMails++;
																	if (((message1.getFlags()
																			& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																		info.addMessage(message1);
																	} else {
																		message1.setMessageFlags(
																				MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																		info.addMessage(message1);

																	}

																}
															} else {
																ConvertOST_PST.count_destination++;
																count_AllMails++;
																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);

																}

															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																ConvertOST_PST.count_destination++;
																count_AllMails++;
																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);

																}

															}
														} else {
															ConvertOST_PST.count_destination++;
															count_AllMails++;
															if (((message1.getFlags()
																	& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																info.addMessage(message1);
															} else {
																message1.setMessageFlags(
																		MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																info.addMessage(message1);

															}

														}

													}

												} catch (OutOfMemoryError ep) {
													mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
															+ mf.namingconventionmapi(message));
												} catch (Exception e) {
													mf.logger.warning("Exception : " + e.getMessage() + "Message" + " "
															+ i + mf.namingconventionmapi(message)
															+ System.lineSeparator());
													e.printStackTrace();
													continue;
												}

											}
											mm.lbl_progressreport.setText("  Total Message Saved Count  "
													+ ConvertOST_PST.count_destination + "  " + Folder
													+ "   Extracting messsage " + message.getSubject());

										} catch (Exception e) {
											continue;
										}

									}
								}
							} else if (Folder.contains("Tasks") || Folder.contains("ToDo")) {

								if (folderInfo.getContentCount() > 0) {
									info = ost.createPredefinedFolder(Folder, StandardIpmFolder.Tasks, true);
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
											de.setConvertAsTnef(true);
											MailMessage mess = message1.toMailMessage(de);
											MapiMessage message = MapiMessage.fromMailMessage(mess, d);
											if (mm.chckbxMigrateOrBackup.isSelected()) {
												message1.getAttachments().clear();
												message.getAttachments().clear();
												message1.getAttachments().clear();
												mess.getAttachments().clear();
											}

											int bct = message.getBodyType();
											if (bct == 0) {
												message.setBodyContent(message.getBodyHtml(), BodyContentType.Html);
											} else {
												message.setBodyContent(message.getBodyRtf(), BodyContentType.Rtf);
											}
											if (main_multiplefile.datefilter.isSelected()) {
												datevalidflag = mm.checkdate(message1, mess);
											}
											if (message1.getMessageClass().equals("IPM.Task")) {
												try {

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
															listduplictask.add(input);
															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	info.addMessage(message1);
																	ConvertOST_PST.count_destination++;

																}
															} else {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														}
													} else {

														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														} else {
															info.addMessage(message1);
															ConvertOST_PST.count_destination++;

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

											} else {

												try {

													if (ost.getRootFolder().getSubFolder(path, true) != null) {

														info = ost.getRootFolder().getSubFolder(path, true);
													} else {

														info = ost.getRootFolder().addSubFolder(path, true);
													}

													if (mm.chckbxRemoveDuplicacy.isSelected()) {

														String input = mm.duplicacymapi(message1);
														input = input.replaceAll("\\s", "");
														input = input.trim();

														if (!listduplicacy.contains(input)) {
															listduplicacy.add(input);

															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	ConvertOST_PST.count_destination++;

																	if (((message1.getFlags()
																			& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																		info.addMessage(message1);
																	} else {
																		message1.setMessageFlags(
																				MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																		info.addMessage(message1);

																	}

																}
															} else {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);

																}

															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);

																}

															}
														} else {
															ConvertOST_PST.count_destination++;

															if (((message1.getFlags()
																	& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																info.addMessage(message1);
															} else {
																message1.setMessageFlags(
																		MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																info.addMessage(message1);

															}

														}

													}

												} catch (OutOfMemoryError ep) {
													mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
															+ mf.namingconventionmapi(message));
												} catch (Exception e) {
													mf.logger.warning("Exception : " + e.getMessage() + "Message" + " "
															+ i + mf.namingconventionmapi(message)
															+ System.lineSeparator());
													e.printStackTrace();
													continue;
												}

											}

											mm.lbl_progressreport.setText("  Total Message Saved Count  "
													+ ConvertOST_PST.count_destination + "  " + Folder
													+ "   Extracting messsage " + message.getSubject());

										} catch (Exception e) {
											continue;
										}

									}
								}
							} else if (Folder.contains("Contacts")) {
								if (folderInfo.getContentCount() > 0) {

									info = ost.createPredefinedFolder(Folder, StandardIpmFolder.Contacts, true);
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
											de.setConvertAsTnef(true);
											MailMessage mess = message1.toMailMessage(de);
											MapiMessage message = MapiMessage.fromMailMessage(mess, d);
											int bct = message.getBodyType();
											if (bct == 0) {
												message.setBodyContent(message.getBodyHtml(), BodyContentType.Html);
											} else {
												message.setBodyContent(message.getBodyRtf(), BodyContentType.Rtf);
											}
											if (main_multiplefile.datefilter.isSelected()) {
												datevalidflag = mm.checkdate(message1, mess);
											}
											if (message1.getMessageClass().equals("IPM.Contact")) {
												try {

													MapiContact con = (MapiContact) message1.toMapiMessageItem();
													MapiContactPersonalInfoPropertySet mp = con.getPersonalInfo();
													mp.getLocation();

													if (mm.chckbxRemoveDuplicacy.isSelected()) {

														String input = mm.duplicacymapiContact(con);
														input = input.replaceAll("\\s", "");
														input = input.trim();

														if (!listdupliccontact.contains(input)) {
															listdupliccontact.add(input);
															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	info.addMessage(message1);
																	ConvertOST_PST.count_destination++;

																}
															} else {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														} else {
															info.addMessage(message1);
															ConvertOST_PST.count_destination++;

														}
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

											} else {

												try {

													if (ost.getRootFolder().getSubFolder(path, true) != null) {

														info = ost.getRootFolder().getSubFolder(path, true);
													} else {

														info = ost.getRootFolder().addSubFolder(path, true);
													}

													if (mm.chckbxRemoveDuplicacy.isSelected()) {

														String input = mm.duplicacymapi(message1);
														input = input.replaceAll("\\s", "");
														input = input.trim();

														if (!listduplicacy.contains(input)) {
															listduplicacy.add(input);

															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	ConvertOST_PST.count_destination++;

																	if (((message1.getFlags()
																			& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																		info.addMessage(message1);
																	} else {
																		message1.setMessageFlags(
																				MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																		info.addMessage(message1);

																	}

																}
															} else {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);

																}

															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);

																}

															}
														} else {
															ConvertOST_PST.count_destination++;

															if (((message1.getFlags()
																	& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																info.addMessage(message1);
															} else {
																message1.setMessageFlags(
																		MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																info.addMessage(message1);

															}

														}

													}

												} catch (OutOfMemoryError ep) {
													mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
															+ mf.namingconventionmapi(message));
												} catch (Exception e) {
													mf.logger.warning("Exception : " + e.getMessage() + "Message" + " "
															+ i + mf.namingconventionmapi(message)
															+ System.lineSeparator());
													e.printStackTrace();
													continue;
												}

											}

											mm.lbl_progressreport.setText("  Total Message Saved Count  "
													+ ConvertOST_PST.count_destination + "  " + Folder
													+ "   Extracting messsage " + message.getSubject());

										} catch (Exception e) {
											continue;
										}

									}
								}
							} else if (Folder.contains("Outbox")) {

								if (folderInfo.getContentCount() > 0) {
									info1 = ost.createPredefinedFolder(Folder, StandardIpmFolder.Outbox, true);
									MessageInfoCollection messageInfoCollection = folderInfo.getContents();
									messageaddOst(messageInfoCollection, info1, Folder);
								}
							} else if (Folder.contains("Draft")) {

								if (folderInfo.getContentCount() > 0) {
									info1 = ost.createPredefinedFolder(Folder, StandardIpmFolder.Drafts, true);
									MessageInfoCollection messageInfoCollection = folderInfo.getContents();
									messageaddOst(messageInfoCollection, info1, Folder);
								}
							} else if (Folder.contains("Junk Email")) {

								if (folderInfo.getContentCount() > 0) {
									info1 = ost.createPredefinedFolder(Folder, StandardIpmFolder.JunkEmail, true);
									MessageInfoCollection messageInfoCollection = folderInfo.getContents();
									messageaddOst(messageInfoCollection, info1, Folder);
								}
							} else if (Folder.contains("Notes")) {

								if (folderInfo.getContentCount() > 0) {
									info1 = ost.createPredefinedFolder(Folder, StandardIpmFolder.Notes, true);
									MessageInfoCollection messageInfoCollection = folderInfo.getContents();
									messageaddOst(messageInfoCollection, info1, Folder);
								}
							} else if (Folder.contains("SyncIssues")) {

								if (folderInfo.getContentCount() > 0) {
									info1 = ost.createPredefinedFolder(Folder, StandardIpmFolder.SyncIssues, true);
									MessageInfoCollection messageInfoCollection = folderInfo.getContents();
									messageaddOst(messageInfoCollection, info1, Folder);
								}
							} else if (Folder.contains("Journal")) {

								if (folderInfo.getContentCount() > 0) {
									info1 = ost.createPredefinedFolder(Folder, StandardIpmFolder.Journal, true);
									MessageInfoCollection messageInfoCollection = folderInfo.getContents();
									messageaddOst(messageInfoCollection, info1, Folder);
								}
							} else if (Folder.contains("Sent")) {

								if (folderInfo.getContentCount() > 0) {
									info1 = ost.createPredefinedFolder(Folder, StandardIpmFolder.SentItems, true);
									MessageInfoCollection messageInfoCollection = folderInfo.getContents();
									messageaddOst(messageInfoCollection, info1, Folder);
								}
							} else {

								if (folderInfo.getContentCount() > 0) {
									MessageInfoCollection messageInfoCollection = folderInfo.getContents();
									listdupliccal.clear();
									listdupliccontact.clear();
									listduplictask.clear();
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
											int bct = message.getBodyType();
											if (bct == 0) {
												message.setBodyContent(message.getBodyHtml(), BodyContentType.Html);
											} else {
												message.setBodyContent(message.getBodyRtf(), BodyContentType.Rtf);
											}
											if (main_multiplefile.datefilter.isSelected()) {
												datevalidflag = mm.checkdate(message1, mess);
											}
											if (ost.getRootFolder().getSubFolder(path, true) != null) {

												info = ost.getRootFolder().getSubFolder(path, true);
											} else {

												info = ost.getRootFolder().addSubFolder(path, true);
											}
											if (message1.getMessageClass().equals("IPM.Contact")) {
												try {

													MapiContact con = (MapiContact) message1.toMapiMessageItem();
													if (mm.chckbxRemoveDuplicacy.isSelected()) {
														String input = mm.duplicacymapiContact(con);
														input = input.replaceAll("\\s", "");
														input = input.trim();
														if (!listdupliccontact.contains(input)) {
															listdupliccontact.add(input);
															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	info.addMessage(message1);
																	ConvertOST_PST.count_destination++;

																}
															} else {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														} else {
															info.addMessage(message1);
															ConvertOST_PST.count_destination++;

														}
													}

												} catch (OutOfMemoryError ep) {
													mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
															+ mf.namingconventionmapi(message));
												} catch (Exception e) {
													mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
															+ i + mf.namingconventionmapi(message)
															+ System.lineSeparator());
													continue;
												}

											} else if (message1.getMessageClass().equals("IPM.Appointment") || message1
													.getMessageClass().equals("IPM.Schedule.Meeting.Request")) {

												try {

													MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();
													if (mm.chckbxRemoveDuplicacy.isSelected()) {

														String input = mm.duplicacymapiCal(cal);
														input = input.replaceAll("\\s", "");
														input = input.trim();

														if (!listdupliccal.contains(input)) {
															listdupliccal.add(input);
															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	ConvertOST_PST.count_destination++;

																	if (((message1.getFlags()
																			& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																		info.addMessage(message1);
																	} else {
																		message1.setMessageFlags(
																				MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																		info.addMessage(message1);
																	}
																}
															} else {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);

																}
															}
														} else {
															info.addMessage(message1);
															ConvertOST_PST.count_destination++;

														}
													}

												} catch (OutOfMemoryError ep) {
													mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
															+ mf.namingconventionmapi(message));
												} catch (Exception e) {
													mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " "
															+ i + mf.namingconventionmapi(message)
															+ System.lineSeparator());
													e.printStackTrace();
													continue;
												}

											} else if (message1.getMessageClass().equals("IPM.Task")) {
												try {
													if (mm.chckbxRemoveDuplicacy.isSelected()) {
														String input = "";
														MapiTask task = null;
														if (message1.getMessageClass().equals("IPM.Task")) {
															task = (MapiTask) message1.toMapiMessageItem();
														}
														if (message1.getMessageClass().equals("IPM.Task")) {
															input = mm.duplicacymapiTask(task);
														}
														input = input.replaceAll("\\s", "");
														input = input.trim();

														if (!listduplictask.contains(input)) {

															listduplictask.add(input);
															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	ConvertOST_PST.count_destination++;

																	if (((message1.getFlags()
																			& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																		info.addMessage(message1);
																	} else {
																		message1.setMessageFlags(
																				MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																		info.addMessage(message1);
																	}
																}
															} else {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);

																}
															}
														} else {
															info.addMessage(message1);
															ConvertOST_PST.count_destination++;

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

											} else {

												try {

													if (ost.getRootFolder().getSubFolder(path, true) != null) {

														info = ost.getRootFolder().getSubFolder(path, true);
													} else {

														info = ost.getRootFolder().addSubFolder(path, true);
													}

													if (mm.chckbxRemoveDuplicacy.isSelected()) {

														String input = mm.duplicacymapi(message1);
														input = input.replaceAll("\\s", "");
														input = input.trim();

														if (!listduplicacy.contains(input)) {
															listduplicacy.add(input);

															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	ConvertOST_PST.count_destination++;

																	if (((message1.getFlags()
																			& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																		info.addMessage(message1);
																	} else {
																		message1.setMessageFlags(
																				MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																		info.addMessage(message1);

																	}

																}
															} else {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);

																}

															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);

																}

															}
														} else {
															ConvertOST_PST.count_destination++;

															if (((message1.getFlags()
																	& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																info.addMessage(message1);
															} else {
																message1.setMessageFlags(
																		MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																info.addMessage(message1);

															}

														}

													}

												} catch (OutOfMemoryError ep) {
													mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
															+ mf.namingconventionmapi(message));
												} catch (Exception e) {
													mf.logger.warning("Exception : " + e.getMessage() + "Message" + " "
															+ i + mf.namingconventionmapi(message)
															+ System.lineSeparator());
													e.printStackTrace();
													continue;
												}

											}
											mm.lbl_progressreport.setText("  Total Message Saved Count  "
													+ ConvertOST_PST.count_destination + "  " + Folder
													+ "   Extracting messsage " + message.getSubject());
										} catch (Exception e) {
											e.printStackTrace();
											continue;
										}

									}
								}
							}

						}
					}
				} catch (Exception e) {

				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderfor_OST_PST(folderInfo, path2);
				}

			} catch (Exception e) {
				continue;
			}

		}

	}

	public void getsubfolderfor_OST_PST(FolderInfo f, String path2) {

		FolderInfoCollection subfolder = f.getSubFolders();

		for (int k = 0; k < subfolder.size(); k++) {
			try {
				if (mm.stop) {
					break;
				}
				FolderInfo folderf = subfolder.get_Item(k);

				String Folder = folderf.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = Main_Frame.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();

				// path = path1 + File.separator + Folder;
				path = path + File.separator + Folder;

				try {
					for (int l = 0; l < pstfolderlist.size(); l++) {
						if (mm.stop) {
							break;
						}
//						String path3 = pstfolderlist.get(l).replace(path2 + File.separator, "");
//						String path3 = pstfolderlist.get(l);

						//
//						String path3 = "";
//						String str = pstfolderlist.get(l);
//						if (pstfolderlist.get(l).contains(word) || pstfolderlist.get(l).contains(word1)) {
//							str = str.replace("\\", "_");
//							String msg[] = str.split("_");
//							String new_str = "";
//
//							for (int i1 = 0; i1 < msg.length; i1++) {
//								if (i1 == 0 && msg[i1].equals(word) || msg[i1].equals(word1)) {
//								} else {
//									check = true;
//									new_str += msg[i1] + "\\";
//								}
//							}
//
//							if (check) {
//								path3 = (String) new_str.subSequence(0, new_str.length() - 1);
//								check = false;
//							}
//						} else {
//							path3 = str;
//						}
						String path3 = pstfolderlist.get(l).trim().replace(path2 + File.separator, "");

						if (path3.equalsIgnoreCase(path)) {
							listduplicacy.clear();
							listdupliccal.clear();
							listdupliccontact.clear();
							listduplictask.clear();
							mm.lbl_progressreport.setText(" Getting Folder " + Folder);
							if (folderf.getContainerClass().contains("IPF.Appointment")) {

								if (folderf.getContentCount() > 0) {

									if (ost.getRootFolder().getSubFolder(path, true) != null) {

										info = ost.getRootFolder().getSubFolder(path, true);
									} else {

										info = ost.getRootFolder().addSubFolder(path, true);
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
											MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
											MailConversionOptions de = new MailConversionOptions();
											de.setConvertAsTnef(true);
											MailMessage mess = message1.toMailMessage(de);
											if (mm.chckbxMigrateOrBackup.isSelected()) {
												message1.getAttachments().clear();
											}
											MapiMessage message = MapiMessage.fromMailMessage(mess, d);

											Date Receiveddate = message.getDeliveryTime();
											if (main_multiplefile.datefilter.isSelected()) {
												datevalidflag = mm.checkdate(message1, mess);
											}
											if (message1.getMessageClass().equals("IPM.Appointment") || message1
													.getMessageClass().equals("IPM.Schedule.Meeting.Request")) {

												try {

													MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();
													if (mm.chckbxRemoveDuplicacy.isSelected()) {

														String input = mm.duplicacymapiCal(cal);
														input = input.replaceAll("\\s", "");
														input = input.trim();

														if (!listdupliccal.contains(input)) {
															listdupliccal.add(input);

															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	info.addMessage(message1);
																	ConvertOST_PST.count_destination++;

																}
															} else {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														} else {
															info.addMessage(message1);
															ConvertOST_PST.count_destination++;

														}
													}

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

											} else if (message1.getMessageClass().equals("IPM.Contact")) {
												try {

													MapiContact con = (MapiContact) message1.toMapiMessageItem();

													Boolean checkcon = false;
													ostcontact(con, Receiveddate, message, info, checkcon, message1);

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

											} else if (message1.getMessageClass().equals("IPM.Task")) {
												try {

													Boolean checktask = false;
													MapiTask task = null;
													if (messageInfo.getMessageClass().equals("IPM.Task")) {
														task = (MapiTask) message1.toMapiMessageItem();
													}

													osttask(task, Receiveddate, message, info, checktask, messageInfo,
															message1);

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

											} else {

												try {

													if (mm.chckbxRemoveDuplicacy.isSelected()) {

														String input = mm.duplicacymapi(message1);
														input = input.replaceAll("\\s", "");
														input = input.trim();

														if (!listduplicacy.contains(input)) {
															listduplicacy.add(input);

															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	ConvertOST_PST.count_destination++;

																	if (((message1.getFlags()
																			& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																		info.addMessage(message1);
																	} else {
																		message1.setMessageFlags(
																				MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																		info.addMessage(message1);
																	}
																}
															} else {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);
																}
															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);
																}
															}
														} else {
															ConvertOST_PST.count_destination++;

															if (((message1.getFlags()
																	& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																info.addMessage(message1);
															} else {
																message1.setMessageFlags(
																		MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																info.addMessage(message1);
															}
														}
													}
												} catch (OutOfMemoryError ep) {
													mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
															+ mf.namingconventionmapi(message));
												} catch (Exception e) {
													mf.logger.warning("Exception : " + e.getMessage() + "Message" + " "
															+ i + mf.namingconventionmapi(message)
															+ System.lineSeparator());
													e.printStackTrace();
													continue;
												}

											}

											mm.lbl_progressreport.setText("  Total Message Saved Count  "
													+ ConvertOST_PST.count_destination + "  " + Folder
													+ "   Extracting messsage " + message.getSubject());

										} catch (Exception e) {
											continue;
										}

									}

								}
							} else if (folderf.getContainerClass().contains("IPF.Task")) {

								if (folderf.getContentCount() > 0) {

									if (ost.getRootFolder().getSubFolder(path, true) != null) {

										info = ost.getRootFolder().getSubFolder(path, true);
									} else {

										info = ost.getRootFolder().addSubFolder(path, true);
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
											MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
											MailConversionOptions de = new MailConversionOptions();
											de.setConvertAsTnef(true);
											MailMessage mess = message1.toMailMessage(de);
											if (mm.chckbxMigrateOrBackup.isSelected()) {
												message1.getAttachments().clear();
											}
											MapiMessage message = MapiMessage.fromMailMessage(mess, d);

											Date Receiveddate = message.getDeliveryTime();
											if (main_multiplefile.datefilter.isSelected()) {
												datevalidflag = mm.checkdate(message1, mess);
											}
											if (message1.getMessageClass().equals("IPM.Task")) {
												try {

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
															listduplictask.add(input);

															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	info.addMessage(message1);
																	ConvertOST_PST.count_destination++;

																}
															} else {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														}
													} else {

														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														} else {
															info.addMessage(message1);
															ConvertOST_PST.count_destination++;

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

											} else if (message1.getMessageClass().equals("IPM.Contact")) {
												try {

													MapiContact con = (MapiContact) message1.toMapiMessageItem();
													Boolean checkcon = false;
													ostcontact(con, Receiveddate, message, info, checkcon, message1);

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

											} else if (message1.getMessageClass().equals("IPM.Appointment") || message1
													.getMessageClass().equals("IPM.Schedule.Meeting.Request")) {

												try {

													MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();
													Boolean checkcal = false;
													ostcalendar(cal, Receiveddate, message, info, checkcal, message1);

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

											} else {

												try {

													if (mm.chckbxRemoveDuplicacy.isSelected()) {

														String input = mm.duplicacymapi(message1);
														input = input.replaceAll("\\s", "");
														input = input.trim();

														if (!listduplicacy.contains(input)) {
															listduplicacy.add(input);

															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	ConvertOST_PST.count_destination++;

																	if (((message1.getFlags()
																			& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																		info.addMessage(message1);
																	} else {
																		message1.setMessageFlags(
																				MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																		info.addMessage(message1);
																	}
																}
															} else {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);
																}
															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);
																}
															}
														} else {
															ConvertOST_PST.count_destination++;

															if (((message1.getFlags()
																	& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																info.addMessage(message1);
															} else {
																message1.setMessageFlags(
																		MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																info.addMessage(message1);
															}
														}
													}
												} catch (OutOfMemoryError ep) {
													mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
															+ mf.namingconventionmapi(message));
												} catch (Exception e) {
													mf.logger.warning("Exception : " + e.getMessage() + "Message" + " "
															+ i + mf.namingconventionmapi(message)
															+ System.lineSeparator());
													e.printStackTrace();
													continue;
												}

											}

											mm.lbl_progressreport.setText("  Total Message Saved Count  "
													+ ConvertOST_PST.count_destination + "  " + Folder
													+ "   Extracting messsage " + message.getSubject());

										} catch (Exception e) {
											continue;
										}

									}
								}
							} else if (folderf.getContainerClass().contains("IPF.Contact")) {

								if (folderf.getContentCount() > 0) {

									if (ost.getRootFolder().getSubFolder(path, true) != null) {

										info = ost.getRootFolder().getSubFolder(path, true);
									} else {

										info = ost.getRootFolder().addSubFolder(path, true);
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
											MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
											MailConversionOptions de = new MailConversionOptions();
											de.setConvertAsTnef(true);
											MailMessage mess = message1.toMailMessage(de);
											if (mm.chckbxMigrateOrBackup.isSelected()) {
												message1.getAttachments().clear();
											}
											MapiMessage message = MapiMessage.fromMailMessage(mess, d);
											Date Receiveddate = message.getDeliveryTime();
											if (main_multiplefile.datefilter.isSelected()) {
												datevalidflag = mm.checkdate(message1, mess);
											}
											if (message1.getMessageClass().equals("IPM.Contact")) {
												try {

													MapiContact con = (MapiContact) message1.toMapiMessageItem();
													if (mm.chckbxRemoveDuplicacy.isSelected()) {

														String input = mm.duplicacymapiContact(con);
														input = input.replaceAll("\\s", "");
														input = input.trim();

														if (!listdupliccontact.contains(input)) {
															listdupliccontact.add(input);
															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	info.addMessage(message1);
																	ConvertOST_PST.count_destination++;

																}
															} else {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																info.addMessage(message1);
																ConvertOST_PST.count_destination++;

															}
														} else {
															info.addMessage(message1);
															ConvertOST_PST.count_destination++;

														}

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

											} else if (message1.getMessageClass().equals("IPM.Contact")) {
												try {

													MapiContact con = (MapiContact) message1.toMapiMessageItem();
													Boolean checkcon = false;
													ostcontact(con, Receiveddate, message, info, checkcon, message1);

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

											} else if (message1.getMessageClass().equals("IPM.Appointment") || message1
													.getMessageClass().equals("IPM.Schedule.Meeting.Request")) {

												try {

													MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();
													Boolean checkcal = false;
													ostcalendar(cal, Receiveddate, message, info, checkcal, message1);

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

											} else if (message1.getMessageClass().equals("IPM.Task")) {
												try {

													Boolean checktask = false;
													MapiTask task = null;
													if (message1.getMessageClass().equals("IPM.Task")) {
														task = (MapiTask) message1.toMapiMessageItem();
													}

													osttask(task, Receiveddate, message, info, checktask, messageInfo,
															message1);

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

											} else {

												try {

													if (mm.chckbxRemoveDuplicacy.isSelected()) {

														String input = mm.duplicacymapi(message1);
														input = input.replaceAll("\\s", "");
														input = input.trim();

														if (!listduplicacy.contains(input)) {
															listduplicacy.add(input);
															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	ConvertOST_PST.count_destination++;

																	if (((message1.getFlags()
																			& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																		info.addMessage(message1);
																	} else {
																		message1.setMessageFlags(
																				MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																		info.addMessage(message1);
																	}
																}
															} else {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);
																}
															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																ConvertOST_PST.count_destination++;

																if (((message1.getFlags()
																		& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																	info.addMessage(message1);
																} else {
																	message1.setMessageFlags(
																			MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																	info.addMessage(message1);
																}
															}
														} else {
															ConvertOST_PST.count_destination++;

															if (((message1.getFlags()
																	& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
																info.addMessage(message1);
															} else {
																message1.setMessageFlags(
																		MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
																info.addMessage(message1);
															}
														}
													}
												} catch (OutOfMemoryError ep) {
													mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
															+ mf.namingconventionmapi(message));
												} catch (Exception e) {
													mf.logger.warning("Exception : " + e.getMessage() + "Message" + " "
															+ i + mf.namingconventionmapi(message)
															+ System.lineSeparator());
													e.printStackTrace();
													continue;
												}
											}

											mm.lbl_progressreport.setText("  Total Message Saved Count  "
													+ ConvertOST_PST.count_destination + "  " + Folder
													+ "   Extracting messsage " + message.getSubject());

										} catch (Exception e) {
											continue;
										}

									}
								}
							} else {
								MessageInfoCollection messageInfoCollection = folderf.getContents();
								if (ost.getRootFolder().getSubFolder(path, true) != null) {

									info1 = ost.getRootFolder().getSubFolder(path, true);
								} else {

									info1 = ost.getRootFolder().addSubFolder(path, true);
								}
								messageaddOst(messageInfoCollection, info1, Folder);
							}
							listdupliccal.clear();
							listduplicacy.clear();
							listdupliccontact.clear();
							listduplictask.clear();

						}
					}
				} catch (Exception e) {

					e.printStackTrace();
				}
				if (folderf.hasSubFolders()) {
					getsubfolderfor_OST_PST(folderf, path2);
				}
				path = mf.removefolder(path);
			} catch (Exception e) {
				continue;
			}

		}

	}

	public void messageaddOst(MessageInfoCollection messageInfoCollection, FolderInfo info, String Folder) {

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
		String sss;
		if (filetype.contains("OST")) {
			sss = ".ost";

		} else {
			sss = ".pst";
		}
		File filechk = new File(destination_path + File.separator + mm.fname + pstindex + sss);

		for (int i = 0; i < messagesize; i++) {
			try {

				if (mm.stop) {
					break;
				}
				if ((i % 100) == 0) {
					System.gc();

				}

				long currentsize = filechk.length();
				if (mm.chckbx_splitpst.isSelected()) {
					if (currentsize > maxsize) {
						pstindex++;
						filechk = new File(destination_path + File.separator + mm.fname + pstindex + sss);
						ost = PersonalStorage.create(filechk.getAbsolutePath(), FileFormatVersion.Unicode);
						ost.getStore().changeDisplayName(mm.fname);
						if (ost.getRootFolder().getSubFolder(path, true) != null) {
							info = ost.getRootFolder().getSubFolder(path, true);
						} else {
							info = ost.getRootFolder().addSubFolder(path, true);
						}
					}
				}

				MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

				MapiMessage message1 = pst.extractMessage(messageInfo);
				MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
				MailConversionOptions de = new MailConversionOptions();
				de.setConvertAsTnef(true);
				MailMessage mess = message1.toMailMessage(de);

				if (mm.chckbxMigrateOrBackup.isSelected()) {
					message1.getAttachments().clear();
				}
				MapiMessage message = MapiMessage.fromMailMessage(mess, d);
				int bct = message.getBodyType();
				if (bct == 0) {
					message.setBodyContent(message.getBodyHtml(), BodyContentType.Html);
				} else {
					message.setBodyContent(message.getBodyRtf(), BodyContentType.Rtf);
				}

				Date Receiveddate = message.getDeliveryTime();
				if (main_multiplefile.datefilter.isSelected()) {
					datevalidflag = mm.checkdate(message1, mess);
				}
				if (message1.getMessageClass().equals("IPM.Contact")) {
					try {

						MapiContact con = (MapiContact) message1.toMapiMessageItem();

						Boolean checkcon = false;
						ostcontact(con, Receiveddate, message, info, checkcon, message1);

					} catch (Error e) {
						mf.logger.warning("ERROR : " + e.getMessage() + System.lineSeparator());
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " " + i
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else if (message1.getMessageClass().equals("IPM.Appointment")
						|| message1.getMessageClass().equals("IPM.Schedule.Meeting.Request")) {

					try {
						MapiCalendar cal = null;
						try {
							cal = (MapiCalendar) message1.toMapiMessageItem();
							Boolean checkcal = false;
							ostcalendar(cal, Receiveddate, message, info, checkcal, message1);
						} catch (Exception e) {
							Boolean checkcal = false;
							ostcalendar(cal, Receiveddate, message, info, checkcal, message1);
						}
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "  " + mf.namingconventionmapi(message)
								+ System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else if (message1.getMessageClass().equals("IPM.Task")) {
					try {

						Boolean checktask = false;
						MapiTask task = null;
						if (message1.getMessageClass().equals("IPM.Task")) {
							task = (MapiTask) message1.toMapiMessageItem();
						}

						osttask(task, Receiveddate, message, info, checktask, messageInfo, message1);

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

						if (mm.chckbxRemoveDuplicacy.isSelected()) {

							String input = mm.duplicacymapi(message1);
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listduplicacy.contains(input)) {
								listduplicacy.add(input);

								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										ConvertOST_PST.count_destination++;

										if (((message1.getFlags()
												& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
											info.addMessage(message1);
										} else {
											message1.setMessageFlags(MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
											info.addMessage(message1);
										}
									}
								} else {
									ConvertOST_PST.count_destination++;

									if (((message1.getFlags()
											& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
										info.addMessage(message1);
									} else {
										message1.setMessageFlags(MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
										info.addMessage(message1);
									}
								}
							}
						} else {

							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									ConvertOST_PST.count_destination++;

									if (((message1.getFlags()
											& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
										info.addMessage(message1);
									} else {
										message1.setMessageFlags(MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
										info.addMessage(message1);
									}
								}
							} else {
								ConvertOST_PST.count_destination++;

								if (((message1.getFlags()
										& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
									info.addMessage(message1);
								} else {
									message1.setMessageFlags(MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);
									info.addMessage(message1);

								}

							}

						}

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

				mm.lbl_progressreport.setText("  Total Message Saved Count  " + ConvertOST_PST.count_destination + "  "
						+ Folder + "   Extracting messsage " + message.getSubject());

			} catch (Exception e) {
				continue;
			}

		}

	}

	void ostcontact(MapiContact contact, Date Receiveddate, MapiMessage message, FolderInfo info, Boolean checkcon,
			MapiMessage message1) {

		if (mm.chckbxRemoveDuplicacy.isSelected()) {

			String input = mm.duplicacymapiContact(contact);
			input = input.replaceAll("\\s", "");
			input = input.trim();

			if (!listdupliccontact.contains(input)) {
				listdupliccontact.add(input);
				if (checkcon) {
					if (main_multiplefile.datefilter.isSelected()) {
						if (datevalidflag) {
							info.addMessage(message1);
							ConvertOST_PST.count_destination++;

						}
					} else {
						info.addMessage(message1);
						ConvertOST_PST.count_destination++;

					}
				} else {
					if (main_multiplefile.datefilter.isSelected()) {
						if (datevalidflag) {
							info.addMessage(message1);
							ConvertOST_PST.count_destination++;

						}
					} else {
						info.addMessage(message1);
						ConvertOST_PST.count_destination++;

					}
				}
			}
		} else {
			if (checkcon) {
				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						info.addMessage(message1);
						ConvertOST_PST.count_destination++;

					}
				} else {
					info.addMessage(message1);
					ConvertOST_PST.count_destination++;

				}
			} else {
				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						info.addMessage(message1);
						ConvertOST_PST.count_destination++;

					}
				} else {
					info.addMessage(message1);
					ConvertOST_PST.count_destination++;

				}

			}

		}

	}

	void ostcalendar(MapiCalendar calendar, Date Receiveddate, MapiMessage msg, FolderInfo info, Boolean checkcal,
			MapiMessage message1) {

		if (mm.chckbxRemoveDuplicacy.isSelected()) {

			String input = mm.duplicacymapiCal(calendar);
			input = input.replaceAll("\\s", "");
			input = input.trim();

			if (!listdupliccal.contains(input)) {
				listdupliccal.add(input);
				if (checkcal) {
					if (main_multiplefile.datefilter.isSelected()) {
						if (datevalidflag) {
							info.addMessage(message1);
							ConvertOST_PST.count_destination++;

						}
					} else {
						info.addMessage(message1);
						ConvertOST_PST.count_destination++;

					}
				} else {
					if (main_multiplefile.datefilter.isSelected()) {
						if (datevalidflag) {
							info.addMessage(message1);
							ConvertOST_PST.count_destination++;

						}
					} else {
						info.addMessage(message1);
						ConvertOST_PST.count_destination++;

					}
				}
			}
		} else {
			if (checkcal) {
				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						info.addMessage(message1);
						ConvertOST_PST.count_destination++;

					}
				} else {
					info.addMessage(message1);
					ConvertOST_PST.count_destination++;

				}
			} else {
				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						info.addMessage(message1);
						ConvertOST_PST.count_destination++;

					}
				} else {
					info.addMessage(message1);
					ConvertOST_PST.count_destination++;

				}
			}
		}
	}

	void osttask(MapiTask task, Date Receiveddate, MapiMessage message, FolderInfo info, Boolean checktask,
			MessageInfo messageInfo, MapiMessage message1) {

		if (mm.chckbxRemoveDuplicacy.isSelected()) {

			String input = "";
			if (message1.getMessageClass().equals("IPM.Task")) {
				input = mm.duplicacymapiTask(task);
			}
			input = input.replaceAll("\\s", "");
			input = input.trim();

			if (!listduplictask.contains(input)) {
				listduplictask.add(input);
				if (checktask) {
					if (main_multiplefile.datefilter.isSelected()) {
						if (datevalidflag) {
							info.addMessage(message1);
							ConvertOST_PST.count_destination++;

						}
					} else {
						info.addMessage(message1);
						ConvertOST_PST.count_destination++;

					}
				} else {
					if (main_multiplefile.datefilter.isSelected()) {
						if (datevalidflag) {
							info.addMessage(message1);
							ConvertOST_PST.count_destination++;

						}
					} else {
						info.addMessage(message1);
						ConvertOST_PST.count_destination++;

					}
				}
			}
		} else {

			if (checktask) {
				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						info.addMessage(message1);
						ConvertOST_PST.count_destination++;

					}
				} else {
					info.addMessage(message1);
					ConvertOST_PST.count_destination++;

				}
			} else {
				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						info.addMessage(message1);
						ConvertOST_PST.count_destination++;

					}
				} else {
					info.addMessage(message1);
					ConvertOST_PST.count_destination++;

				}
			}

		}

	}

}
