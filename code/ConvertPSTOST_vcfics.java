package email.code;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.aspose.email.AppointmentSaveFormat;
import com.aspose.email.ContactSaveFormat;
import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.MessageInfo;
import com.aspose.email.MessageInfoCollection;
import com.aspose.email.PersonalStorage;

public class ConvertPSTOST_vcfics implements Runnable {
	File f1 = null;
	File faa = null;
	String account = null;
	private List<String> listdupliccontact = new ArrayList<String>();
	private List<String> listdupliccal = new ArrayList<String>();
	long foldermessagecount;
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

	public ConvertPSTOST_vcfics(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm) {
		this.mf = mf;
		this.mm = mm;
		this.filetype = filetype;
		this.filepath = filepath;
		this.destination_path = destination_path;
		ConvertPSTOST_vcfics.count_destination = count_destination;
		this.pstfolderlist = pstfolderlist;
		this.fromList = fromList;
		this.toList = toList;
		this.temppathm = temppathm;
	}

	public void run() {
		convertPSTOST_vcfics(mf, filetype, destination_path, count_destination, filepath, mm, pstfolderlist, fromList,
				toList, temppathm);
		main_multiplefile.count_destination = ConvertPSTOST_vcfics.count_destination;
	}

	void convertPSTOST_vcfics(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm) {
		ConvertPSTOST_vcfics.count_destination = 0;
		pst = PersonalStorage.fromFile(filepath);
		String path1 = "";
		FolderInfo folderInfo2 = pst.getRootFolder();
		Folder = folderInfo2.getDisplayName();
		Folder = Folder.replace(",", "").replace(".", "");
		Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
		Folder = Folder.replaceAll("[\\[\\]]", "");
		Folder = Folder.trim();
		if (Folder.equalsIgnoreCase("")) {
			Folder = "Root Folder";
		}
		String clonepath = path;
		path = path + File.separator + Folder;

		path1 = Folder;

		MessageInfoCollection messageInfoCollection1 = folderInfo2.getContents();

		int messagesize1;
		mf.listduplicacy.clear();
		if (main_multiplefile.demo) {
			if (messageInfoCollection1.size() <= All_Data.demo_count) {
				messagesize1 = messageInfoCollection1.size();
			} else {
				messagesize1 = All_Data.demo_count;
			}

		} else {
			messagesize1 = messageInfoCollection1.size();
		}
		listdupliccal.clear();
		listdupliccontact.clear();
		for (int i = 0; i < messagesize1; i++) {
			try {

				if (mm.stop) {
					break;
				}
				if ((i % 100) == 0) {
					System.gc();

				}
				int countr = 0;
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
					datevalidflag = mm.checkdate(message1, mess1);
				}
				if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
						&& mm.chckbxSavePdfAttachment.isSelected() || message1.getMessageClass().equals("IPM.Contact")
						|| message1.getMessageClass().equals("IPM.Appointment")
						|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {
					if (mm.chckbxMaintainFolderStructure.isSelected()) {
						new File(destination_path + File.separator + path).mkdirs();
						clonepath = path;
					}
					f1 = new File(destination_path + File.separator + clonepath + File.separator + "Attachment");
					f1.mkdirs();
				}
				if (message1.getMessageClass().equals("IPM.Contact")) {
					if (message1.getMessageClass().equals("IPM.Contact")) {
						try {
							MapiContact con = (MapiContact) message1.toMapiMessageItem();
							con.setBody(mess1.getBody());
							if (mm.chckbxRemoveDuplicacy.isSelected()) {

								String input = mm.duplicacymapiContact(con);
								input = input.replaceAll("\\s", "");
								input = input.trim();

								if (!listdupliccontact.contains(input)) {
									listdupliccontact.add(input);
									if (main_multiplefile.datefilter.isSelected()) {
										if (datevalidflag) {

											if (!mm.chckbxMigrateOrBackup.isSelected()
													&& message1.getAttachments().size() > 0
													&& mm.chckbxSavePdfAttachment.isSelected()) {
												saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
												con.getAttachments().clear();
											}
											con.save(
													destination_path + File.separator + clonepath + File.separator
															+ mm.namingconventionmapi(message1) + "_" + i + ".vcf",
													ContactSaveFormat.VCard);
											ConvertPSTOST_vcfics.count_destination++;
										}
									} else {

										if (!mm.chckbxMigrateOrBackup.isSelected()
												&& message1.getAttachments().size() > 0
												&& mm.chckbxSavePdfAttachment.isSelected()) {
											saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
											con.getAttachments().clear();
										}
										con.save(
												destination_path + File.separator + clonepath + File.separator
														+ mm.namingconventionmapi(message1) + "_" + i + ".vcf",
												ContactSaveFormat.VCard);
										ConvertPSTOST_vcfics.count_destination++;
									}
								}
							} else {
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {

										if (!mm.chckbxMigrateOrBackup.isSelected()
												&& message1.getAttachments().size() > 0
												&& mm.chckbxSavePdfAttachment.isSelected()) {
											saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
											con.getAttachments().clear();
										}
										con.save(
												destination_path + File.separator + clonepath + File.separator
														+ mm.namingconventionmapi(message1) + "_" + i + ".vcf",
												ContactSaveFormat.VCard);
										ConvertPSTOST_vcfics.count_destination++;
									}
								} else {

									if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
											&& mm.chckbxSavePdfAttachment.isSelected()) {
										saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
										con.getAttachments().clear();
									}
									con.save(
											destination_path + File.separator + clonepath + File.separator
													+ mm.namingconventionmapi(message1) + "_" + i + ".vcf",
											ContactSaveFormat.VCard);
									ConvertPSTOST_vcfics.count_destination++;
								}
							}
							countr++;
						} catch (OutOfMemoryError ep) {
							mf.logger.info(
									"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
						} catch (Exception e) {
							mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " " + countr
									+ System.lineSeparator());
							e.printStackTrace();
							continue;
						}
					} else {
						continue;
					}
				} else if (message1.getMessageClass().equals("IPM.Appointment")
						|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {
					if (filetype.equalsIgnoreCase("ICS")) {
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

											if (!mm.chckbxMigrateOrBackup.isSelected()
													&& message1.getAttachments().size() > 0
													&& mm.chckbxSavePdfAttachment.isSelected()) {
												saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
												cal.getAttachments().clear();
											}
											cal.save(
													destination_path + File.separator + clonepath + File.separator
															+ mm.namingconventionmapi(message) + "_" + i + ".ics",
													AppointmentSaveFormat.Ics);
											ConvertPSTOST_vcfics.count_destination++;
										}
									} else {

										if (!mm.chckbxMigrateOrBackup.isSelected()
												&& message1.getAttachments().size() > 0
												&& mm.chckbxSavePdfAttachment.isSelected()) {
											saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
											cal.getAttachments().clear();
										}
										cal.save(
												destination_path + File.separator + clonepath + File.separator
														+ mm.namingconventionmapi(message) + "_" + i + ".ics",
												AppointmentSaveFormat.Ics);
										ConvertPSTOST_vcfics.count_destination++;
									}
								}
							} else {
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {

										if (!mm.chckbxMigrateOrBackup.isSelected()
												&& message1.getAttachments().size() > 0
												&& mm.chckbxSavePdfAttachment.isSelected()) {
											saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
											cal.getAttachments().clear();
										}
										cal.save(
												destination_path + File.separator + clonepath + File.separator
														+ mm.namingconventionmapi(message) + "_" + i + ".ics",
												AppointmentSaveFormat.Ics);
										ConvertPSTOST_vcfics.count_destination++;
									}
								} else {

									if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
											&& mm.chckbxSavePdfAttachment.isSelected()) {
										saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
										cal.getAttachments().clear();
									}
									cal.save(
											destination_path + File.separator + clonepath + File.separator
													+ mm.namingconventionmapi(message) + "_" + i + ".ics",
											AppointmentSaveFormat.Ics);
									ConvertPSTOST_vcfics.count_destination++;
								}
							}
							countr++;
						} catch (OutOfMemoryError ep) {
							mf.logger.info(
									"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
						} catch (Exception e) {
							mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + i
									+ mf.namingconventionmapi(message) + System.lineSeparator());
							e.printStackTrace();
							continue;
						}
					} else {
						continue;
					}

				} else {
					continue;
				}

				mm.lbl_progressreport.setText("Total Message Saved : " + ConvertPSTOST_vcfics.count_destination + "    "
						+ Folder + " Extracting messsage " + message.getSubject());

			} catch (Exception e) {
				e.printStackTrace();
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
				String Folder = folderInfo.getDisplayName();

				Folder = Folder.replace(",", "").replace(".", "");
				Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();
				path = path + File.separator + Folder;
				String path3 = path1 + File.separator + Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {

					if (mm.stop) {
						break;
					}
					if (pstfolderlist.get(l).trim().equalsIgnoreCase(path3)) {
						mm.lbl_progressreport.setText(" Getting Folder " + Folder);
						int countr = 1;

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
						listdupliccal.clear();
						listdupliccontact.clear();
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
									datevalidflag = mm.checkdate(message1, mess);
								}
								if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
										&& mm.chckbxSavePdfAttachment.isSelected()
										|| message1.getMessageClass().equals("IPM.Contact")
										|| message1.getMessageClass().equals("IPM.Appointment")
										|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {
									if (mm.chckbxMaintainFolderStructure.isSelected()) {
										new File(destination_path + File.separator + path).mkdirs();
										clonepath = path;
									}
									f1 = new File(destination_path + File.separator + clonepath + File.separator
											+ "Attachment");
									f1.mkdirs();
								}
								if (message1.getMessageClass().equals("IPM.Contact")) {
									if (filetype.equalsIgnoreCase("VCF")) {
										try {
											MapiContact con = (MapiContact) message1.toMapiMessageItem();
											con.setBody(mess.getBody());
											if (mm.chckbxRemoveDuplicacy.isSelected()) {

												String input = mm.duplicacymapiContact(con);
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listdupliccontact.contains(input)) {
													listdupliccontact.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {

															if (!mm.chckbxMigrateOrBackup.isSelected()
																	&& message1.getAttachments().size() > 0
																	&& mm.chckbxSavePdfAttachment.isSelected()) {
																saveContactAttachment(mess, message1, destination_path,
																		clonepath, f1);
																con.getAttachments().clear();
															}
															con.save(destination_path + File.separator + clonepath
																	+ File.separator + mm.namingconventionmapi(message1)
																	+ "_" + i + ".vcf", ContactSaveFormat.VCard);
															ConvertPSTOST_vcfics.count_destination++;
														}
													} else {

														if (!mm.chckbxMigrateOrBackup.isSelected()
																&& message1.getAttachments().size() > 0
																&& mm.chckbxSavePdfAttachment.isSelected()) {
															saveContactAttachment(mess, message1, destination_path,
																	clonepath, f1);
															con.getAttachments().clear();
														}
														con.save(destination_path + File.separator + clonepath
																+ File.separator + mm.namingconventionmapi(message1)
																+ "_" + i + ".vcf", ContactSaveFormat.VCard);
														ConvertPSTOST_vcfics.count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {

														if (!mm.chckbxMigrateOrBackup.isSelected()
																&& message1.getAttachments().size() > 0
																&& mm.chckbxSavePdfAttachment.isSelected()) {
															saveContactAttachment(mess, message1, destination_path,
																	clonepath, f1);
															con.getAttachments().clear();
														}
														con.save(destination_path + File.separator + clonepath
																+ File.separator + mm.namingconventionmapi(message1)
																+ "_" + i + ".vcf", ContactSaveFormat.VCard);
														ConvertPSTOST_vcfics.count_destination++;
													}
												} else {

													if (!mm.chckbxMigrateOrBackup.isSelected()
															&& message1.getAttachments().size() > 0
															&& mm.chckbxSavePdfAttachment.isSelected()) {
														saveContactAttachment(mess, message1, destination_path,
																clonepath, f1);
														con.getAttachments().clear();
													}
													con.save(destination_path + File.separator + clonepath
															+ File.separator + mm.namingconventionmapi(message1) + "_"
															+ i + ".vcf", ContactSaveFormat.VCard);
													ConvertPSTOST_vcfics.count_destination++;
												}
											}

											countr++;
										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mf.namingconventionmapi(message));
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " " + countr
													+ System.lineSeparator());
											e.printStackTrace();
											continue;
										}
									} else {
										continue;
									}
								} else if (message1.getMessageClass().equals("IPM.Appointment")
										|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {

									if (filetype.equalsIgnoreCase("ICS")) {
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
															if (!mm.chckbxMigrateOrBackup.isSelected()
																	&& message1.getAttachments().size() > 0
																	&& mm.chckbxSavePdfAttachment.isSelected()) {
																saveContactAttachment(mess, message1, destination_path,
																		clonepath, f1);
																cal.getAttachments().clear();
															}
															cal.save(destination_path + File.separator + clonepath
																	+ File.separator + mm.namingconventionmapi(message)
																	+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
															ConvertPSTOST_vcfics.count_destination++;
														}
													} else {

														if (!mm.chckbxMigrateOrBackup.isSelected()
																&& message1.getAttachments().size() > 0
																&& mm.chckbxSavePdfAttachment.isSelected()) {
															saveContactAttachment(mess, message1, destination_path,
																	clonepath, f1);
															cal.getAttachments().clear();
														}
														cal.save(destination_path + File.separator + clonepath
																+ File.separator + mm.namingconventionmapi(message)
																+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
														ConvertPSTOST_vcfics.count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {

														if (!mm.chckbxMigrateOrBackup.isSelected()
																&& message1.getAttachments().size() > 0
																&& mm.chckbxSavePdfAttachment.isSelected()) {
															saveContactAttachment(mess, message1, destination_path,
																	clonepath, f1);
															cal.getAttachments().clear();
														}
														cal.save(destination_path + File.separator + clonepath
																+ File.separator + mm.namingconventionmapi(message)
																+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
														ConvertPSTOST_vcfics.count_destination++;
													}
												} else {
													if (!mm.chckbxMigrateOrBackup.isSelected()
															&& message1.getAttachments().size() > 0
															&& mm.chckbxSavePdfAttachment.isSelected()) {
														saveContactAttachment(mess, message1, destination_path,
																clonepath, f1);
														cal.getAttachments().clear();
													}
													cal.save(destination_path + File.separator + clonepath
															+ File.separator + mm.namingconventionmapi(message) + "_"
															+ i + ".ics", AppointmentSaveFormat.Ics);
													ConvertPSTOST_vcfics.count_destination++;
												}
											}
											countr++;
										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mf.namingconventionmapi(message));
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + i
													+ mf.namingconventionmapi(message) + System.lineSeparator());
											e.printStackTrace();
											continue;
										}
									} else {
										continue;
									}
								} else {
									continue;
								}
								mm.lbl_progressreport
										.setText("Total Message Saved : " + ConvertPSTOST_vcfics.count_destination
												+ "    " + Folder + " Extracting messsage " + message.getSubject());
							} catch (Exception e) {
								continue;
							}
						}
					}
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderforPSTOSTvcfics(folderInfo, path3, clonepath);
				}

				path = mm.removefolder(path);
			} catch (Exception e) {
				continue;
			}

		}

	}

	void getsubfolderforPSTOSTvcfics(FolderInfo f, String path3, String clonepath) {

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

				path = path + File.separator + Folder;
				path3 = path3 + File.separator + Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}

					if (pstfolderlist.get(l).equalsIgnoreCase(path3)) {
						mm.lbl_progressreport.setText("Getting : " + Folder);

						MessageInfoCollection messageInfoCollection = null;
						try {
							messageInfoCollection = folderf.getContents();
						} catch (Exception e1) {

							e1.printStackTrace();
						}

						if (!(messageInfoCollection == null)) {

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
							listdupliccal.clear();
							listdupliccontact.clear();
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
										datevalidflag = mm.checkdate(message1, mess);
									}
									if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
											&& mm.chckbxSavePdfAttachment.isSelected()
											|| message1.getMessageClass().equals("IPM.Contact")
											|| message1.getMessageClass().equals("IPM.Appointment")
											|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {
										if (mm.chckbxMaintainFolderStructure.isSelected()) {
											new File(destination_path + File.separator + path).mkdirs();
											clonepath = path;
										}
										f1 = new File(destination_path + File.separator + clonepath + File.separator
												+ "Attachment");
										f1.mkdirs();
									}
									if (message1.getMessageClass().equals("IPM.Contact")) {
										if (filetype.equalsIgnoreCase("VCF")) {
											try {
												MapiContact con = (MapiContact) message1.toMapiMessageItem();
												con.setBody(mess.getBody());
												if (mm.chckbxRemoveDuplicacy.isSelected()) {
													String input = mm.duplicacymapiContact(con);
													input = input.replaceAll("\\s", "");
													input = input.trim();
													if (!listdupliccontact.contains(input)) {
														listdupliccontact.add(input);
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {

																if (!mm.chckbxMigrateOrBackup.isSelected()
																		&& message1.getAttachments().size() > 0
																		&& mm.chckbxSavePdfAttachment.isSelected()) {
																	saveContactAttachment(mess, message1,
																			destination_path, clonepath, f1);
																	con.getAttachments().clear();
																}
																con.save(
																		destination_path + File.separator + clonepath
																				+ File.separator
																				+ mm.namingconventionmapi(message1)
																				+ "_" + i + ".vcf",
																		ContactSaveFormat.VCard);

																ConvertPSTOST_vcfics.count_destination++;
															}
														} else {

															if (!mm.chckbxMigrateOrBackup.isSelected()
																	&& message1.getAttachments().size() > 0
																	&& mm.chckbxSavePdfAttachment.isSelected()) {
																saveContactAttachment(mess, message1, destination_path,
																		clonepath, f1);
																con.getAttachments().clear();
															}
															con.save(destination_path + File.separator + clonepath
																	+ File.separator + mm.namingconventionmapi(message1)
																	+ "_" + i + ".vcf", ContactSaveFormat.VCard);

															ConvertPSTOST_vcfics.count_destination++;
														}
													}
												} else {
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {

															if (!mm.chckbxMigrateOrBackup.isSelected()
																	&& message1.getAttachments().size() > 0
																	&& mm.chckbxSavePdfAttachment.isSelected()) {
																saveContactAttachment(mess, message1, destination_path,
																		clonepath, f1);
																con.getAttachments().clear();
															}
															con.save(destination_path + File.separator + clonepath
																	+ File.separator + mm.namingconventionmapi(message1)
																	+ "_" + i + ".vcf", ContactSaveFormat.VCard);

															ConvertPSTOST_vcfics.count_destination++;
														}
													} else {

														if (!mm.chckbxMigrateOrBackup.isSelected()
																&& message1.getAttachments().size() > 0
																&& mm.chckbxSavePdfAttachment.isSelected()) {
															saveContactAttachment(mess, message1, destination_path,
																	clonepath, f1);
															con.getAttachments().clear();
														}
														con.save(destination_path + File.separator + clonepath
																+ File.separator + mm.namingconventionmapi(message1)
																+ "_" + i + ".vcf", ContactSaveFormat.VCard);

														ConvertPSTOST_vcfics.count_destination++;
													}
												}
											} catch (Exception e) {
												mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
														+ System.lineSeparator());
												continue;
											}
										} else {
											continue;
										}

									} else if (message1.getMessageClass().equals("IPM.Appointment")
											|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {
										if (filetype.equalsIgnoreCase("ICS")) {
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

																if (!mm.chckbxMigrateOrBackup.isSelected()
																		&& message1.getAttachments().size() > 0
																		&& mm.chckbxSavePdfAttachment.isSelected()) {
																	saveContactAttachment(mess, message1,
																			destination_path, clonepath, f1);
																	cal.getAttachments().clear();
																}
																cal.save(
																		destination_path + File.separator + clonepath
																				+ File.separator
																				+ mm.namingconventionmapi(message1)
																						.trim()
																				+ "_" + i + ".ics",
																		AppointmentSaveFormat.Ics);
																ConvertPSTOST_vcfics.count_destination++;
															}
														} else {

															if (!mm.chckbxMigrateOrBackup.isSelected()
																	&& message1.getAttachments().size() > 0
																	&& mm.chckbxSavePdfAttachment.isSelected()) {
																saveContactAttachment(mess, message1, destination_path,
																		clonepath, f1);
																cal.getAttachments().clear();
															}
															cal.save(
																	destination_path + File.separator + clonepath
																			+ File.separator
																			+ mm.namingconventionmapi(message1).trim()
																			+ "_" + i + ".ics",
																	AppointmentSaveFormat.Ics);
															ConvertPSTOST_vcfics.count_destination++;
														}
													}
												} else {
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {

															if (!mm.chckbxMigrateOrBackup.isSelected()
																	&& message1.getAttachments().size() > 0
																	&& mm.chckbxSavePdfAttachment.isSelected()) {
																saveContactAttachment(mess, message1, destination_path,
																		clonepath, f1);
																cal.getAttachments().clear();
															}
															cal.save(
																	destination_path + File.separator + clonepath
																			+ File.separator
																			+ mm.namingconventionmapi(message1).trim()
																			+ "_" + i + ".ics",
																	AppointmentSaveFormat.Ics);
															ConvertPSTOST_vcfics.count_destination++;
														}
													} else {

														if (!mm.chckbxMigrateOrBackup.isSelected()
																&& message1.getAttachments().size() > 0
																&& mm.chckbxSavePdfAttachment.isSelected()) {
															saveContactAttachment(mess, message1, destination_path,
																	clonepath, f1);
															cal.getAttachments().clear();
														}
														cal.save(destination_path + File.separator + clonepath
																+ File.separator
																+ mm.namingconventionmapi(message1).trim() + "_" + i
																+ ".ics", AppointmentSaveFormat.Ics);
														ConvertPSTOST_vcfics.count_destination++;
													}
												}
											} catch (Exception e) {
												mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " "
														+ System.lineSeparator());
												continue;
											}
										} else {
											continue;
										}

									} else {
										continue;
									}

									mm.lbl_progressreport.setText(
											"  Total Message Saved Count  " + ConvertPSTOST_vcfics.count_destination
													+ "  " + Folder + "   Extracting messsage " + message.getSubject());
								} catch (Exception e) {
									continue;
								}

							}
						}
					}
				}
				if (folderf.hasSubFolders()) {
					getsubfolderforPSTOSTvcfics(folderf, path3, clonepath);
				}
				path = mm.removefolder(path);
				path3 = mm.removefolder(path3);

			} catch (Exception e) {
				continue;
			}
		}

	}

	void saveContactAttachment(MailMessage mess, MapiMessage message, String destination_path, String clonepath,
			File f1) {
		File f2 = new File(f1.getAbsolutePath() + File.separator + mm.namingconventionmapi(message).trim());
		if (f2.exists()) {
			f2 = new File(f1.getAbsolutePath() + File.separator + mm.namingconventionmapi(message).trim() + "_"
					+ count_destination);
		}
		f2.mkdir();
		for (MapiAttachment attachment : message.getAttachments()) {
			try {
				attachment.save(f2.getAbsolutePath() + File.separator + attachment.getDisplayName().trim());
			} catch (Exception e) {
				attachment.save(f2.getAbsolutePath() + File.separator + attachment.getLongFileName().trim());
			}
		}
	}

}
