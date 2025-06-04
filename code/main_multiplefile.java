package email.code;

import java.awt.AWTException;
import java.awt.AlphaComposite;
import java.awt.CardLayout;
import java.awt.Color;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.Font;
import java.awt.GradientPaint;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.HeadlessException;
import java.awt.MenuItem;
import java.awt.PopupMenu;
import java.awt.SystemTray;
import java.awt.Toolkit;
import java.awt.TrayIcon;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.math.BigInteger;
import java.net.InetAddress;
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.net.URLConnection;
import java.net.UnknownHostException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Enumeration;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TimeZone;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.Action;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JEditorPane;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JPopupMenu;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JSpinner;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.RowFilter;
import javax.swing.SpinnerModel;
import javax.swing.SpinnerNumberModel;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.WindowConstants;
import javax.swing.border.EmptyBorder;
import javax.swing.border.LineBorder;
import javax.swing.border.TitledBorder;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import javax.swing.table.TableRowSorter;
import javax.swing.text.DefaultEditorKit;
import javax.swing.text.html.HTMLEditorKit;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreeNode;
import javax.swing.tree.TreePath;

import com.aspose.email.AppointmentSaveFormat;
import com.aspose.email.Attachment;
import com.aspose.email.ContactSaveFormat;
import com.aspose.email.EWSClient;
import com.aspose.email.EmailClient;
import com.aspose.email.EmlSaveOptions;
import com.aspose.email.ExchangeMessageInfo;
import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.IConnection;
import com.aspose.email.IEWSClient;
import com.aspose.email.ImapClient;
import com.aspose.email.ImapMessageFlags;
import com.aspose.email.ImapMessageInfo;
import com.aspose.email.MailAddress;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiAttachmentCollection;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiContactElectronicAddress;
import com.aspose.email.MapiContactEventPropertySet;
import com.aspose.email.MapiContactNamePropertySet;
import com.aspose.email.MapiContactPersonalInfoPropertySet;
import com.aspose.email.MapiContactPhysicalAddress;
import com.aspose.email.MapiContactPhysicalAddressPropertySet;
import com.aspose.email.MapiContactProfessionalPropertySet;
import com.aspose.email.MapiContactTelephonePropertySet;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiElectronicAddress;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiMessageFlags;
import com.aspose.email.MapiTask;
import com.aspose.email.MapiTaskUsers;
import com.aspose.email.MboxrdStorageReader;
import com.aspose.email.MboxrdStorageWriter;
import com.aspose.email.MessageInfo;
import com.aspose.email.MessageInfoCollection;
import com.aspose.email.OAuthNetworkCredential;
import com.aspose.email.OlmFolder;
import com.aspose.email.OlmStorage;
import com.aspose.email.PersonalStorage;
import com.aspose.email.SaveOptions;
import com.aspose.email.SecurityOptions;
import com.aspose.email.system.NetworkCredential;
import com.aspose.email.system.exceptions.FileNotFoundException;
import com.aspose.email.system.io.FileAccess;
import com.aspose.email.system.io.FileMode;
import com.aspose.email.system.io.FileStream;
import com.opencsv.CSVWriter;
import com.toedter.calendar.JDateChooser;
import com.toedter.calendar.JTextFieldDateEditor;

import components.GradientButton;
import components.Label;
import email.activation.ActivationFrame;
import email.activation.OnlineActivation;
import email.activation.Starting_Frame;
import email.activation.Uninstall;
import email.design.CustomTreeNode;
import it.cnr.imaa.essi.lablib.gui.checkboxtree.CheckboxTree;
import it.cnr.imaa.essi.lablib.gui.checkboxtree.DefaultCheckboxTreeCellRenderer;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.ExchangeService;

import microsoft.exchange.webservices.data.property.definition.ExtendedPropertyDefinition;
import microsoft.exchange.webservices.data.property.definition.ExtendedPropertyDefinition.*;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import email.activation.OnlineActivation;
import email.activation.Starting_Frame;
import email.design.CustomTreeNode;
import it.cnr.imaa.essi.lablib.gui.checkboxtree.CheckboxTree;
import it.cnr.imaa.essi.lablib.gui.checkboxtree.DefaultCheckboxTreeCellRenderer;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.SendInvitationsMode;
import microsoft.exchange.webservices.data.core.enumeration.service.TaskStatus;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.exception.service.remote.ServiceRequestException;
import microsoft.exchange.webservices.data.core.service.folder.ContactsFolder;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Contact;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.recurrence.pattern.Recurrence;
import microsoft.exchange.webservices.data.property.complex.time.TimeZoneDefinition;
import microsoft.exchange.webservices.data.core.enumeration.property.*;
import microsoft.exchange.webservices.data.property.complex.*;
//import microsoft.exchange.webservices.data.core.enumeration.property.MapiPropertyType.*;
import microsoft.exchange.webservices.data.core.enumeration.property.MapiPropertyType.*;
import microsoft.exchange.webservices.data.core.service.item.*;

import javax.swing.border.EtchedBorder;
import javax.swing.JRadioButton;
import javax.swing.ButtonGroup;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.UIManager;

@SuppressWarnings("deprecation")
public class main_multiplefile extends JFrame {

	boolean fir = false;
	boolean second = false;
	boolean third = false;
	boolean fourth = false;
	boolean fifth = false;

	String x1 = "";
	String parent1 = "";
	String sepreter = "";
	JButton btnActivate;
	JButton btn_buy;
	JCheckBox chckbxShowPassword_p3;
	JLabel lblNewLabel;
	JLabel lblNewLabel_1;
	LoadingDialog loadingDialog;
	List<String> listdupliccal = new ArrayList<String>();
	List<String> listduplictask = new ArrayList<String>();
	List<String> listdupliccontact = new ArrayList<String>();
	List<String> listduplicacy = new ArrayList<String>();
	public static main_multiplefile main_m;
	static public String projectTitle;
	static File licFileon;
	static boolean datevalidflag = false;
	boolean checkDate = false;
	static Date fromdate;
	 int modelRow;
	static Date todate;
	ArrayList<Date> fromList = new ArrayList<Date>();
	ArrayList<Date> toList = new ArrayList<Date>();
	String from;
	String to;
	JTextFieldDateEditor fromdateeditor;
	JTextFieldDateEditor todateeditor;
	static File listOfFile;
	GradientButton btn_converter_1;
	JPopupMenu jPopupMenu;
	String first = null, middle = null, last = null;
	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	private static JTable table;
	static long count_destination;
	static JFileChooser jFileChooser;
	static File[] files;
	long maxsize = 0;
	int pstindex = 0;
	String pstfilename = "";
	int ijj = 1;
	int splitcount = 0;
	String splitpath = "";
	@SuppressWarnings("rawtypes")
	JComboBox comboBox_setsize;
	JCheckBox chckbx_splitpst;
	LoadingThreadclass obTh;
	JSpinner spinner_sizespinner;
	JCheckBox chckbx_seperatepst;
	Cursor cursor = new Cursor(Cursor.HAND_CURSOR);
	JPanel panel_8;
	long foldercountcheck = 0;
	JCheckBox chckbx_convert_pdf_to_pdf;
	JLabel label_16;
	JLabel lblNewLabel_8;
	JLabel lblNewLabel_9;
	Date c1;
	JCheckBox chckbxSavePdfAttachment;
	JLabel label_pdf_to_pdf;
	String destination = "";
	Date pdf_date;
	JCheckBox chckbxSaveMboxIn;
	int ids = 1;
	double SPACE_KB = 1024;
	double SPACE_MB = 1024 * SPACE_KB;
	double SPACE_GB = 1024 * SPACE_MB;
	double SPACE_TB = 1024 * SPACE_GB;
	LoadingDialog ld;
	Boolean checkmboxpstost = true;
	JCheckBox chckbxRestoreToDefault;
	JCheckBox chckbxMaintainFolderStructure;
	JPanel panel_Contact;
	JPanel panel_Callendar;
	String fileoption;
	Boolean contatcheck;
	Boolean calendarcheck;
	JLabel lblFullName;
	JLabel lblEMailAd;
	JLabel lblPhoneNo;
	JLabel label_contactcompany;
	JLabel llabel_contactfullname;
	JLabel lblCompany;
	JLabel label_contactemail;
	String OS = System.getProperty("os.name").toLowerCase();
	JLabel label_contactfullname;
	JLabel label_contactphonenumber;
	JLabel label_contacticon;
	JLabel label_calendarstartdate;
	JLabel lblEndDate;
	JLabel label_Calendarsubject;
	JLabel label_Calendaricon;
	JTextArea textArea_contact;
	boolean folder_check = false;
	static DefaultTableModel modeli;
	static JComboBox<String> comboBox_FiletypeChooser;
	static String fileoptionm;
	File f;
	JCheckBox chckbxRemoveDuplicacy;
	JPanel panel_3_1_2;
	public static JLabel label_11;
	JLabel lblMakeSureYou;
	JLabel lblEnableImap_p3;
	JLabel lblTurnOffTwo_p3;
	JComboBox<String> comboBox;
	File filetem;
	JButton btn_signout_p3;
	JButton updateBtn;
	Calendar cal;
	JLabel lbl_splitpst;
	JPanel panel_progress;
	JPanel panel_taskfilter;
	long count_eml_msg_emlx = 0;
	int portnofiletype;
	File file;
	static int countforfile = 0;
	static PersonalStorage pst;
	PersonalStorage ost;
	Main_Frame mf;
	List<String> pstfolderlist;
	ArrayList<String> pstfolderlist2;
	static List<DefaultMutableTreeNode> lists = new ArrayList<DefaultMutableTreeNode>();
	static List<String> listst = new ArrayList<String>();
	static Date mailfilterstartdate;
	static Date mailfilterenddate;
	static Date Calenderfilterstartdate;
	static Date Calenderfilterenddate;
	static Date taskfilterenddate;
	static Date taskfilterstartdate;
	GradientButton btn_previous_p2;
	GradientButton btn_remove;
	Boolean output = false;
	JLabel lbl_Domain;
	HashMap<String, List<String>> hm;
	String path4 = "";
	static JRadioButton radioFileFormat;
	static String fa = "";
	JLabel lbl_connecting_p3;
	JLabel lbl_progressreport;
	FolderInfo info = new FolderInfo();
	FolderInfo folderInfo;
	static DefaultMutableTreeNode root;
	static DefaultTreeModel model;
	DefaultMutableTreeNode mainnode;
	static DefaultMutableTreeNode lastNode;
	static DefaultTableModel mode;
	GradientButton btn_select_folder;
	JButton btn_previous_p3;
	JButton btn_Destination;
	JButton btn_cancel;
	JButton btnStop;
//	private JButton btn_converter_1;
	GradientButton btn_Next;
	JButton btnDowloadReport;
	private JButton btnDowloadReport_1;
	String domain_p3;
	static String username_p3;
	static String password_p3;
	String filetype;
	String calendertime;
	String parentfolder;
	String subfolderfile;
	String reportpath;
	String filepath;

	String Status = "Completed";
	static String fname;

	String filetemp = "";
	String foldername = "";
	String foldername2 = "";
	String foldername3 = "";
	String foldername4 = "";
	String foldername5 = "";
	String foldername6 = "";
	String subfolder = "";
	String subfolder2 = "";
	String subfolder3 = "";
	String subfolder4 = "";
	String subfolder5 = "";
	String subfolder6 = "";
	static String path = "";
	String path2;
	String path3 = "";
	String parentname;
	String path1 = "";
	String s = "";
	String parent = "";
	String[] filesfin;
	String Folder;
	String Folderuri;
	String destination_path;
	static String mailboxUri = "https://outlook.office365.com/EWS/Exchange.asmx";
	String buyurl;
	String infourl;
	String helpurl;
	MboxrdStorageWriter wr;
	OlmStorage storage;
	int filesno = 1;
	int folder = 0;
	CSVWriter writer;
	boolean connectioncheck = true;
	JPanel panel_3;
	JPanel panel_3_;
	JPanel panel_3_1;
	JPanel panel_3_1_1;
	JLabel lblLive_Chat_p3;
	JPanel panel_3_2;
	JPanel innercardlayout;
	JPanel Cardlayout;
	boolean cropted;
	JTextField textField_username_p3;
	GradientButton btnNewButton_2;
	JTextField tf_Destination_Location;
	JPasswordField passwordField_p3;
	Thread th;
	static JRadioButton rdbtnEmailClients;
	static IEWSClient clientforexchange_output;
	static ImapClient clientforimap_output;
	static IConnection iconnforimap_output;
	JProgressBar progressBar_message_p3;
	JTable table_fileConvertionreport_panel4;
	JComboBox<String> comboBox_fileDestination_type;
//	JButton btn_next_pane2;
	Calendar Cal;
	static Boolean demo = true;
	Boolean btnfile = false;
	Boolean btnfolder = false;
	Boolean stop = false;
	Boolean pstCalenderfoldercheck = true;
	Boolean pstContactfoldercheck = true;
	Boolean psttaskfoldercheck = true;
	Boolean foldercheck = true;
	Boolean checkconvertagain = false;
	Boolean checkdestination = true;
	Boolean Stoppreview = false;
	int ret;
	static CheckboxTree tree;
	OlmFolder folderi;
	long foldermessagecount;
	FolderInfo info1;
	JDateChooser dateChooser_calender_start;
	JDateChooser dateChooser_mail_fromdate;
	JDateChooser dateChooser_task_end_date;
	JDateChooser dateChooser_task_start_date;
	JDateChooser dateChooser_mail_tilldate;
	JDateChooser dateChooser_calendar_end;
	GradientButton btnremove_all;
	Boolean checky = true;
	JCheckBox chckbx_Mail_Filter;
	JCheckBox chckbx_calender_box;
	private JPanel panel_4;
	private JTable table_fileinformation;
	private JPanel attachmenttable;
	private JEditorPane editorPane;
	private JScrollPane scrollPane_3;
	private JTable table_1;
	private JPanel panel_3_1_2_1;
	private JTextField textField_domain_name_p3;
	Set<File> hashset = new LinkedHashSet<File>();
	private JScrollPane scrollPane_4;
	JCheckBox chckbxCustomFolderName;
	JCheckBox task_box;
	private List<MailMessage> listmail = new ArrayList<MailMessage>();
	private List<MapiMessage> listmapi = new ArrayList<MapiMessage>();
	private List<ImapMessageInfo> listImapmesinfo = new ArrayList<ImapMessageInfo>();
	private List<ExchangeMessageInfo> listExchangemesingo = new ArrayList<ExchangeMessageInfo>();
	private List<MessageInfo> listPSTOSTgemesingo = new ArrayList<MessageInfo>();
	private JLabel lblLoadingPleaseWait;
	private JLabel label_10;
	String logpathm = "";
	String temppathm = "";
	static JLabel Progressbar;
	private JLabel lblPortNo;
	private JTextField tf_portNo_p3;
	private JPanel panel;
	private JPanel panel_5;
	private JLabel lblnamingconvention;
	private JLabel lblnamingconvention_1;
	private JPanel panel_6;
	JTextField textField_customfolder;
	private JLabel label_Calendarenddate;
	private JLabel lblNotes;
	private JPanel panel_7;
	private JLabel lblNewLabel_5;
	private JLabel lblemailAddress;
	String version;
	private JLabel lblTotalMessageCount;
	private JCheckBox chckbxSaveInSame;
	private JLabel label_12;
	private JLabel label_13;
	private JLabel label_14;
	private JLabel label_15;
	JCheckBox chckbxMigrateOrBackup;
	private JLabel label_17;
	private JPanel date_filter;
	private JCheckBox checkBox;
	static JCheckBox datefilter;
	private JLabel lblNewLabel_6;
	private JDateChooser dateChooser_NewFrom;
	GradientButton btn_Next_1 ;
	private JDateChooser dateChooser_NewTo;
	private JTable table_2;
	JButton add;
	JButton remove;
	private GradientButton btn_ChoseFile;
	public static String messageboxtitle = All_Data.messageboxtitle;
	public static String match;
	public static String parent_path;
	private JRadioButton basic_Authentication;
	static JRadioButton modern_Authentication;
	private final ButtonGroup buttonGroup = new ButtonGroup();
	private JButton techHelp;
	private JButton btnNewButton_1;
	private JPanel panel_10;
	private JPanel panel_12;
	private JPanel panel_13;
	private final ButtonGroup buttonGroup_1 = new ButtonGroup();
///ews
	public static HashMap<String, FolderId> map = new HashMap<String, FolderId>();
	public static FolderId rootfolderid;
	public static FolderId folderid;
	EWSOffice ows;
	static ExchangeService service;
	public static boolean except = false;
	static EWSOffice ews;
	static String filterselected = "mailboxsel";
	static long count_destination_total = 0;
	public static JRadioButton mailbox;
	public static JRadioButton publicfolder;
	public static JRadioButton archive;
	public static boolean firstFolderGeneralContact = false;
	public static boolean firstFolderGeneralAppointment = false;
	public static boolean firstFolderGeneralTask = false;
	public static HashMap<String, FolderId> generalFolderMap = new HashMap<String, FolderId>();
	public static String locdat;
	private GradientButton btn_next_pane2;
	private JLabel lblNewLabel_4;
	private JTextField searchField;
	private TableRowSorter<TableModel> rowSorter;

	@SuppressWarnings({ "rawtypes", "unchecked", "resource" }) // @SuppressWarnings()
	public main_multiplefile(JFrame parent1, Boolean demo, String messageboxtitle) {
		super();
		main_m = main_multiplefile.this;
		mf = (Main_Frame) parent1;
		main_multiplefile.demo = demo;
//		main_multiplefile.messageboxtitle = All_Data.messageboxtitle;
		// fileoptionm = Main_Frame.fileoption;

		fileoptionm = All_Data.input_default;
		this.fileoption = All_Data.input_default;
		logpathm = mf.logpath;
		version = mf.version;
		temppathm = mf.temppath;
		calendertime = Main_Frame.calendertime;
		buyurl = mf.buyurl;
		infourl = mf.infourl;
		helpurl = mf.helpurl;

		addWindowListener(new WindowAdapter() {

			public void windowClosing(WindowEvent arg0) {
				String warn = "Do you want to close the Application?";
				int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {
					setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
					System.exit(0);
					main_multiplefile.this.dispose();
				}
			}

		});

		setIconImage(Toolkit.getDefaultToolkit().getImage(Main_Frame.class.getResource("/128x128.png")));

		if (demo) {
			String center = messageboxtitle;
			setTitle(center);
		} else {
			String center = messageboxtitle;
			setTitle(center);
		}

		setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		setLocationRelativeTo(null);
		setResizable(true);
		setBounds(100, 100, 1190, 718);
		contentPane = new JPanel();
		contentPane.setBackground(Color.WHITE);
		contentPane.setBorder(null);
		setContentPane(contentPane);

		jPopupMenu = new JPopupMenu();
		jPopupMenu.setBackground(Color.WHITE);
		Action cut = new DefaultEditorKit.CutAction();
		cut.putValue(Action.NAME, "Cut");
		cut.putValue(Action.ACCELERATOR_KEY, KeyStroke.getKeyStroke("control X"));
		jPopupMenu.add(cut);

		Action copy = new DefaultEditorKit.CopyAction();
		copy.putValue(Action.NAME, "Copy");
		copy.putValue(Action.ACCELERATOR_KEY, KeyStroke.getKeyStroke("control C"));
		jPopupMenu.add(copy);

		Action paste = new DefaultEditorKit.PasteAction();
		paste.putValue(Action.NAME, "Paste");
		paste.putValue(Action.ACCELERATOR_KEY, KeyStroke.getKeyStroke("control V"));

		jPopupMenu.add(paste);

		Action selectAll = new SelectAll();
		jPopupMenu.add(selectAll);

		Cardlayout = new JPanel();
		Cardlayout.setBounds(198, 73, 973, 606);
		Cardlayout.setBackground(Color.LIGHT_GRAY);
		Cardlayout.setLayout(new CardLayout(0, 0));

		JPanel panel_1 = new JPanel();
		panel_1.setBackground(Color.WHITE);
		Cardlayout.add(panel_1, "panel_1");
		btnActivate = new JButton("");
		btnActivate.setBounds(98, 430, 80, 80);
		
		
		btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/key-act-btn.png")));
		if (demo) {
			btnActivate.setVisible(true);
		} else {
			btnActivate.setVisible(true);

		}
		if (demo) {
			btnActivate.setToolTipText("Click here to activate the software.");
			btnActivate.addMouseListener(new MouseAdapter() {
				public void mouseEntered(MouseEvent arg0) {
					btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/key-act-btn_hrv.png")));
				}

				public void mouseExited(MouseEvent e) {
					btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/key-act-btn.png")));
				}
			});
			btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/key-act-btn.png")));
		} else {
			btnActivate.setToolTipText("Click here to deactivate the software.");
//			btnActivate.addMouseListener(new MouseAdapter() {
//				public void mouseEntered(MouseEvent arg0) {
//					btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/deactivate-hvr-btn.png")));
//				}
//
//				public void mouseExited(MouseEvent e) {
//					btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/deactivate-btn.png")));
//				}
//			});
			btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/deactivate-btn.png")));

		}
		btn_buy = new JButton("");
		btn_buy.setBounds(101, 338, 80, 80);
		if (demo) {
			btn_buy.setVisible(true);
		} else {
			btn_buy.setVisible(false);
		}

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(0, 0, 973, 545);

		table = new JTable() {

			@Override
			public void paintComponent(Graphics g) {
				try {
					super.paintComponent(g);

					((Graphics2D) g)
							.setPaint(new GradientPaint(0f, 0f, Color.WHITE, getWidth(), getHeight(), Color.ORANGE));

					((Graphics2D) g).setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER).derive(0.3f));
					g.fillRect(0, 0, getWidth(), getHeight());
				} catch (Exception e) {
					// e.printStackTrace();
				}
			}

			/**
			* 
			*/
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {

				return false;
			}
		};
		table.setFont(new Font("Tahoma", Font.PLAIN, 11));
		scrollPane.setViewportView(table);
		table.getTableHeader().setReorderingAllowed(false);
		table.setModel(new DefaultTableModel(new Object[][] {},
				new String[] { "<html><b>" + "File no.", "<html><b>" + " PST File Name",
						"<html><b>" + "Selected File Path", "<html><b>" + "<html><b>" + "File Type",
						"<html><b>" + "Size Of File" }));
		table.getColumnModel().getColumn(0).setPreferredWidth(5);
		table.getColumnModel().getColumn(1).setPreferredWidth(150);
		table.getColumnModel().getColumn(2).setPreferredWidth(330);
		table.getColumnModel().getColumn(3).setPreferredWidth(20);
		table.getColumnModel().getColumn(4).setPreferredWidth(30);
		//
		// TableColumnModel tcm = table.getColumnModel();
		table.getColumnModel().getColumn(3).setMinWidth(0);
		table.getColumnModel().getColumn(3).setMaxWidth(0);
		table.getColumnModel().getColumn(3).setWidth(0);
		Stoppreview = false;
		panel_1.setLayout(null);

		btn_ChoseFile = new GradientButton("select-files");
		btn_ChoseFile.setText("Select File(s)");
		btn_ChoseFile.setGradientColor1(new Color(70, 130, 180));
		btn_ChoseFile.setGradientColor2(new Color(70, 130, 180));
		btn_ChoseFile.setForeground(new Color(255, 255, 255));

		btn_ChoseFile.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				btn_ChoseFile.setGradientColor1(new Color(255, 255, 255));
				btn_ChoseFile.setGradientColor2(new Color(255, 255, 255));
				btn_ChoseFile.setForeground(new Color(80, 80, 80));

			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_ChoseFile.setGradientColor1(new Color(70, 130, 180));
				btn_ChoseFile.setGradientColor2(new Color(70, 130, 180));
				btn_ChoseFile.setForeground(new Color(255, 255, 255));
			}
		});

		btn_ChoseFile.setShadowColor(new Color(0, 0, 205));
		btn_ChoseFile.setBounds(79, 557, 150, 40);
		panel_1.add(btn_ChoseFile);
		btn_ChoseFile.setToolTipText("Click here to Select the File(s).");
		btn_ChoseFile.setRolloverEnabled(false);
		btn_ChoseFile.setRequestFocusEnabled(false);
		btn_ChoseFile.setOpaque(false);
		btn_ChoseFile.setFocusable(false);
		btn_ChoseFile.setFocusTraversalKeysEnabled(false);
		btn_ChoseFile.setFocusPainted(false);
		btn_ChoseFile.setDefaultCapable(false);
		btn_ChoseFile.setContentAreaFilled(false);
		btn_ChoseFile.setBorderPainted(false);
		btn_ChoseFile.setBackground(Color.WHITE);
		btn_ChoseFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				try {
					filter_file();
					btnfile = true;
					// btn_Next.setEnabled(true);
					// btn_Next.setVisible(true);
					btn_remove.setEnabled(true);
					btnremove_all.setEnabled(true);

					btn_previous_p2.setEnabled(true);
					btn_Next.setEnabled(true);

					if (countforfile == 0) {

						btnfile = false;
						// btn_Next.setEnabled(false);
						// btn_Next.setVisible(false);
						btn_remove.setEnabled(false);
						btnremove_all.setEnabled(false);

						btn_previous_p2.setEnabled(false);
						btn_Next.setEnabled(false);
					}
				} catch (Exception e) {

				}

			}
		});

		btn_ChoseFile.setFont(new Font("Verdana", Font.BOLD, 12));

		btn_select_folder = new GradientButton("Select Folder(s)");
		btn_select_folder.setGradientColor1(new Color(70, 130, 180));
		btn_select_folder.setGradientColor2(new Color(70, 130, 180));
		btn_select_folder.setForeground(new Color(255, 255, 255));

		btn_select_folder.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				btn_select_folder.setGradientColor1(new Color(255, 255, 255));
				btn_select_folder.setGradientColor2(new Color(255, 255, 255));
				btn_select_folder.setForeground(new Color(80, 80, 80));

			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_select_folder.setGradientColor1(new Color(70, 130, 180));
				btn_select_folder.setGradientColor2(new Color(70, 130, 180));
				btn_select_folder.setForeground(new Color(255, 255, 255));
			}
		});

		btn_select_folder.setShadowColor(new Color(0, 0, 128));
		btn_select_folder.setBounds(422, 557, 150, 40);
		panel_1.add(btn_select_folder);
		btn_select_folder.setToolTipText("Click here to Select the Folder(s).");
		btn_select_folder.setOpaque(false);
		btn_select_folder.setFocusable(false);
		btn_select_folder.setFocusTraversalKeysEnabled(false);
		btn_select_folder.setFocusPainted(false);
		btn_select_folder.setDefaultCapable(false);
		btn_select_folder.setContentAreaFilled(false);
		btn_select_folder.setBorderPainted(false);
		btn_select_folder.setRolloverEnabled(false);
		btn_select_folder.setRequestFocusEnabled(false);
		btn_select_folder.setFont(new Font("Tahoma", Font.BOLD, 12));

		btn_select_folder.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				files = null;
				btnfolder = true;

				btn_Next.setVisible(true);
				btn_Next.setEnabled(true);

				btn_remove.setEnabled(true);
				btnremove_all.setEnabled(true);

				btn_previous_p2.setEnabled(true);
				btn_Next.setEnabled(true);

				jFileChooser = new JFileChooser();

				jFileChooser.setMultiSelectionEnabled(true);

				jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

				int r = jFileChooser.showOpenDialog(main_multiplefile.this);
				if ((r == JFileChooser.APPROVE_OPTION)) {
					listOfFile = jFileChooser.getSelectedFile();
					try {
						files = listOfFile.listFiles();

						lisstOfFiles(files);
					} catch (Exception e) {

						JOptionPane.showMessageDialog(main_multiplefile.this, "Please select correct Folder.",
								messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));
						return;

					}
				}

				DefaultTableModel model = (DefaultTableModel) table.getModel();

				while (model.getRowCount() > 0) {

					for (int i = 0; i < model.getRowCount(); ++i) {

						model.removeRow(i);
						filesno--;
					}
				}

				Iterator<File> itr = hashset.iterator();
				while (itr.hasNext()) {

					modeli = (DefaultTableModel) table.getModel();
					File fo = itr.next();

					String filet = "";
					if (fo.isFile()) {
						filet = "File";
					} else {
						filet = "Folder";
					}
					long sizeInBytes = fo.length();
					modeli.addRow(new Object[] { filesno, fo.getName(), fo.getAbsolutePath(), filet,
							bytes2String(sizeInBytes) });
					filesno++;
					countforfile++;
				}
				if (countforfile == 0) {

					btn_Next.setEnabled(false);
					btn_Next.setVisible(false);
					btn_remove.setEnabled(false);
					btnremove_all.setEnabled(false);

					btn_previous_p2.setEnabled(false);
					btn_Next.setEnabled(false);

				}

			}

		});

		btn_remove = new GradientButton("Remove File");
		btn_remove.setText("Remove Selected File");
		btn_remove.setGradientColor1(new Color(70, 130, 180));
		btn_remove.setGradientColor2(new Color(70, 130, 180));
		btn_remove.setForeground(new Color(255, 255, 255));
		btn_remove.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				btn_remove.setGradientColor1(new Color(255, 255, 255));
				btn_remove.setGradientColor2(new Color(255, 255, 255));
				btn_remove.setForeground(new Color(80, 80, 80));

			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_remove.setGradientColor1(new Color(70, 130, 180));
				btn_remove.setGradientColor2(new Color(70, 130, 180));
				btn_remove.setForeground(new Color(255, 255, 255));
			}
		});

		btn_remove.setShadowColor(new Color(0, 0, 128));
		btn_remove.setBounds(746, 557, 192, 40);
		panel_1.add(btn_remove);
		btn_remove.setToolTipText("Click here to Remove Selected File(s).");
		btn_remove.setRolloverEnabled(false);
		btn_remove.setRequestFocusEnabled(false);
		btn_remove.setOpaque(false);
		btn_remove.setFocusable(false);
		btn_remove.setFocusTraversalKeysEnabled(false);
		btn_remove.setFocusPainted(false);
		btn_remove.setDefaultCapable(false);
		btn_remove.setContentAreaFilled(false);
		btn_remove.setBorderPainted(false);
		btn_remove.setEnabled(false);
		btn_remove.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				try {
					DefaultTableModel model = (DefaultTableModel) table.getModel();
					int selected = table.getSelectedRow();

					hashset.remove(new File(table.getValueAt(selected, 2).toString().replace("<html><b>", "")));

					model.removeRow(selected);
					countforfile--;
					filesno--;
					if (countforfile == 0) {

						btn_Next.setEnabled(false);
						// btn_Next.setVisible(false);
						btn_remove.setEnabled(false);
						btnremove_all.setEnabled(false);

						btn_previous_p2.setEnabled(false);
						btn_Next.setEnabled(false);
						btn_next_pane2.setEnabled(false);
						btn_converter_1.setEnabled(false);
						btnNewButton_2.setEnabled(false);

					}
				} catch (Exception e) {
					JOptionPane.showMessageDialog(mf, "Please select  file or Folder you want to remove",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));
				}

			}
		});
		btn_remove.setFont(new Font("Tahoma", Font.BOLD, 12));

		btnremove_all = new GradientButton("Remove-All");
		btnremove_all.setGradientColor1(new Color(70, 130, 180));
		btnremove_all.setGradientColor2(new Color(70, 130, 180));
		btnremove_all.setForeground(new Color(255, 255, 255));
		btnremove_all.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				btnremove_all.setGradientColor1(new Color(255, 255, 255));
				btnremove_all.setGradientColor2(new Color(255, 255, 255));
				btnremove_all.setForeground(new Color(80, 80, 80));

			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnremove_all.setGradientColor1(new Color(70, 130, 180));
				btnremove_all.setGradientColor2(new Color(70, 130, 180));
				btnremove_all.setForeground(new Color(255, 255, 255));
			}
		});

		btnremove_all.setShadowColor(new Color(0, 0, 128));
		btnremove_all.setBounds(790, 557, 150, 40);
		// panel_1.add(btnremove_all);
		btnremove_all.setToolTipText("Click here to Remove All Files.");
		btnremove_all.setRolloverEnabled(false);
		btnremove_all.setRequestFocusEnabled(false);
		btnremove_all.setOpaque(false);
		btnremove_all.setFocusable(false);
		btnremove_all.setFocusTraversalKeysEnabled(false);
		btnremove_all.setFocusPainted(false);
		btnremove_all.setDefaultCapable(false);
		btnremove_all.setContentAreaFilled(false);
		btnremove_all.setBorderPainted(false);
		btnremove_all.setEnabled(false);
		btnremove_all.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				DefaultTableModel model = (DefaultTableModel) table.getModel();
				while (model.getRowCount() > 0) {

					for (int i = 0; i < model.getRowCount(); ++i) {

						model.removeRow(i);

					}
				}
				hashset.clear();
				filesno = 1;
				countforfile = 0;
				btn_Next.setEnabled(false);
				btn_Next.setVisible(false);
				btn_remove.setEnabled(false);
				btnremove_all.setEnabled(false);

				btn_previous_p2.setEnabled(false);
				btn_Next.setEnabled(false);
				btn_next_pane2.setEnabled(false);
				btn_converter_1.setEnabled(false);
				btnNewButton_2.setEnabled(false);

			}

		});
		btnremove_all.setFont(new Font("Tahoma", Font.BOLD, 12));
		panel_1.add(scrollPane);

		JPanel panel_2 = new JPanel();
		panel_2.setBackground(Color.WHITE);
		Cardlayout.add(panel_2, "panel_2");

		JButton btnViewer = new JButton("");
		btnViewer.setBounds(625, 233, 111, 31);
		btnViewer.setToolTipText("Click here to Preview");
		btnViewer.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnViewer.setIcon(new ImageIcon(Main_Frame.class.getResource("/preview-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				btnViewer.setIcon(new ImageIcon(Main_Frame.class.getResource("/preview-btn.png")));
			}
		});

		btnViewer.setIcon(new ImageIcon(Main_Frame.class.getResource("/preview-btn.png")));
		btnViewer.setRolloverEnabled(false);
		btnViewer.setRequestFocusEnabled(false);
		btnViewer.setOpaque(false);
		btnViewer.setFocusable(false);
		btnViewer.setFocusTraversalKeysEnabled(false);
		btnViewer.setFocusPainted(false);
		btnViewer.setDefaultCapable(false);
		btnViewer.setContentAreaFilled(false);
		btnViewer.setBorderPainted(false);
		btnViewer.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (contatcheck) {

					CardLayout card1 = (CardLayout) innercardlayout.getLayout();
					card1.show(innercardlayout, "panel_Contact");

				} else if (calendarcheck) {
					CardLayout card1 = (CardLayout) innercardlayout.getLayout();
					card1.show(innercardlayout, "panel_Callendar");
				} else {
					CardLayout card = (CardLayout) innercardlayout.getLayout();

					card.show(innercardlayout, "viewer");
				}

			}
		});

		label_10 = new JLabel("");
		label_10.setBounds(82, 569, 668, 31);
		label_10.setRequestFocusEnabled(false);
		label_10.setFocusable(false);
		label_10.setFocusTraversalKeysEnabled(false);
		label_10.setIcon(new ImageIcon(Main_Frame.class.getResource("/progress-bar.gif")));
		label_10.setVisible(false);
		btnViewer.setFont(new Font("Tahoma", Font.BOLD, 13));

		JButton btnAttachment = new JButton("");
		btnAttachment.setBounds(736, 233, 111, 31);
		btnAttachment.setToolTipText("Click here to view Attachment");
		btnAttachment.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnAttachment.setIcon(new ImageIcon(Main_Frame.class.getResource("/attachment-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				btnAttachment.setIcon(new ImageIcon(Main_Frame.class.getResource("/attachment-btn.png")));
			}
		});

		btnAttachment.setIcon(new ImageIcon(Main_Frame.class.getResource("/attachment-btn.png")));
		btnAttachment.setDefaultCapable(false);
		btnAttachment.setFocusTraversalKeysEnabled(false);
		btnAttachment.setRolloverEnabled(false);
		btnAttachment.setRequestFocusEnabled(false);
		btnAttachment.setOpaque(false);
		btnAttachment.setFocusable(false);
		btnAttachment.setFocusPainted(false);
		btnAttachment.setContentAreaFilled(false);
		btnAttachment.setBorderPainted(false);
		btnAttachment.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				CardLayout card = (CardLayout) innercardlayout.getLayout();
				card.show(innercardlayout, "attachmenttable");

			}
		});
		btnAttachment.setFont(new Font("Tahoma", Font.BOLD, 13));

		JScrollPane scrollPane_1 = new JScrollPane();
		scrollPane_1.setBounds(0, 0, 295, 606);

		tree = new CheckboxTree() {
			@Override
			public void paintComponent(Graphics g) {
				try {
					super.paintComponent(g);

					((Graphics2D) g)
							.setPaint(new GradientPaint(0f, 0f, Color.WHITE, getWidth(), getHeight(), Color.ORANGE));

					((Graphics2D) g).setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER).derive(0.3f));
					g.fillRect(0, 0, getWidth(), getHeight());
				} catch (Exception e) {
					// e.printStackTrace();
				}
			}

		};
		scrollPane_1.setViewportView(tree);
		tree.addMouseListener(new MouseAdapter() {

			public void mouseClicked(MouseEvent arg0) {
				th = new Thread(new Runnable() {

					public void run() {

						try {
							listmail.clear();
							listmapi.clear();
							btn_cancel.setVisible(true);
							btn_cancel.setEnabled(true);
							listExchangemesingo.clear();
							listImapmesinfo.clear();
							Stoppreview = false;
							listPSTOSTgemesingo.clear();
							btnAttachment.setEnabled(false);
							btnViewer.setEnabled(false);
							btn_next_pane2.setEnabled(false);
							btn_previous_p2.setEnabled(false);
							lblTotalMessageCount.setText("<html><b>" + "  Total Message in this Folder : ");
							CardLayout card = (CardLayout) innercardlayout.getLayout();
							listmail.clear();
							card.show(innercardlayout, "viewer");
							editorPane.setText("");
							if (arg0.getClickCount() == 2) {

								loadingDialog = new LoadingDialog(mf, true);
								loadingDialog.setLocationRelativeTo(mf);

								Thread innerthread = new Thread(new Runnable() {
									public void run() {
										loadingDialog.setVisible(true);
									}
								});
								innerthread.start();

								TreePath tp = tree.getSelectionPath();

								DefaultMutableTreeNode node = (DefaultMutableTreeNode) tp.getLastPathComponent();

								foldername = node.getUserObject().toString();

								DefaultTableModel model = (DefaultTableModel) table_fileinformation.getModel();

								while (model.getRowCount() > 0) {

									for (int i = 0; i < model.getRowCount(); ++i) {

										model.removeRow(i);
									}
								}

								DefaultTableModel model1 = (DefaultTableModel) table_1.getModel();

								while (model1.getRowCount() > 0) {

									for (int i = 0; i < model1.getRowCount(); ++i) {

										model1.removeRow(i);
									}
								}
								editorPane.setText("");

//								lblNew_setemail.setText("");
//
//								lblNew_setsubject.setText("");
//
//								label_date.setText("");

								if (fileoptionm.equalsIgnoreCase("MBOX")) {

									path = ((CustomTreeNode) tp.getPathComponent(3)).filepath;

									if (node.isLeaf()) {
										fileInformation_on_mbox();
									}

								} else if (fileoptionm.equalsIgnoreCase("OLM File (.olm)")) {
									foldername = "";
									TreeNode[] folder = node.getPath();

									for (int i = 2; i < folder.length; i++) {
										String s = folder[i].toString().trim();

										// System.out.println(s);
										s = s.replace("<html><b>", "");
										if (i == 2) {
											path2 = ((CustomTreeNode) tp.getPathComponent(i)).filepath;
										} else if (i == 3) {
											foldername = s;

										} else if (i > 3) {
											foldername = foldername + File.separator + s;
										}

									}

									if (!foldername.equalsIgnoreCase(root.toString())) {
										fileinformation_olm();
									}

								} else if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")
										|| fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
									foldername = "";

									TreeNode[] folder = node.getPath();

									for (int i = 2; i < folder.length; i++) {
										String s = folder[i].toString().trim();

										s = s.replace("<html><b>", "");
										if (i == 2) {
											path2 = ((CustomTreeNode) tp.getPathComponent(i)).filepath;
										} else if (i == 3) {
											foldername = s;

										} else if (i > 3) {
											foldername = foldername + File.separator + s;
										}

									}
									if (!foldername.equalsIgnoreCase(root.toString())) {

										try {
											pst = PersonalStorage.fromFile(path2);
											fileInhformation_Ost_Pst();

										} catch (Exception e) {
											JOptionPane.showMessageDialog(mf,
													"Selected File is Currupted  Please Choose another file  "
															+ filepath,
													messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
													new ImageIcon(Main_Frame.class.getResource("/information.png")));

										}

									}

								} else {

									path2 = ((CustomTreeNode) tp.getLastPathComponent()).filepath;

									file = new File(path2);
									ids = 1;
									if (file.isFile()) {
										fileInformation_on_mail();
									} else if (file.isDirectory()) {
										readewwds(file);
									}

								}
							}
							lblLoadingPleaseWait.setVisible(false);
							label_10.setVisible(false);
							btn_cancel.setVisible(false);
							btn_cancel.setEnabled(false);
							btnAttachment.setEnabled(true);
							table_fileinformation.setEnabled(true);
							btnViewer.setEnabled(true);
							btn_next_pane2.setEnabled(true);
							btn_previous_p2.setEnabled(true);

						} catch (Exception e) {

						} finally {

							if (loadingDialog != null) {

								loadingDialog.dispose();
							}

						}

					}

				});
				th.start();

			}
		});

		tree.setModel(new DefaultTreeModel(new DefaultMutableTreeNode("root") {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;

			{

			}
		}));

		JScrollPane scrollPane_2 = new JScrollPane();
		scrollPane_2.setBounds(298, 0, 675, 237);

		innercardlayout = new JPanel();
		innercardlayout.setBounds(298, 265, 675, 341);
		innercardlayout.setLayout(new CardLayout(0, 0));

		CardLayout card = (CardLayout) innercardlayout.getLayout();
		card.show(innercardlayout, "viewer");

		table_fileinformation = new JTable() {

			@Override
			public void paintComponent(Graphics g) {
				try {
					super.paintComponent(g);

					((Graphics2D) g)
							.setPaint(new GradientPaint(0f, 0f, Color.WHITE, getWidth(), getHeight(), Color.ORANGE));

					((Graphics2D) g).setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER).derive(0.3f));
					g.fillRect(0, 0, getWidth(), getHeight());
				} catch (Exception e) {
					// e.printStackTrace();
				}
			}

			/**
			* 
			*/
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {

				return false;
			}
		};

		table_fileinformation.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent arg0) {
				try {
					SwingUtilities.invokeLater(new Runnable() {

						public void run() {
							if (arg0.getClickCount() == 1) {
								editorPane.setText("");
								DefaultTableModel model = (DefaultTableModel) table_1.getModel();
								while (model.getRowCount() > 0) {

									for (int i = 0; i < model.getRowCount(); ++i) {

										model.removeRow(i);
									}
								}
								contatcheck = false;
								calendarcheck = false;
								CardLayout card = (CardLayout) innercardlayout.getLayout();

								card.show(innercardlayout, "viewer");
								if (fileoptionm.equalsIgnoreCase("MBOX")) {
									MailMessage message = listmail.get(table_fileinformation.getSelectedRow());

									try {
//										lblNew_setemail.setText(message.getFrom().toString());
//
//										lblNew_setsubject.setText(message.getSubject());
//
//										label_date.setText(message.getDate().toString());

										HTMLEditorKit kit = new HTMLEditorKit();
										editorPane.setEditorKit(kit);
										FileOutputStream os = new FileOutputStream(
												temppathm + File.separator + "previewHtml.html");
										message.save(os, EmlSaveOptions.getDefaultHtml());
										os.close();
										URL url = new URL("file:///" + temppathm + File.separator + "previewHtml.html");
										editorPane.setPage(url);

										for (int j = 0; j < message.getAttachments().size(); j++) {
											Attachment att = message.getAttachments().get_Item(j);

											String attFileName = att.getName();
											ImageIcon icon = null;

											if (attFileName.endsWith(".pdf")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
											} else if (attFileName.endsWith(".txt")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
											} else if (attFileName.endsWith(".docx")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
											} else if (attFileName.endsWith(".zip")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
											} else {
												icon = new ImageIcon(
														Main_Frame.class.getResource("/attachment-icon.png"));
											}
											JLabel imagelabl = new JLabel();
											imagelabl.setIcon(icon);
											DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
											modeli.addRow(new Object[] { "<html><b>" + (j + 1),
													"<html><b>" + attFileName, imagelabl });
										}

									} catch (Exception e) {
									}

								} else if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
										|| fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

									MapiMessage message = listmapi.get(modelRow);

									  System.out.println("this is selected row  "  +  modelRow);
									
									//MapiMessage message = listmapi.get(table_fileinformation.getSelectedRow());
									MailConversionOptions de = new MailConversionOptions();
									MailMessage mess = message.toMailMessage(de);
//									try {
//										lblNew_setemail.setText(message.getSenderEmailAddress());
//									} catch (Exception a) {
//										lblNew_setemail.setText("");
//									}
//									try {
//										lblNew_setsubject.setText(message.getSubject());
//									} catch (Exception a) {
//										lblNew_setsubject.setText("");
//									}
//									try {
//										label_date.setText(message.getDeliveryTime().toString());
//									} catch (Exception a) {
//										label_date.setText("");
//									}

//									if (message.getMessageClass().equals("IPM.Task")
//											|| message.getMessageClass().equals("IPM.StickyNote")
//											|| message.getMessageClass().equals("IPM.Contact")
//											|| message.getMessageClass().equals("IPM.Appointment")
//											|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {
//										int bct = message.getBodyType();
//										if (bct == 0) {
//											message.setBody(mess.getBody());
//										} else {
//											message.setBody(mess.getBody());
//										}
//									}

									if (message.getMessageClass().equals("IPM.Contact")) {
										CardLayout card1 = (CardLayout) innercardlayout.getLayout();
										card1.show(innercardlayout, "panel_Contact");

										MapiContact con = (MapiContact) message.toMapiMessageItem();
										try {
//											String[] compa = con.getCompanies();
//											label_contactcompany.setText(compa[0]);
											label_contactcompany.setText(con.getProfessionalInfo().getCompanyName());
										} catch (Exception e) {
											label_contactcompany.setText("");
										}
										try {
											String gn = "";
											String mn = "";
											String sn = "";
											try {
												gn = con.getNameInfo().getGivenName();
											} catch (Exception e) {

											}
											if (gn == null) {
												gn = "";
											}
											try {
												mn = con.getNameInfo().getMiddleName();
											} catch (Exception e) {

											}
											if (mn == null) {
												mn = "";
											}
											try {
												sn = con.getNameInfo().getSurname();
											} catch (Exception e) {

											}
											if (sn == null) {
												sn = "";
											}

											String fn = gn + " " + mn + " " + sn;
											label_contactfullname.setText(fn);
										} catch (Exception e) {
											label_contactfullname.setText("");
										}
										try {
											label_contactemail.setText(
													con.getElectronicAddresses().getEmail1().getEmailAddress());
										} catch (Exception e) {
											label_contactemail.setText("");
										}
										try {
											label_contactphonenumber
													.setText(con.getTelephones().getMobileTelephoneNumber());
										} catch (Exception e) {
											label_contactphonenumber.setText("");
										}
										try {
//											textArea_contact.setText(con.getPersonalInfo().getNotes());
											textArea_contact.setText(message.getBody());
											if (message.getBody() == null) {
												textArea_contact.setText(con.getPersonalInfo().getNotes().toString());
											}
										} catch (Exception e) {
											textArea_contact.setText("");
										}

										contatcheck = true;

									} else if (message.getMessageClass().equals("IPM.Appointment")
											|| message.getMessageClass().contains("IPM.Schedule.Meeting") && !message
													.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
										CardLayout card1 = (CardLayout) innercardlayout.getLayout();
										card1.show(innercardlayout, "panel_Callendar");
										System.out.println("check");
										MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

										calendarcheck = true;
										try {
											label_Calendarsubject.setText(cal.getSubject());
										} catch (Exception e) {
											label_Calendarsubject.setText("");
										}
										try {
											label_calendarstartdate.setText(cal.getStartDate().toString());
										} catch (Exception e) {
											label_calendarstartdate.setText("");
										}
										try {
											label_Calendarenddate.setText(cal.getEndDate().toString());
										} catch (Exception e) {
											label_Calendarenddate.setText("");
										}

									} else {

										try {

											HTMLEditorKit kit = new HTMLEditorKit();
											editorPane.setEditorKit(kit);
											FileOutputStream os = new FileOutputStream(
													temppathm + File.separator + "previewHtml.html");
											message.save(os, EmlSaveOptions.getDefaultHtml());
											os.close();
											URL url = new URL(
													"file:///" + temppathm + File.separator + "previewHtml.html");
											editorPane.setPage(url);

										} catch (Error e) {
											mf.logger.warning("ERROR : " + e.getMessage() + System.lineSeparator());
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
											editorPane.setContentType("text/html");
											editorPane.setText("<html>Page not found.</html>");
										}

									}

									int k = 1;
									for (int j = 0; j < message.getAttachments().size(); j++) {
										MapiAttachment att = message.getAttachments().get_Item(j);

										String attFileName = "";
										try {
											attFileName = att.getDisplayName();
										} catch (Exception e) {
											attFileName = att.getLongFileName();
										}
										ImageIcon icon = null;

										if (attFileName.endsWith(".pdf")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
										} else if (attFileName.endsWith(".txt")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
										} else if (attFileName.endsWith(".docx")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
										} else if (attFileName.endsWith(".zip")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
										} else {
											icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
										}
										JLabel imagelabl = new JLabel();
										imagelabl.setIcon(icon);

										DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
										modeli.addRow(
												new Object[] { "<html><b>" + k, "<html><b>" + attFileName, imagelabl });
										k++;
										// System.out.println(attFileName);

									}

								} else if (fileoptionm.equalsIgnoreCase("OLM File (.olm)")) {

									MapiMessage message = listmapi.get(table_fileinformation.getSelectedRow());
//									try {
//										lblNew_setemail.setText(message.getSenderEmailAddress());
//									} catch (Exception a) {
//										lblNew_setemail.setText("");
//									}
//									try {
//										lblNew_setsubject.setText(message.getSubject());
//									} catch (Exception a) {
//										lblNew_setsubject.setText("");
//									}
//									try {
//										label_date.setText(message.getDeliveryTime().toString());
//									} catch (Exception a) {
//										label_date.setText("");
//									}
//
//									try {
//										lblNew_setemail.setText(message.getSenderEmailAddress());
//									} catch (Exception a) {
//										lblNew_setemail.setText("");
//									}
//									try {
//										lblNew_setsubject.setText(message.getSubject());
//									} catch (Exception a) {
//										lblNew_setsubject.setText("");
//									}
//									try {
//										label_date.setText(message.getDeliveryTime().toString());
//									} catch (Exception a) {
//										label_date.setText("");
//									}

									if (message.getMessageClass().equals("IPM.Contact")) {
										CardLayout card1 = (CardLayout) innercardlayout.getLayout();
										card1.show(innercardlayout, "panel_Contact");

										MapiContact con = (MapiContact) message.toMapiMessageItem();
										try {
											String[] compa = con.getCompanies();

											label_contactcompany.setText(compa[0]);
										} catch (Exception e) {
											label_contactcompany.setText("");
										}
										try {
											String gn = "";
											String mn = "";
											String sn = "";
											try {
												gn = con.getNameInfo().getGivenName();
											} catch (Exception e) {

											}
											try {
												mn = con.getNameInfo().getMiddleName();
											} catch (Exception e) {

											}
											try {
												sn = con.getNameInfo().getSurname();
											} catch (Exception e) {

											}

											String fn = gn + " " + mn + " " + sn;
											label_contactfullname.setText(fn);
										} catch (Exception e) {
											label_contactfullname.setText("");
										}
										try {
											label_contactemail.setText(
													con.getElectronicAddresses().getEmail1().getEmailAddress());
										} catch (Exception e) {
											label_contactemail.setText("");
										}
										try {
											label_contactphonenumber
													.setText(con.getTelephones().getMobileTelephoneNumber());
										} catch (Exception e) {
											label_contactphonenumber.setText("");
										}
										try {
											textArea_contact.setText(con.getPersonalInfo().getNotes());
										} catch (Exception e) {
											textArea_contact.setText("");
										}

										contatcheck = true;

									} else if (message.getMessageClass().equals("IPM.Appointment")
											|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {
										CardLayout card1 = (CardLayout) innercardlayout.getLayout();
										card1.show(innercardlayout, "panel_Callendar");
										System.out.println("check");
										MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

										calendarcheck = true;
										try {
											label_Calendarsubject.setText(cal.getSubject());
										} catch (Exception e) {
											label_Calendarsubject.setText("");
										}
										try {
											label_calendarstartdate.setText(cal.getStartDate().toString());
										} catch (Exception e) {
											label_calendarstartdate.setText("");
										}
										try {
											label_Calendarenddate.setText(cal.getEndDate().toString());
										} catch (Exception e) {
											label_Calendarenddate.setText("");
										}

									} else {

										try {

											HTMLEditorKit kit = new HTMLEditorKit();
											editorPane.setEditorKit(kit);
											FileOutputStream os = new FileOutputStream(
													temppathm + File.separator + "previewHtml.html");
											message.save(os, EmlSaveOptions.getDefaultHtml());
											os.close();
											URL url = new URL(
													"file:///" + temppathm + File.separator + "previewHtml.html");
											editorPane.setPage(url);

										} catch (Error e) {
											mf.logger.warning("ERROR : " + e.getMessage() + System.lineSeparator());
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
											editorPane.setContentType("text/html");
											editorPane.setText("<html>Page not found.</html>");
										}

									}

									int k = 1;
									for (int j = 0; j < message.getAttachments().size(); j++) {
										MapiAttachment att = message.getAttachments().get_Item(j);

										String attFileName = att.getLongFileName();
										ImageIcon icon = null;

										if (attFileName.endsWith(".pdf")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
										} else if (attFileName.endsWith(".txt")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
										} else if (attFileName.endsWith(".docx")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
										} else if (attFileName.endsWith(".zip")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
										} else {
											icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
										}
										JLabel imagelabl = new JLabel();
										imagelabl.setIcon(icon);

										DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
										modeli.addRow(
												new Object[] { "<html><b>" + k, "<html><b>" + attFileName, imagelabl });
										k++;
										// System.out.println(attFileName);

									}
								} else if (fileoptionm.equalsIgnoreCase("Nodes Storage (.nsf)")) {

								} else {
									MailMessage message = listmail.get(table_fileinformation.getSelectedRow());

									// System.out.println("found");

									try {
//										try {
//											lblNew_setemail.setText(message.getFrom().toString());
//										} catch (Exception e) {
//
//										}
//										try {
//											lblNew_setsubject.setText(message.getSubject());
//										} catch (Exception e) {
//
//										}
//										try {
//											label_date.setText(message.getDate().toString());
//										} catch (Exception e) {
//
//										}
										HTMLEditorKit kit = new HTMLEditorKit();
										editorPane.setEditorKit(kit);
										FileOutputStream os = new FileOutputStream(
												temppathm + File.separator + "previewHtml.html");
										message.save(os, EmlSaveOptions.getDefaultHtml());
										os.close();
										URL url = new URL("file:///" + temppathm + File.separator + "previewHtml.html");
										editorPane.setPage(url);

										for (int j = 0; j < message.getAttachments().size(); j++) {
											Attachment att = message.getAttachments().get_Item(j);

											String attFileName = att.getName();
											ImageIcon icon = null;

											if (attFileName.endsWith(".pdf")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
											} else if (attFileName.endsWith(".txt")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
											} else if (attFileName.endsWith(".docx")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
											} else if (attFileName.endsWith(".zip")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
											} else {
												icon = new ImageIcon(
														Main_Frame.class.getResource("/attachment-icon.png"));
											}
											JLabel imagelabl = new JLabel();
											imagelabl.setIcon(icon);

											DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
											modeli.addRow(new Object[] { "<html><b>" + (j + 1),
													"<html><b>" + attFileName, imagelabl });
											// System.out.println(attFileName);
										}
									} catch (Exception e) {

										// e.printStackTrace();
									}
								}
								
								 
								

							}
						}
					});

				} catch (Exception e) {

				}
			}
		});
		table_fileinformation.setModel(
				new DefaultTableModel(new Object[][] {}, new String[] { "<html><b>" + "From", "<html><b>" + "Subject",
						"<html><b>" + "Date", "<html><b>" + "Attachment count", "<html><b>" + "Attachment" }));
		table_fileinformation.getColumn("<html><b>" + "Attachment").setCellRenderer(new Renderer());
		table_fileinformation.setColumnSelectionAllowed(false);
		table_fileinformation.setRowSelectionAllowed(true);
		DefaultTableCellRenderer tablerenderer = (DefaultTableCellRenderer) table_fileinformation.getTableHeader()
				.getDefaultRenderer();
		tablerenderer.setHorizontalAlignment(0);
		table_fileinformation.getTableHeader().setReorderingAllowed(false);

		table_fileinformation.getColumnModel().getColumn(0).setPreferredWidth(200);
		table_fileinformation.getColumnModel().getColumn(1).setPreferredWidth(200);
		table_fileinformation.getColumnModel().getColumn(2).setPreferredWidth(150);
		table_fileinformation.getColumnModel().getColumn(3).setPreferredWidth(50);
		table_fileinformation.getColumnModel().getColumn(4).setPreferredWidth(50);

		scrollPane_2.setViewportView(table_fileinformation);

		JPanel viewer = new JPanel();
		innercardlayout.add(viewer, "viewer");
		viewer.setLayout(null);

		scrollPane_3 = new JScrollPane();
		scrollPane_3.setBounds(0, 0, 675, 371);

		editorPane = new JEditorPane();
		scrollPane_3.setViewportView(editorPane);
		editorPane.setEditable(false);
		viewer.add(scrollPane_3);

		panel_Callendar = new JPanel();
		panel_Callendar.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_Callendar.setBackground(Color.WHITE);
		innercardlayout.add(panel_Callendar, "panel_Callendar");
		panel_Callendar.setLayout(null);

		label_calendarstartdate = new JLabel("");
		label_calendarstartdate.setBounds(78, 151, 351, 17);
		panel_Callendar.add(label_calendarstartdate);

		lblEndDate = new JLabel("End Date");
		lblEndDate.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblEndDate.setBounds(10, 199, 65, 17);
		panel_Callendar.add(lblEndDate);

		JLabel lblStartDate = new JLabel("Start Date");
		lblStartDate.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblStartDate.setBounds(10, 148, 65, 22);
		panel_Callendar.add(lblStartDate);

		JLabel lblSubject = new JLabel("Subject");
		lblSubject.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblSubject.setBounds(10, 92, 52, 28);
		panel_Callendar.add(lblSubject);

		label_Calendarsubject = new JLabel("");
		label_Calendarsubject.setBounds(85, 92, 344, 28);
		panel_Callendar.add(label_Calendarsubject);

		label_Calendaricon = new JLabel("");
		label_Calendaricon.setIcon(new ImageIcon(Main_Frame.class.getResource("/calender.png")));
		label_Calendaricon.setBounds(10, 11, 116, 77);
		panel_Callendar.add(label_Calendaricon);

		label_Calendarenddate = new JLabel("");
		label_Calendarenddate.setBounds(78, 199, 351, 22);
		panel_Callendar.add(label_Calendarenddate);

		panel_Contact = new JPanel();
		panel_Contact.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_Contact.setBackground(Color.WHITE);
		innercardlayout.add(panel_Contact, "panel_Contact");
		panel_Contact.setLayout(null);

		lblFullName = new JLabel("Full Name ");
		lblFullName.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblFullName.setBounds(10, 67, 64, 14);
		panel_Contact.add(lblFullName);

		label_contactfullname = new JLabel("");
		label_contactfullname.setBounds(83, 67, 215, 14);
		panel_Contact.add(label_contactfullname);

		lblEMailAd = new JLabel("E Mail :");
		lblEMailAd.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblEMailAd.setBounds(10, 92, 64, 21);
		panel_Contact.add(lblEMailAd);

		lblCompany = new JLabel("Company");
		lblCompany.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblCompany.setBounds(10, 124, 64, 14);
		panel_Contact.add(lblCompany);

		label_contactemail = new JLabel("");
		label_contactemail.setBounds(83, 95, 215, 18);
		panel_Contact.add(label_contactemail);

		label_contactcompany = new JLabel("");
		label_contactcompany.setBounds(83, 124, 215, 21);
		panel_Contact.add(label_contactcompany);

		lblPhoneNo = new JLabel("Phone No");
		lblPhoneNo.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblPhoneNo.setBounds(10, 149, 56, 21);
		panel_Contact.add(lblPhoneNo);

		label_contactphonenumber = new JLabel("");
		label_contactphonenumber.setBounds(80, 152, 218, 18);
		panel_Contact.add(label_contactphonenumber);

		textArea_contact = new JTextArea();
		textArea_contact.setEditable(false);
		textArea_contact.setBounds(10, 198, 655, 139);
		panel_Contact.add(textArea_contact);

		label_contacticon = new JLabel("");
		label_contacticon.setIcon(new ImageIcon(Main_Frame.class.getResource("/User-Chat-icon.png")));
		label_contacticon.setBounds(10, 11, 64, 51);
		panel_Contact.add(label_contacticon);

		lblNotes = new JLabel("Notes");
		lblNotes.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNotes.setBounds(10, 172, 48, 21);
		panel_Contact.add(lblNotes);

		attachmenttable = new JPanel();
		innercardlayout.add(attachmenttable, "attachmenttable");

		scrollPane_4 = new JScrollPane();
		scrollPane_4.setBounds(0, 0, 675, 348);

		Object[][] data1 = {};

		String[] cols1 = { "S.No", "File Name", "File Type" };

		DefaultTableModel tablemodel = new DefaultTableModel(data1, cols1) {

			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {

				return false;

			}

		};
		attachmenttable.setLayout(null);

		table_1 = new JTable(tablemodel);
		scrollPane_4.setViewportView(table_1);
		attachmenttable.add(scrollPane_4);

		lblTotalMessageCount = new JLabel("<html><b>" + "Total Message Count :");
		lblTotalMessageCount.setBounds(301, 238, 190, 23);
		lblTotalMessageCount.setForeground(UIManager.getColor("CheckBox.foreground"));

//		btn_next_pane2 = new JButton("");
//		btn_next_pane2.setBounds(169, 21, 111, 31);
		btn_next_pane2 = new GradientButton("Filters");
		btn_next_pane2.setEnabled(false);
		btn_next_pane2.setToolTipText("Click here to Go Forward.");
		btn_next_pane2.setDefaultCapable(false);
		btn_next_pane2.setContentAreaFilled(false);
		btn_next_pane2.setBorderPainted(false);
		btn_next_pane2.setRequestFocusEnabled(false);
		btn_next_pane2.setOpaque(false);
		btn_next_pane2.setRolloverEnabled(false);
		btn_next_pane2.setFocusable(false);
		btn_next_pane2.setFocusTraversalKeysEnabled(false);
		btn_next_pane2.setFocusPainted(false);

		btn_next_pane2.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btn_next_pane2.isEnabled()) {
					btn_next_pane2.setGradientColor1(new Color(70, 130, 180));
					btn_next_pane2.setGradientColor2(new Color(70, 130, 180));
					btn_next_pane2.setForeground(new Color(255, 255, 255));
				}
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_next_pane2.setGradientColor1(new Color(255, 255, 255));
				btn_next_pane2.setGradientColor2(new Color(255, 255, 255));
				btn_next_pane2.setForeground(new Color(80, 80, 80));
			}
		});

		btn_next_pane2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				
				filetype=GetSelectedDest();
				
				
				
				third = true;

				Buttonclick("btn_next_pane2", third);

				radioFileFormat.setSelected(true);
				radioFileFormat.doClick();
				Stoppreview = true;
				TreePath[] tp = tree.getCheckingPaths();

				TreePath[] checktp11 = tree.getCheckingPaths();
//				if (checktp11.length == 0) {
//					JOptionPane.showMessageDialog(mf, "Please Select File From the Tree", messageboxtitle,
//							JOptionPane.ERROR_MESSAGE,
//							new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
//				} else {

				hm = new HashMap<String, List<String>>();

				pstfolderlist = null;
				destination_path = "";
				for (int i = 0; i < tp.length; i++) {

					pstfolderlist2 = new ArrayList<String>();
					String[] str = (tp[i].toString().replace("<html><b>", "")).split(",");

					String sfile = "";
					StringBuilder strbr = new StringBuilder();
					for (int j = 3; j < str.length; j++) {

						if (j != (str.length - 1)) {
							strbr.append(str[j].trim());
							if (!pstfolderlist2.contains(strbr.toString().trim())) {
								pstfolderlist2.add(strbr.toString().trim());
							}
							if (fileoptionm.equalsIgnoreCase("Maildir")) {
								strbr.append(",");

							} else {
								strbr.append(File.separator);
							}

						} else if (j == str.length - 1) {
							strbr.append(str[j].replace("]", "").trim());
							if (!pstfolderlist2.contains(strbr.toString().trim())) {
								pstfolderlist2.add(strbr.toString().trim());
							}
						}

						if (fileoptionm.equalsIgnoreCase("EML File (.eml)")
								|| fileoption.equalsIgnoreCase("EMLX File (.emlx)")
								|| fileoption.equalsIgnoreCase("OFT File (.oft)")
								|| fileoptionm.equalsIgnoreCase("Message File (.msg)")
								|| fileoptionm.equalsIgnoreCase("Maildir")) {

							sfile = ((CustomTreeNode) tp[i].getLastPathComponent()).filepath;

						}

					}
					if (fileoptionm.equalsIgnoreCase("EML File (.eml)")
							|| fileoption.equalsIgnoreCase("EMLX File (.emlx)")
							|| fileoption.equalsIgnoreCase("OFT File (.oft)")
							|| fileoptionm.equalsIgnoreCase("Message File (.msg)")
							|| fileoptionm.equalsIgnoreCase("Maildir")) {
						if (new File(sfile).isFile()) {
							hm.put(sfile, null);
						} else {

							addemldd(new File(sfile));
						}

					}

					if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
							|| fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")
							|| fileoptionm.equalsIgnoreCase("OLM File (.olm)")) {
						String fname;
						try {
							fname = ((CustomTreeNode) tp[i].getPathComponent(2)).filepath;
						} catch (Exception e33) {
							continue;
						}

						if (hm.size() != 0) {
							pstfolderlist = hm.get(fname);
							if (pstfolderlist != null) {
								if (!pstfolderlist.contains(strbr.toString().trim())) {
									pstfolderlist.add(strbr.toString().trim());
								}
							} else {
								pstfolderlist = new ArrayList<String>();

								pstfolderlist.addAll(pstfolderlist2);
								if (!pstfolderlist.contains(strbr.toString().trim())) {
									pstfolderlist.add(strbr.toString().trim());
								}
								hm.put(fname, pstfolderlist);

							}

						} else {
							pstfolderlist = new ArrayList<String>();
							pstfolderlist.add(strbr.toString().trim());
							hm.put(fname, pstfolderlist);
						}
					} else {
						DefaultMutableTreeNode d1 = (DefaultMutableTreeNode) tp[i].getLastPathComponent();
						if (d1.isLeaf()) {

							if (filetype.equalsIgnoreCase("Maildir")) {
								sfile = str[2].trim();

								sfile = sfile + pstfolderlist2.get(3);

							} else {
								if (fileoptionm.equalsIgnoreCase("EML File (.eml)")
										|| fileoption.equalsIgnoreCase("EMLX File (.emlx)")
										|| fileoption.equalsIgnoreCase("OFT File (.oft)")
										|| fileoptionm.equalsIgnoreCase("Message File (.msg)")) {
									sfile = str[2].replace("]", "").trim();

								} else {
									sfile = ((CustomTreeNode) tp[i].getPathComponent(3)).filepath;

								}
							}
							if (!(fileoptionm.equalsIgnoreCase("EML File (.eml)")
									|| fileoption.equalsIgnoreCase("EMLX File (.emlx)")
									|| fileoption.equalsIgnoreCase("OFT File (.oft)")
									|| fileoptionm.equalsIgnoreCase("Message File (.msg)"))) {

								hm.put(sfile, null);

							}

						}
					}

				}
				if (fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")
						|| fileoption.equalsIgnoreCase("OLM File (.olm)")
						|| fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
					try {

						try {

							for (int i = 0; i < Main_Frame.file_sfd.length; i++) {

								mf.imageMap_output.put(Main_Frame.file_sfd[i],
										new ImageIcon(Main_Frame.class.getResource(Main_Frame.filesfd_img[i])));

							}
						} catch (Exception ex) {
							ex.printStackTrace();
						}

						// comboBox_fileDestination_type.addItem("VCF");
						// comboBox_fileDestination_type.addItem("ICS");
						//
						// mf.imageMap_output.put("VCF", new
						// ImageIcon(Main_Frame.class.getResource("/vcf.png")));
						// mf.imageMap_output.put("ICS", new
						// ImageIcon(Main_Frame.class.getResource("/ics.png")));

					} catch (Exception e1) {

					}

				} else {

					try {
						comboBox_fileDestination_type.removeItem("VCF");

					} catch (Exception e1) {
						e1.printStackTrace();
					}
					try {
						comboBox_fileDestination_type.removeItem("ICS");

					} catch (Exception e1) {
						e1.printStackTrace();
					}

				}
				panel_progress.setVisible(false);

				if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

					comboBox_fileDestination_type.removeItem("OST");
				} else if (fileoptionm.equalsIgnoreCase("EML File (.eml)")) {

					comboBox_fileDestination_type.removeItem("EML");
				} else if (fileoptionm.equalsIgnoreCase("EMLX File (.emlx)")) {

					comboBox_fileDestination_type.removeItem("EMLX");

				} else if (fileoptionm.equalsIgnoreCase("Message File (.msg)")) {

					comboBox_fileDestination_type.removeItem("MSG");

				} else if (fileoptionm.equalsIgnoreCase("MBOX")) {
					comboBox_fileDestination_type.removeItem("MBOX");
					comboBox_fileDestination_type.removeItem("Thunderbird");
					comboBox_fileDestination_type.removeItem("Opera Mail");
				}
				if (fileoptionm.equalsIgnoreCase("EML File (.eml)") || fileoptionm.equalsIgnoreCase("Thunderbird")
						|| fileoptionm.equalsIgnoreCase("Opera Mail")
						|| fileoption.equalsIgnoreCase("EMLX File (.emlx)")
						|| fileoption.equalsIgnoreCase("OFT File (.oft)")
						|| fileoptionm.equalsIgnoreCase("Message File (.msg)")) {

					task_box.setVisible(false);
					panel_taskfilter.setVisible(false);

				}
				panel_3_2.setVisible(false);
				panel_3_.setVisible(false);
				panel_3_1_2.setVisible(false);
				btn_converter_1.setEnabled(false);
				comboBox.setVisible(false);
				btnStop.setVisible(false);

				tf_Destination_Location.setText(System.getProperty("user.home") + File.separator + "Desktop");

				lbl_progressreport.setText("");

				panel_3_.setVisible(true);

				CardLayout card1 = (CardLayout) panel_3_.getLayout();
				card1.show(panel_3_, "panel_3_1_1");
				panel_3_2.setVisible(true);
//																			panel_progress.setVisible(true);
				filetype = "PST";
				if (fileoptionm.equalsIgnoreCase("MBOX")) {
					label_16.setVisible(true);
					chckbxSaveMboxIn.setVisible(true);
				}
				btn_converter_1.setEnabled(true);
				comboBox_fileDestination_type.setSelectedItem("PST");
				if (Main_Frame.versiontype == 1) {
					rdbtnEmailClients.setVisible(false);
				} else {
					rdbtnEmailClients.setVisible(true);
				}
				
				
				
				
				//selecteddestination():
				
				
				
				
				
				
				
				
				
				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_3");
				cal = Calendar.getInstance();
				calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());

				chckbxMaintainFolderStructure.setSelected(true);
				// }

			}
		});
		btn_next_pane2.setFont(new Font("Tahoma", Font.BOLD, 12));
		table_1.getColumn("File Type").setCellRenderer(new Renderer());

		panel_3 = new JPanel();
		panel_3.setBackground(Color.WHITE);
		Cardlayout.add(panel_3, "panel_3");
//		panel_3.setLayout(null);

		textField_customfolder = new JTextField();
		panel_3_ = new JPanel();
		panel_3_2 = new JPanel();
		panel_3_1_2 = new JPanel();
		lbl_splitpst = new JLabel("");
		chckbxCustomFolderName = new JCheckBox("Custom Folder Name");
		chckbx_splitpst = new JCheckBox("Split Resultant PST\r\n");
		comboBox_fileDestination_type = new JComboBox(Main_Frame.file_sfd);
		panel_3_1_2_1 = new JPanel();
		chckbxSavePdfAttachment = new JCheckBox("Save attachment separately");
		label_15 = new JLabel("");
		textField_domain_name_p3 = new JTextField();
		tf_Destination_Location = new JTextField();
		textField_username_p3 = new JTextField();
		chckbxSaveInSame = new JCheckBox("Save in Same(Source and Destination Folder are same)");

		comboBox_fileDestination_type.setBounds(509, 11, 310, 29);

		mf.imageMap_output = mf.createImageMap_output(mf.l_output);
		comboBox_fileDestination_type.setRenderer(mf.new ListRenderer_output());

		comboBox_fileDestination_type.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				try {
					chckbxCustomFolderName.setSelected(false);
					if (chckbxCustomFolderName.isSelected()) {
						textField_customfolder.setEnabled(true);
						textField_customfolder.setEditable(true);
					} else {

						textField_customfolder.setText("");
						textField_customfolder.setEnabled(false);
						textField_customfolder.setEditable(false);
					}

					panel_3_.setVisible(false);

					panel_3_2.setVisible(false);

					panel_3_1_2.setVisible(false);

					panel_3_1_2.setVisible(false);
					lbl_splitpst.setVisible(false);
					chckbx_splitpst.setSelected(false);
					chckbx_splitpst.setVisible(false);
					panel_3_1_2_1.setVisible(false);
					chckbxSavePdfAttachment.setVisible(false);
					label_15.setVisible(false);
					textField_domain_name_p3.setText("");
					output = false;
					tf_Destination_Location.setText(System.getProperty("user.home") + File.separator + "Desktop");
					textField_username_p3.setText("");
					chckbxSaveInSame.setSelected(false);
					try {
						passwordField_p3.setText("");
						lbl_progressreport.setText("");
						btn_converter_1.setEnabled(false);
						lbl_splitpst.setVisible(false);
						chckbx_splitpst.setSelected(false);
						panel_progress.setVisible(true);
						tf_portNo_p3.setVisible(false);
						lblPortNo.setVisible(false);
						comboBox.setVisible(false);
						btn_signout_p3.setVisible(false);
						chckbx_convert_pdf_to_pdf.setVisible(false);
						label_pdf_to_pdf.setVisible(false);
						lblMakeSureYou.setVisible(true);
						lblEnableImap_p3.setVisible(true);
						lblTurnOffTwo_p3.setVisible(true);
						label_16.setVisible(false);
						chckbxSaveMboxIn.setVisible(false);
						chckbxRestoreToDefault.setVisible(false);
						panel_5.setVisible(false);
						panel_8.setVisible(false);
						mailbox.setVisible(false);
						archive.setVisible(false);
						publicfolder.setVisible(false);
						if (arg0.getSource() == comboBox_fileDestination_type) {

							JComboBox cb = (JComboBox) arg0.getSource();

							filetype = (String) cb.getSelectedItem();
							if (cb.getSelectedItem() == null) {
								return;
							}
						}
					} catch (Exception e) {
						System.out.println("here we riched 2582");
						// TODO Auto-generated catch block
						// e.printStackTrace();
					}

//if(!cb.getSelectedItem()==null) {
					try {
						if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")
								|| filetype.equalsIgnoreCase("YAHOO MAIL") || filetype.equalsIgnoreCase("Icloud")
								|| filetype.equalsIgnoreCase("GoDaddy email")
								|| filetype.equalsIgnoreCase("Hostgator email")
								|| filetype.equalsIgnoreCase("Amazon WorkMail")
								|| filetype.equalsIgnoreCase("OFFICE 365") || filetype.equalsIgnoreCase("AOL")
								|| filetype.equalsIgnoreCase("Live Exchange")
								|| filetype.equalsIgnoreCase("Yandex Mail") || filetype.equalsIgnoreCase("Zoho Mail")
								|| filetype.equalsIgnoreCase("HOTMAIL") || filetype.equalsIgnoreCase("IMAP")) {
							lblEnableImap_p3.setText("<HTML><U>To Enable IMAP</U><HTML>");
							lbl_splitpst.setVisible(false);
							chckbx_splitpst.setSelected(false);
							chckbxSaveInSame.setVisible(false);
							label_13.setVisible(false);
							lblEnableImap_p3.setVisible(false);
							lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
									+ " account , you'll need to generate and use an app password.</U></HTML>");
							lblMakeSureYou.setText("Please  Click on The Link");
							lblNewLabel_5.setVisible(true);

							if (!filetype.equalsIgnoreCase("OFFICE 365")) {
								lblNewLabel_1.setVisible(true);
								passwordField_p3.setVisible(true);
								chckbxShowPassword_p3.setVisible(true);
							}

							if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("OFFICE 365")
									|| filetype.equalsIgnoreCase("G-SUITE"))) {
								basic_Authentication.setSelected(true);
								System.out.println("here we riched");
								textField_username_p3.setEnabled(true);
								passwordField_p3.setEnabled(true);
								tf_portNo_p3.setEnabled(true);
								chckbxShowPassword_p3.setEnabled(true);
								lblPortNo.setEnabled(true);
								lblNewLabel_5.setEnabled(true);
								lblNewLabel_1.setEnabled(true);
								lblNewLabel.setEnabled(true);
								lblemailAddress.setEnabled(true);

							}
							if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")) {
								modern_Authentication.setVisible(true);
								basic_Authentication.setVisible(true);
								basic_Authentication.setEnabled(true);
								basic_Authentication.setSelected(true);
								textField_username_p3.setEnabled(true);
								passwordField_p3.setEnabled(true);
								chckbxShowPassword_p3.setEnabled(true);
								lblNewLabel.setEnabled(true);
								lblemailAddress.setEnabled(true);
								lblNewLabel_1.setEnabled(true);
								lblNewLabel_5.setEnabled(true);

							} else {
								modern_Authentication.setVisible(false);
								basic_Authentication.setVisible(false);
//						lblNewLabel.setEnabled(false);
//						lblemailAddress.setEnabled(false);
//						lblNewLabel_1.setEnabled(false);
//						lblNewLabel_5.setEnabled(false);
							}

							if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")
									|| filetype.equalsIgnoreCase("Zoho Mail")) {
								lblEnableImap_p3.setVisible(true);
								lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
										+ " account , you'll need to generate and use an app password or"
										+ System.lineSeparator() + " turn on less secure app</U></HTML>");
								textField_username_p3.setEnabled(true);
								passwordField_p3.setEnabled(true);
							}

							if (filetype.equalsIgnoreCase("Live Exchange")) {
								panel_3_.setVisible(true);
								basic_Authentication.setSelected(true);
								CardLayout card = (CardLayout) panel_3_.getLayout();
								card.show(panel_3_, "panel_3_1_2");
								lbl_Domain.setText("IP or Computer Name");
								panel_3_1_2_1.setVisible(true);
								lblMakeSureYou.setVisible(false);
								lblEnableImap_p3.setVisible(false);
								lblTurnOffTwo_p3.setVisible(false);
								lblNewLabel_5.setVisible(false);
								lblTurnOffTwo_p3.setText("");
								lblMakeSureYou.setText("");
								lblEnableImap_p3.setText("");

							} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {
								panel_3_.setVisible(true);
								basic_Authentication.setSelected(true);
								CardLayout card = (CardLayout) panel_3_.getLayout();
								card.show(panel_3_, "panel_3_1_2");
								lbl_Domain.setText("Amazon Domain Name");
								panel_3_1_2_1.setVisible(true);
								tf_portNo_p3.setVisible(true);
								lblPortNo.setVisible(true);
								lblTurnOffTwo_p3.setText("");
								lblMakeSureYou.setText("");
								lblEnableImap_p3.setText("");
								lblNewLabel_5.setVisible(false);
								lblMakeSureYou.setVisible(false);
								lblEnableImap_p3.setVisible(false);
								lblTurnOffTwo_p3.setVisible(false);
								lblNewLabel_5.setVisible(false);

							} else if (filetype.equalsIgnoreCase("Hostgator email")) {
								panel_3_.setVisible(true);
								CardLayout card = (CardLayout) panel_3_.getLayout();
								basic_Authentication.setSelected(true);
								card.show(panel_3_, "panel_3_1_2");
								lbl_Domain.setText("Hostgator HOST");
								panel_3_1_2_1.setVisible(true);
								tf_portNo_p3.setVisible(true);
								lblPortNo.setVisible(true);
								lblTurnOffTwo_p3.setText("");
								lblMakeSureYou.setText("");
								lblEnableImap_p3.setText("");
								lblNewLabel_5.setVisible(false);
								lblMakeSureYou.setVisible(false);
								lblEnableImap_p3.setVisible(false);
								lblTurnOffTwo_p3.setVisible(false);
								lblNewLabel_5.setVisible(false);

							} else if (filetype.equalsIgnoreCase("IMAP")) {
								panel_3_.setVisible(true);
								CardLayout card = (CardLayout) panel_3_.getLayout();
								card.show(panel_3_, "panel_3_1_2");
								lbl_Domain.setText("IMAP HOST");
								textField_username_p3.setEnabled(true);
								passwordField_p3.setEnabled(true);
								basic_Authentication.setSelected(true);
								panel_3_1_2_1.setVisible(true);
								tf_portNo_p3.setVisible(true);
								chckbxSaveInSame.setVisible(false);
								label_13.setVisible(false);
								lblPortNo.setVisible(true);
								lblTurnOffTwo_p3.setText("");
								panel.setVisible(false);
								lblMakeSureYou.setText("");
								lblEnableImap_p3.setText("");
								lblNewLabel_5.setVisible(false);
								lblMakeSureYou.setVisible(false);
								lblEnableImap_p3.setVisible(false);
								lblTurnOffTwo_p3.setVisible(false);
								lblNewLabel_5.setVisible(false);

							} else if (filetype.equalsIgnoreCase("GoDaddy email")) {
								panel_3_.setVisible(true);
								basic_Authentication.setSelected(true);
								CardLayout card = (CardLayout) panel_3_.getLayout();
								card.show(panel_3_, "panel_3_1_2");
								lblTurnOffTwo_p3.setText("");
								textField_username_p3.setEnabled(true);
								passwordField_p3.setEnabled(true);
								lblMakeSureYou.setText("");
								lblEnableImap_p3.setText("");
								lblNewLabel_5.setVisible(false);
								lblMakeSureYou.setVisible(false);
								lblEnableImap_p3.setVisible(false);
								lblTurnOffTwo_p3.setVisible(false);
								lblNewLabel_5.setVisible(false);

							} else {

								panel_3_.setVisible(true);
								panel.setVisible(true);

								if (filetype.equalsIgnoreCase("OFFICE 365")) {
									modern_Authentication.setSelected(true);
									modern_Authentication.setVisible(true);
//							added by false to true
									textField_username_p3.setEnabled(true);
									passwordField_p3.setEnabled(false);
									chckbxShowPassword_p3.setEnabled(false);
									basic_Authentication.setVisible(true);
									lblNewLabel_1.setVisible(false);
									passwordField_p3.setVisible(false);
									chckbxShowPassword_p3.setVisible(false);
									mailbox.setVisible(true);
									archive.setVisible(true);
									publicfolder.setVisible(true);
									basic_Authentication.setEnabled(false);
								} else if (!(filetype.equalsIgnoreCase("GMAIL")
										|| filetype.equalsIgnoreCase("G-SUITE"))) {
									basic_Authentication.setSelected(false);
									modern_Authentication.setVisible(false);
								}

								if (filetype.equalsIgnoreCase("OFFICE 365")
										|| filetype.equalsIgnoreCase("Live Exchange")) {

									lblEnableImap_p3.setText("<HTML><U>To Enable IMAP</U><HTML>");

									lblEnableImap_p3.setVisible(false);
									lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
											+ " account , you'll need to generate and use an app password.</U></HTML>");
									lblMakeSureYou.setText("Please  Click on The Link");

									panel.setVisible(false);

									lblNewLabel_5.setVisible(false);

//							panel.setVisible(true);
									if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")
											|| fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
											|| fileoptionm.equalsIgnoreCase("OLM File (.olm)")) {
										chckbxRestoreToDefault.setVisible(true);
										task_box.setVisible(true);
										panel_taskfilter.setVisible(true);
									}
									lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
											+ " account , you'll need to generate and use an app password.</U></HTML>");
								}

								CardLayout card = (CardLayout) panel_3_.getLayout();
								card.show(panel_3_, "panel_3_1_2");
							}
							lbl_splitpst.setVisible(false);
							chckbx_splitpst.setSelected(false);
						} else {
							panel_progress.setVisible(true);
							panel_3_.setVisible(true);
							chckbx_splitpst.setVisible(false);
							lbl_splitpst.setVisible(false);
							if (filetype.equalsIgnoreCase("pst")) {

								chckbx_splitpst.setVisible(true);
								lbl_splitpst.setVisible(true);
							}
							if (filetype.equalsIgnoreCase("pdf")) {
								chckbxSavePdfAttachment.setVisible(true);
								label_15.setVisible(true);
								chckbx_convert_pdf_to_pdf.setVisible(true);
								label_pdf_to_pdf.setVisible(true);

							}
							if (filetype.equalsIgnoreCase("DOCX") || filetype.equalsIgnoreCase("DOC")
									|| filetype.equalsIgnoreCase("DOCM") || filetype.equalsIgnoreCase("DOCM")
									|| filetype.equalsIgnoreCase("TIFF") || filetype.equalsIgnoreCase("TXT")
									|| filetype.equalsIgnoreCase("GIF") || filetype.equalsIgnoreCase("JPG")
									|| filetype.equalsIgnoreCase("PNG") || filetype.equalsIgnoreCase("Json")
									|| filetype.equalsIgnoreCase("JPG")) {
								chckbxSavePdfAttachment.setVisible(true);
							}

							if (filetype.equalsIgnoreCase("EML") || filetype.equalsIgnoreCase("MSG")
									|| filetype.equalsIgnoreCase("EMLX") || filetype.equalsIgnoreCase("HTML")
									|| filetype.equalsIgnoreCase("MHTML")) {
								chckbxSavePdfAttachment.setVisible(true);
								label_15.setVisible(true);
							}
							if (filetype.equalsIgnoreCase("VCF") || filetype.equalsIgnoreCase("ICS")) {
								chckbxSavePdfAttachment.setVisible(true);
							}

							System.out.println("this is fileoptionm" + fileoptionm);

							if (fileoptionm.equalsIgnoreCase("MBOX") && filetype.equalsIgnoreCase("PST")) {
								label_16.setVisible(true);
								chckbxSaveMboxIn.setVisible(true);

							}

							if (fileoptionm.equalsIgnoreCase("Maildir") && filetype.equalsIgnoreCase("PST")) {
								chckbxRestoreToDefault.setVisible(true);
								panel_8.setVisible(true);
							}

							CardLayout card = (CardLayout) panel_3_.getLayout();
							card.show(panel_3_, "panel_3_1_1");

							if (!(fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")
									|| fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
									|| fileoptionm.equalsIgnoreCase("OLM File (.olm)"))) {
								chckbxMaintainFolderStructure.setVisible(false);
								label_14.setVisible(false);
								task_box.setVisible(false);
								panel_taskfilter.setVisible(false);
								chckbxRestoreToDefault.setVisible(true);
							} else {
								task_box.setVisible(true);
								panel_taskfilter.setVisible(true);
							}

							panel_3_2.setVisible(true);
							btn_converter_1.setVisible(true);
							btn_converter_1.setEnabled(true);

							if (filetype.equalsIgnoreCase("Opera Mail")) {

								String str = null;

								if (OS.contains("windows")) {
									str = System.getenv("APPDATA").replace("Roaming", "Local") + File.separator
											+ "Opera Mail" + File.separator + "Opera Mail" + File.separator + "Mail"
											+ File.separator + "store";
								} else {
									str = System.getProperty("user.home") + File.separator + "Library" + File.separator
											+ "Application Support" + File.separator + "Opera Mail" + File.separator
											+ "mail";
								}

								System.out.println(str);

								if (new File(str).exists()) {

									tf_Destination_Location.setText(str);

								} else {
									String warn = filetype + " Not Installed Do you want to proceed ?";
									int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle,
											JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
									if (ans == JOptionPane.YES_OPTION) {

									} else {
										SwingUtilities.invokeLater(new Runnable() {

											public void run() {
												comboBox_fileDestination_type.setSelectedItem("PST");
											}
										});
									}

								}

							} else if (filetype.equalsIgnoreCase("Thunderbird")) {

								String str = null;

								if (OS.contains("windows")) {
									str = System.getenv("APPDATA") + File.separator + "Thunderbird" + File.separator
											+ "Profiles";
								} else {
									str = System.getProperty("user.home") + File.separator + "Library" + File.separator
											+ "Thunderbird" + File.separator + "Profiles";
								}

								if (new File(str).exists()) {

									File[] f = new File(str).listFiles();
									for (File fl : f) {
										if (fl != null) {
											if (fl.isDirectory()) {
												String filename = fl.getName();
												String extension = filename.substring(filename.lastIndexOf(".") + 1,
														filename.length());
												String ext = "default";
												if (ext.equals(extension)) {
													// System.out.println(file);

													String defaultfolder = fl.getName();

													str = str + File.separator + defaultfolder + File.separator + "Mail"
															+ File.separator + "Local Folders";

													tf_Destination_Location.setText(str);
													break;
												} else {

												}
											}
										}
									}
								} else {
									String warn = filetype + " Not Installed Do you want to proceed ?";
									int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle,
											JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
									if (ans == JOptionPane.YES_OPTION) {

									} else {
										SwingUtilities.invokeLater(new Runnable() {

											public void run() {
												comboBox_fileDestination_type.setSelectedItem("PST");
											}
										});
									}

								}

							}

							if (!(filetype.equalsIgnoreCase("PST") || filetype.equalsIgnoreCase("Thunderbird")
									|| filetype.equalsIgnoreCase("Opera Mail") || filetype.equalsIgnoreCase("OST")
									|| filetype.equalsIgnoreCase("MBOX") || filetype.equalsIgnoreCase("CSV"))) {
								comboBox.setVisible(true);
								panel_5.setVisible(true);
							}
							if (filetype.equalsIgnoreCase("HOTMAIL")) {
								basic_Authentication.setSelected(true);
								textField_username_p3.setEnabled(true);
								passwordField_p3.setEnabled(true);
							}
						}
					} catch (Exception e) {
						// TODO Auto-generated catch block
						// e.printStackTrace();
						System.out.println("here we riched line no 2984");
					}
				} catch (HeadlessException e) {
					// TODO Auto-generated catch block
					// e.printStackTrace();
					System.out.println("here we riched line no 2989");
				}

			}
		});
		comboBox_fileDestination_type.setFont(new Font("Tahoma", Font.BOLD, 15));
		SwingUtilities.invokeLater(new Runnable() {

			@Override
			public void run() {
				comboBox_fileDestination_type.setSelectedItem("PST");
			}
		});
		panel_3.add(comboBox_fileDestination_type);

		btn_signout_p3 = new JButton("");
		btn_signout_p3.setBounds(841, 9, 142, 31);
		btn_signout_p3.setToolTipText("Click here to Sign Out.");
		btn_signout_p3.setFocusable(false);
		btn_signout_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_signout_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sing-out-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_signout_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sing-out-btn.png")));
			}
		});

		btn_signout_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sing-out-btn.png")));
		btn_signout_p3.setFocusTraversalKeysEnabled(false);
		btn_signout_p3.setFocusPainted(false);
		btn_signout_p3.setDefaultCapable(false);
		btn_signout_p3.setContentAreaFilled(false);
		btn_signout_p3.setBorderPainted(false);
		btn_signout_p3.setRolloverEnabled(false);
		btn_signout_p3.setRequestFocusEnabled(false);
		btn_signout_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					String warn = "Do you want to sign out?";

					int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
					if (ans == JOptionPane.YES_OPTION) {

						if (filetype.equals("OFFICE 365") || filetype.equals("Live Exchange")
								|| filetype.equals("Hotmail")) {
							try {
								if (filetype.equals("Live Exchange") || filetype.equals("Hotmail")) {
									clientforexchange_output.dispose();
								}
								if (filetype.equals("OFFICE 365")) {
									service.close();
								}
							} catch (Exception e) {

							}

						} else {

							try {
								iconnforimap_output.dispose();
							} catch (Exception e) {

							}
						}

						textField_domain_name_p3.setText("");
						passwordField_p3.setText("");
						textField_username_p3.setText("");
						textField_domain_name_p3.setText("");
						btn_signout_p3.setVisible(false);
						btn_converter_1.setEnabled(false);
						radioFileFormat.setEnabled(true);
						btn_previous_p3.setEnabled(true);
						basic_Authentication.setEnabled(true);
						modern_Authentication.setEnabled(true);
						if (filetype.equals("OFFICE 365")) {
							modern_Authentication.setSelected(true);
							lblNewLabel_1.setVisible(false);
							passwordField_p3.setVisible(false);
							chckbxShowPassword_p3.setVisible(false);
							basic_Authentication.setEnabled(false);
						}

						CardLayout card = (CardLayout) panel_3_.getLayout();
						card.show(panel_3_, "panel_3_1_2");
					}

				} catch (Error e) {
					mf.logger.warning("ERROR : " + e.getMessage() + System.lineSeparator());
				} catch (Exception e) {
					mf.logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
					return;
				} finally {
					comboBox_fileDestination_type.setEnabled(true);
				}
			}
		});
		btn_signout_p3.setVisible(false);
		panel_3.add(btn_signout_p3);

		panel_3_2.setBounds(10, 398, 1061, 50);
		panel_3_2.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_3_2.setBackground(Color.WHITE);
		panel_3.add(panel_3_2, "panel_3_2");
		panel_3_2.setLayout(null);
		panel_3_2.setVisible(false);

		tf_Destination_Location.setBackground(Color.WHITE);
		tf_Destination_Location.setBounds(10, 12, 841, 30);
		panel_3_2.add(tf_Destination_Location);
		tf_Destination_Location.setEditable(false);
		tf_Destination_Location.setColumns(10);

		btn_Destination = new JButton("");
		btn_Destination.setToolTipText("Click here to Select Destination Path.");
		btn_Destination.setRolloverEnabled(false);
		btn_Destination.setRequestFocusEnabled(false);
		btn_Destination.setOpaque(false);
		btn_Destination.setFocusable(false);
		btn_Destination.setFocusTraversalKeysEnabled(false);
		btn_Destination.setFocusPainted(false);
		btn_Destination.setDefaultCapable(false);
		btn_Destination.setContentAreaFilled(false);
		btn_Destination.setBorderPainted(false);
		btn_Destination.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_Destination.setIcon(new ImageIcon(Main_Frame.class.getResource("/path-to-save-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_Destination.setIcon(new ImageIcon(Main_Frame.class.getResource("/path-to-save-btn.png")));
			}
		});

		btn_Destination.setIcon(new ImageIcon(Main_Frame.class.getResource("/path-to-save-btn.png")));
		btn_Destination.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					destinationPath();
				} catch (Exception e1) {

					e1.printStackTrace();
				}

			}
		});
		btn_Destination.setBounds(876, 12, 111, 30);
		panel_3_2.add(btn_Destination);
		btn_Destination.setFont(new Font("Tahoma", Font.BOLD, 12));

		progressBar_message_p3 = new JProgressBar();
		progressBar_message_p3.setBackground(Color.WHITE);
		progressBar_message_p3.setBounds(10, 464, 6, 7);
		// progressBar_message_p3.setVisible(false);

		// panel_3.add(progressBar_message_p3);

		panel_3_.setBounds(12, 46, 1059, 334);
		panel_3.add(panel_3_);
		panel_3_.setLayout(new CardLayout(0, 0));

		panel_3_1_1 = new JPanel();
		panel_3_1_1.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_3_1_1.setBackground(Color.WHITE);
		panel_3_.add(panel_3_1_1, "panel_3_1_1");

		JPanel panel_mailfilter = new JPanel();
		panel_mailfilter.setBounds(547, 231, 278, 8);
		panel_mailfilter.setVisible(false);

		datefilter = new JCheckBox("Date Filter");
		datefilter.setBounds(560, 15, 109, 23);
		datefilter.setRolloverEnabled(false);
		datefilter.setRequestFocusEnabled(false);
		datefilter.setOpaque(false);
		datefilter.setFocusable(false);
		datefilter.setFocusPainted(false);
		datefilter.setContentAreaFilled(false);
		datefilter.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				if (datefilter.isSelected()) {
					dateChooser_NewFrom.setEnabled(true);
					dateChooser_NewTo.setEnabled(true);
					remove.setEnabled(true);
					add.setEnabled(true);
				} else {
					dateChooser_NewFrom.setEnabled(false);
					dateChooser_NewTo.setEnabled(false);
					remove.setEnabled(false);
					add.setEnabled(false);
				}
				if (!datefilter.isSelected()) {
					dateChooser_NewFrom.setDate(null);
					dateChooser_NewTo.setDate(null);
					DefaultTableModel model = (DefaultTableModel) table_2.getModel();
					while (model.getRowCount() > 0) {

						for (int i = 0; i < model.getRowCount(); ++i) {

							model.removeRow(i);

						}
					}
				}

			}
		});
		datefilter.setBackground(Color.WHITE);
		panel_mailfilter.setBackground(Color.WHITE);

		dateChooser_mail_fromdate = new JDateChooser();
		dateChooser_mail_fromdate.setBounds(80, 11, 23, 22);
		dateChooser_mail_fromdate.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_mail_fromdate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_mail_fromdate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_mail_fromdate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_mail_fromdate.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				Calendar cal2 = Calendar.getInstance();

				cal2.set(Calendar.HOUR_OF_DAY, 00);
				cal2.set(Calendar.MINUTE, 00);
				cal2.set(Calendar.SECOND, 00);
				Date startdate = cal2.getTime();
				dateChooser_mail_fromdate.setMaxSelectableDate(startdate);
			}
		});
		panel_mailfilter.setLayout(null);
		dateChooser_mail_fromdate.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));
		dateChooser_mail_fromdate.setEnabled(false);
		panel_mailfilter.add(dateChooser_mail_fromdate);

		dateChooser_mail_tilldate = new JDateChooser();
		dateChooser_mail_tilldate.setBounds(424, 11, 23, 19);
		dateChooser_mail_tilldate.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_mail_tilldate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_mail_tilldate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_mail_tilldate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_mail_tilldate.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal3 = Calendar.getInstance();
				cal3.set(Calendar.HOUR_OF_DAY, 23);
				cal3.set(Calendar.MINUTE, 59);
				cal3.set(Calendar.SECOND, 59);
				Date enddate = cal3.getTime();
				dateChooser_mail_tilldate.setMaxSelectableDate(enddate);
				try {
					Calendar calendarstartdate = dateChooser_mail_fromdate.getCalendar();
					calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
					calendarstartdate.set(Calendar.MINUTE, 00);
					calendarstartdate.set(Calendar.SECOND, 00);
					dateChooser_mail_tilldate.setMinSelectableDate(calendarstartdate.getTime());
				} catch (Exception e1) {
					return;
				}
			}
		});
		dateChooser_mail_tilldate.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));
		dateChooser_mail_tilldate.setEnabled(false);
		panel_mailfilter.add(dateChooser_mail_tilldate);

		JLabel label_1 = new JLabel("Start Date");
		label_1.setBounds(198, 11, 9, 20);
		label_1.setFont(new Font("Tahoma", Font.BOLD, 13));
		panel_mailfilter.add(label_1);

		JLabel label_9 = new JLabel("End Date");
		label_9.setBounds(217, 11, 9, 19);
		panel_mailfilter.add(label_9);
		label_9.setFont(new Font("Tahoma", Font.BOLD, 13));

		chckbx_Mail_Filter = new JCheckBox("Mail Filter");
		chckbx_Mail_Filter.setBounds(285, 8, 23, 25);
		panel_mailfilter.add(chckbx_Mail_Filter);
		chckbx_Mail_Filter.setRolloverEnabled(false);
		chckbx_Mail_Filter.setRequestFocusEnabled(false);
		chckbx_Mail_Filter.setOpaque(false);
		chckbx_Mail_Filter.setFocusable(false);
		chckbx_Mail_Filter.setFocusPainted(false);
		chckbx_Mail_Filter.setContentAreaFilled(false);
		chckbx_Mail_Filter.setFont(new Font("Tahoma", Font.BOLD, 12));
		chckbx_Mail_Filter.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					panel_mailfilter.setEnabled(true);
					dateChooser_mail_fromdate.setEnabled(true);
					dateChooser_mail_tilldate.setEnabled(true);
				}

				else {
					panel_mailfilter.setEnabled(false);
					dateChooser_mail_fromdate.setEnabled(false);
					dateChooser_mail_tilldate.setEnabled(false);
				}

			}
		});

		chckbx_Mail_Filter.setBackground(Color.WHITE);

		JPanel panel_checkboxcalenderfilter = new JPanel();
		panel_checkboxcalenderfilter.setBackground(Color.WHITE);
		panel_checkboxcalenderfilter.setBounds(12, 86, 119, 36);
		// panel_3_1_1.add(panel_checkboxcalenderfilter);
		panel_checkboxcalenderfilter.setLayout(null);

		JPanel panel_Calender = new JPanel();
		panel_Calender.setBackground(Color.WHITE);
		panel_Calender.setBounds(12, 119, 1050, 53);
		// panel_3_1_1.add(panel_Calender);
		panel_Calender.setLayout(null);

		chckbx_calender_box = new JCheckBox("Calendar Filter");
		chckbx_calender_box.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					panel_Calender.setEnabled(true);
					dateChooser_calendar_end.setEnabled(true);
					dateChooser_calender_start.setEnabled(true);
				} else {
					panel_Calender.setEnabled(false);
					dateChooser_calendar_end.setEnabled(false);
					dateChooser_calender_start.setEnabled(false);
				}
			}
		});
		chckbx_calender_box.setBackground(Color.WHITE);
		chckbx_calender_box.setBounds(0, 9, 113, 25);
		panel_checkboxcalenderfilter.add(chckbx_calender_box);

		dateChooser_calender_start = new JDateChooser();
		dateChooser_calender_start.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_calender_start.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_calender_start.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_calender_start.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_calender_start.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal2 = Calendar.getInstance();

				cal2.set(Calendar.HOUR_OF_DAY, 00);
				cal2.set(Calendar.MINUTE, 00);
				cal2.set(Calendar.SECOND, 00);
				Date startdate = cal2.getTime();
				dateChooser_calender_start.setMaxSelectableDate(startdate);
			}
		});
		dateChooser_calender_start.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));
		dateChooser_calender_start.setEnabled(false);
		dateChooser_calender_start.setBounds(101, 25, 168, 20);
		panel_Calender.add(dateChooser_calender_start);

		dateChooser_calendar_end = new JDateChooser();
		dateChooser_calendar_end.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_calendar_end.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_calendar_end.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_calendar_end.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_calendar_end.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal3 = Calendar.getInstance();
				cal3.set(Calendar.HOUR_OF_DAY, 23);
				cal3.set(Calendar.MINUTE, 59);
				cal3.set(Calendar.SECOND, 59);
				Date enddate = cal3.getTime();
				dateChooser_calendar_end.setMaxSelectableDate(enddate);
				try {
					Calendar calendarstartdate = dateChooser_mail_fromdate.getCalendar();
					calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
					calendarstartdate.set(Calendar.MINUTE, 00);
					calendarstartdate.set(Calendar.SECOND, 00);
					dateChooser_calendar_end.setMinSelectableDate(calendarstartdate.getTime());
				} catch (Exception e1) {
					return;
				}
			}
		});
		dateChooser_calendar_end.setBounds(893, 25, 147, 19);
		panel_Calender.add(dateChooser_calendar_end);
		dateChooser_calendar_end.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));
		dateChooser_calendar_end.setEnabled(false);

		JLabel label = new JLabel("End Date");
		label.setBounds(815, 25, 79, 19);
		panel_Calender.add(label);
		label.setFont(new Font("Tahoma", Font.BOLD, 15));

		JLabel label_2 = new JLabel("Start Date");
		label_2.setBounds(12, 25, 91, 20);
		panel_Calender.add(label_2);
		label_2.setFont(new Font("Tahoma", Font.BOLD, 15));

		panel_taskfilter = new JPanel();
		panel_taskfilter.setBounds(833, 231, 208, 8);
		panel_taskfilter.setVisible(false);
		panel_taskfilter.setBackground(Color.WHITE);
		panel_taskfilter.setLayout(null);
		JLabel label_8 = new JLabel("Start Date");
		label_8.setFont(new Font("Tahoma", Font.BOLD, 13));
		label_8.setBounds(405, 9, 11, 20);
		panel_taskfilter.add(label_8);

		dateChooser_task_start_date = new JDateChooser();
		dateChooser_task_start_date.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_task_start_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_task_start_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_task_start_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_task_start_date.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal2 = Calendar.getInstance();
				cal2.set(Calendar.HOUR_OF_DAY, 00);
				cal2.set(Calendar.MINUTE, 00);
				cal2.set(Calendar.SECOND, 00);
				Date startdate = cal2.getTime();
				dateChooser_task_start_date.setMaxSelectableDate(startdate);

			}
		});
		dateChooser_task_start_date.setEnabled(false);
		dateChooser_task_start_date.setBounds(181, 9, 23, 22);
		panel_taskfilter.add(dateChooser_task_start_date);

		dateChooser_task_end_date = new JDateChooser();
		dateChooser_task_end_date.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_task_end_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_task_end_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_task_end_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_task_start_date.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal3 = Calendar.getInstance();
				cal3.set(Calendar.HOUR_OF_DAY, 23);
				cal3.set(Calendar.MINUTE, 59);
				cal3.set(Calendar.SECOND, 59);
				Date enddate = cal3.getTime();
				dateChooser_task_end_date.setMaxSelectableDate(enddate);
				try {
					Calendar calendarstartdate = dateChooser_mail_fromdate.getCalendar();
					calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
					calendarstartdate.set(Calendar.MINUTE, 00);
					calendarstartdate.set(Calendar.SECOND, 00);
					dateChooser_task_end_date.setMinSelectableDate(calendarstartdate.getTime());
				} catch (Exception e1) {
					return;
				}
			}
		});
		dateChooser_task_end_date.setEnabled(false);
		dateChooser_task_end_date.setBounds(426, 7, 23, 22);
		panel_taskfilter.add(dateChooser_task_end_date);

		JLabel label_9_1 = new JLabel("End Date");
		label_9_1.setFont(new Font("Tahoma", Font.BOLD, 13));
		label_9_1.setBounds(475, 10, 11, 19);
		panel_taskfilter.add(label_9_1);

		task_box = new JCheckBox("Task Filter");
		task_box.setBounds(115, 13, 11, 16);
		panel_taskfilter.add(task_box);
		task_box.setRolloverEnabled(false);
		task_box.setRequestFocusEnabled(false);
		task_box.setOpaque(false);
		task_box.setFocusable(false);
		task_box.setFocusPainted(false);
		task_box.setContentAreaFilled(false);
		task_box.setFont(new Font("Tahoma", Font.BOLD, 12));
		task_box.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {

					dateChooser_task_end_date.setEnabled(true);
					dateChooser_task_start_date.setEnabled(true);
				} else {

					dateChooser_task_end_date.setEnabled(false);
					dateChooser_task_start_date.setEnabled(false);
				}
			}
		});
		task_box.setBackground(Color.WHITE);

		panel_5 = new JPanel();
		panel_5.setBounds(547, 240, 510, 43);
		panel_5.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_5.setBackground(Color.WHITE);
		panel_5.setVisible(false);

		comboBox = new JComboBox<String>();
		comboBox.setBounds(179, 13, 251, 26);
		comboBox.setBackground(Color.WHITE);
		comboBox.addItem("Subject");
		comboBox.addItem("Subject_Date(DD-MM-YYYY)");
		comboBox.addItem("Subject_Date(MM-DD-YYYY)");
		comboBox.addItem("Subject_Date(YYYY-MM-DD)");
		comboBox.addItem("Subject_Date(YYYY-DD-MM)");
		comboBox.addItem("(DD-MM-YYYY)Date_Subject");
		comboBox.addItem("(MM-DD-YYYY)Date_Subject");
		comboBox.addItem("(YYYY-MM-DD)Date_Subject");
		comboBox.addItem("(YYYY-DD-MM)Date_Subject");
		comboBox.addItem("From_Subject_Date(DD-MM-YYYY)");
		comboBox.addItem("From_Subject_Date(MM-DD-YYYY)");
		comboBox.addItem("From_Subject_Date(YYYY-MM-DD)");
		comboBox.addItem("From_Subject_Date(YYYY-DD-MM)");
		comboBox.addItem("(DD-MM-YYYY)Date_From_Subject");
		comboBox.addItem("(MM-DD-YYYY)Date_From_Subject");
		comboBox.addItem("(YYYY-MM-DD)Date_From_Subject");
		comboBox.addItem("(YYYY-DD-MM)Date_From_Subject");

		comboBox.setVisible(false);

		panel_3_1_2.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_3_1_2.setBackground(Color.WHITE);
		panel_3_.add(panel_3_1_2, "panel_3_1_2");

		lblLive_Chat_p3 = new JLabel("More Help");
		lblLive_Chat_p3.setBounds(880, 16, 66, 25);
		lblLive_Chat_p3.setForeground(Color.RED);
		lblLive_Chat_p3.setCursor(cursor);
		lblLive_Chat_p3.setFont(new Font("Tahoma", Font.PLAIN, 14));
		lblLive_Chat_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {

				openBrowser("http://messenger.providesupport.com/messenger/0pi295uz3ga080c7lxqxxuaoxr.html");

			}
		});

		lbl_connecting_p3 = new JLabel("");
		lbl_connecting_p3.setBounds(447, 201, 85, 32);
		lbl_connecting_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/loading.gif")));
		lbl_connecting_p3.setVisible(false);

		lblNewLabel = new JLabel("User Name");
		lblNewLabel.setBounds(29, 70, 85, 25);
		lblNewLabel.setFont(new Font("Tahoma", Font.BOLD, 14));

		lblNewLabel_1 = new JLabel("Password ");
		lblNewLabel_1.setBounds(29, 97, 77, 25);
		lblNewLabel_1.setFont(new Font("Tahoma", Font.BOLD, 14));

		textField_username_p3.setBounds(332, 67, 375, 25);
		textField_username_p3.setHorizontalAlignment(JTextField.CENTER);
		textField_username_p3.setComponentPopupMenu(jPopupMenu);
		textField_username_p3.setColumns(10);

		passwordField_p3 = new JPasswordField();
		passwordField_p3.setBounds(332, 103, 375, 25);
		passwordField_p3.setHorizontalAlignment(JTextField.CENTER);
		passwordField_p3.setComponentPopupMenu(jPopupMenu);

		chckbxShowPassword_p3 = new JCheckBox("Show Password");
		chckbxShowPassword_p3.setBounds(736, 102, 143, 25);
		chckbxShowPassword_p3.setRolloverEnabled(false);
		chckbxShowPassword_p3.setRequestFocusEnabled(false);
		chckbxShowPassword_p3.setOpaque(false);
		chckbxShowPassword_p3.setFocusable(false);
		chckbxShowPassword_p3.setFocusPainted(false);
		chckbxShowPassword_p3.setContentAreaFilled(false);
		chckbxShowPassword_p3.setBackground(Color.WHITE);
		chckbxShowPassword_p3.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					passwordField_p3.setEchoChar((char) 0);
				}

				else {
					passwordField_p3.setEchoChar('●');
				}
			}
		});
		chckbxShowPassword_p3.setFont(new Font("Tahoma", Font.BOLD, 13));

		JButton btn_Sign_p3 = new JButton("");
		btn_Sign_p3.setBounds(786, 187, 123, 38);
		btn_Sign_p3.setDefaultCapable(false);
		btn_Sign_p3.setBorderPainted(false);
		btn_Sign_p3.setRolloverEnabled(false);
		btn_Sign_p3.setRequestFocusEnabled(false);
		btn_Sign_p3.setOpaque(false);
		btn_Sign_p3.setFocusable(false);
		btn_Sign_p3.setFocusTraversalKeysEnabled(false);
		btn_Sign_p3.setFocusPainted(false);
		btn_Sign_p3.setContentAreaFilled(false);
		btn_Sign_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_Sign_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_Sign_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-btn.png")));
			}
		});

		btn_Sign_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-btn.png")));
		btn_Sign_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				try {
					domain_p3 = textField_domain_name_p3.getText().replaceAll("//s", "");
					domain_p3 = domain_p3.trim();
				} catch (Exception a) {
					domain_p3 = "";
				}
				try {
					username_p3 = textField_username_p3.getText().replaceAll("//s", "");
					username_p3 = username_p3.trim();
				} catch (Exception a) {
					username_p3 = "";
				}
				try {
					password_p3 = new String(passwordField_p3.getPassword());
					password_p3 = password_p3.trim();
				} catch (Exception a) {
					password_p3 = "";
				}
				try {
					portnofiletype = Integer.parseInt(tf_portNo_p3.getText().replaceAll("//s", ""));
				} catch (Exception a) {

				}
				chckbxShowPassword_p3.setEnabled(false);
				tf_portNo_p3.setEnabled(false);
				comboBox_fileDestination_type.setEnabled(false);
				radioFileFormat.setEnabled(false);
				passwordField_p3.setEnabled(false);
				textField_username_p3.setEnabled(false);
				textField_domain_name_p3.setEnabled(false);

				if (modern_Authentication.isSelected()) {

					th = new Thread(new Runnable() {

						@Override
						public void run() {
							lbl_connecting_p3.setVisible(true);

							try {
								basic_Authentication.setEnabled(false);
								modern_Authentication.setEnabled(false);
								radioFileFormat.setEnabled(false);
								btn_previous_p3.setEnabled(false);
								btn_converter_1.setEnabled(false);
								textField_username_p3.setEnabled(true);
								if (filetype.equalsIgnoreCase("OFFICE 365")) {
									// change to modern auth
									ows = new EWSOffice();
									service = ows.loginEWS(textField_username_p3.getText().trim());
									service.validate();
//									ConnectionToOffice.conntiontooffice365_output();
								} else if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")) {
									String token = GetToken.tokenForGmail_output();
									if (token.equals("") || token == null) {
										btn_previous_p3.setEnabled(true);
										comboBox_fileDestination_type.setEnabled(true);
										basic_Authentication.setEnabled(true);
										modern_Authentication.setEnabled(true);

										CardLayout card = (CardLayout) panel_3_.getLayout();
										card.show(panel_3_, "panel_3_1_2");
										return;
									}
									clientforimap_output = GetToken.loginGmail_output(token);
								}
								panel_3_1_2.setVisible(false);
								panel_3_1_2.setVisible(true);
								CardLayout card = (CardLayout) panel_3_.getLayout();
								card.show(panel_3_, "panel_3_1_1");
								btn_signout_p3.setVisible(true);
								btn_converter_1.setEnabled(true);
								btn_converter_1.setVisible(true);
								output = true;
								radioFileFormat.setEnabled(false);
								btn_previous_p3.setEnabled(false);
							} catch (Exception e) {
								basic_Authentication.setEnabled(true);
								modern_Authentication.setEnabled(true);
								radioFileFormat.setEnabled(true);
								btn_previous_p3.setEnabled(true);
								comboBox_fileDestination_type.setEnabled(true);
								if (filetype.equalsIgnoreCase("Gmail") || filetype.equalsIgnoreCase("G-SUITE")) {
									basic_Authentication.setSelected(true);
									if (e.getMessage().contains(
											"AE_1_2_0002 NO [AUTHENTICATIONFAILED] Invalid credentials (Failure)")) {
										JOptionPane.showMessageDialog(mf,
												"Connection Not Established with Gmail please check your Credential OR Otherwise allow 3rd party app to access your account",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else if (e.getMessage().contains(" Application-specific password required: ")) {
										JOptionPane.showMessageDialog(mf, "Application specific password required",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
										comboBox_fileDestination_type.setEnabled(true);
									} else {
										JOptionPane.showMessageDialog(mf, "Connection not established", messageboxtitle,
												JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
										comboBox_fileDestination_type.setEnabled(true);
									}
								} else if (filetype.equalsIgnoreCase("Yahoo Mail")) {
									if (e.getMessage().contains(
											"AE_3_2_0002 NO [AUTHORIZATIONFAILED] LOGIN Invalid credentials")) {
										JOptionPane.showMessageDialog(mf,
												"Connection Not Established with Yahoo Mail please check your Credential Or Otherwise allow 3rd party app to access your account",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else if (e.getMessage().contains(" Application-specific password required: ")) {
										JOptionPane.showMessageDialog(mf, "Application specific password required",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else {
										JOptionPane.showMessageDialog(mf, "Connection not established", messageboxtitle,
												JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));

									}
								} else if (e.getMessage().contains(" Application-specific password required: ")) {
									JOptionPane.showMessageDialog(mf, "Application specific password required",
											messageboxtitle, JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));
								} else {
									JOptionPane.showMessageDialog(mf, "Connection not established", messageboxtitle,
											JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));
								}

							} finally {
								lbl_connecting_p3.setVisible(false);

								tf_portNo_p3.setEnabled(true);

								passwordField_p3.setEnabled(true);
								textField_username_p3.setEnabled(true);
								textField_domain_name_p3.setEnabled(true);
								chckbxShowPassword_p3.setEnabled(true);
								basic_Authentication.setEnabled(true);
								if (filetype.equalsIgnoreCase("OFFICE 365")) {
									lblNewLabel_1.setVisible(false);
									passwordField_p3.setVisible(false);
									chckbxShowPassword_p3.setVisible(false);
									basic_Authentication.setEnabled(false);
								} else {

									lblNewLabel_1.setVisible(true);
									passwordField_p3.setVisible(true);
									chckbxShowPassword_p3.setVisible(true);

								}
//								btn_previous_p3.setEnabled(true);

							}

						}
					});
					th.start();

				} else {
					basic_Authentication.setEnabled(false);

					modern_Authentication.setEnabled(false);

					if (username_p3.equalsIgnoreCase("") || password_p3.equalsIgnoreCase("")) {

						if (username_p3.equalsIgnoreCase("") && password_p3.equalsIgnoreCase("")) {
							JOptionPane.showMessageDialog(mf, "User name and Password fields can't be empty",
									messageboxtitle, JOptionPane.ERROR_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/information.png")));

						} else if (username_p3.equalsIgnoreCase("")) {

							JOptionPane.showMessageDialog(mf, "User name field can't be empty", messageboxtitle,
									JOptionPane.ERROR_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/information.png")));

						} else if (password_p3.equalsIgnoreCase("")) {

							JOptionPane.showMessageDialog(mf, "Password field can't be empty", messageboxtitle,
									JOptionPane.ERROR_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/information.png")));

						}

						chckbxShowPassword_p3.setEnabled(true);
						radioFileFormat.setEnabled(true);
						comboBox_fileDestination_type.setEnabled(true);
						passwordField_p3.setEnabled(true);
						tf_portNo_p3.setEnabled(true);
						textField_username_p3.setEnabled(true);
						textField_domain_name_p3.setEnabled(true);
						basic_Authentication.setEnabled(true);

						modern_Authentication.setEnabled(true);
					} else if (filetype.equalsIgnoreCase("Live Exchange") && domain_p3.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(mf, "Computer Name or IP Address field can not be empty",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));
						comboBox_FiletypeChooser.setEnabled(true);

						btn_next_pane2.setEnabled(true);
						tf_portNo_p3.setEnabled(true);
						radioFileFormat.setEnabled(true);
						comboBox_fileDestination_type.setEnabled(true);
						passwordField_p3.setEnabled(true);
						textField_username_p3.setEnabled(true);
						chckbxShowPassword_p3.setEnabled(true);
						textField_domain_name_p3.setEnabled(true);

					} else if (filetype.equalsIgnoreCase("IMAP") && domain_p3.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(mf, "IMAP Host field can't be empty", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));
						comboBox_FiletypeChooser.setEnabled(true);

						btn_next_pane2.setEnabled(true);
						tf_portNo_p3.setEnabled(true);
						radioFileFormat.setEnabled(true);
						comboBox_fileDestination_type.setEnabled(true);
						passwordField_p3.setEnabled(true);
						textField_username_p3.setEnabled(true);
						chckbxShowPassword_p3.setEnabled(true);
						textField_domain_name_p3.setEnabled(true);

					} else if (filetype.equalsIgnoreCase("IMAP") && tf_portNo_p3.getText().isEmpty()) {

						JOptionPane.showMessageDialog(mf, "Port No field can't be empty", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));
						comboBox_FiletypeChooser.setEnabled(true);
						tf_portNo_p3.setEnabled(true);
						radioFileFormat.setEnabled(true);
						comboBox_fileDestination_type.setEnabled(true);
						passwordField_p3.setEnabled(true);
						textField_username_p3.setEnabled(true);
						chckbxShowPassword_p3.setEnabled(true);
						textField_domain_name_p3.setEnabled(true);

						btn_next_pane2.setEnabled(true);

					} else if (!(isValid(username_p3))) {

						JOptionPane.showMessageDialog(mf, "Please enter a valid username", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));
						tf_portNo_p3.setEnabled(true);
						radioFileFormat.setEnabled(true);
						comboBox_fileDestination_type.setEnabled(true);
						passwordField_p3.setEnabled(true);
						textField_username_p3.setEnabled(true);
						chckbxShowPassword_p3.setEnabled(true);
						textField_domain_name_p3.setEnabled(true);
					} else {

						th = new Thread(new Runnable() {

							@Override
							public void run() {
								lbl_connecting_p3.setVisible(true);

								try {
									radioFileFormat.setEnabled(false);
									btn_previous_p3.setEnabled(false);
									btn_converter_1.setEnabled(false);

									if (filetype.equalsIgnoreCase("OFFICE 365")) {
										conntiontooffice365_output();
									}
									if (filetype.equalsIgnoreCase("HOTMAIL")) {
										conntiontohotmail_output();

									} else if (filetype.equalsIgnoreCase("Yandex Mail")) {
										connectiontoYandex_output();

									} else if (filetype.equalsIgnoreCase("Zoho Mail")) {
										connectiontozoho_output();

									} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {
										connectiontoinaws_output();

									} else if (filetype.equalsIgnoreCase("Hostgator email")) {
										connectiontoHostgator_output();

									} else if (filetype.equalsIgnoreCase("Icloud")) {
										connectiontoicloud_output();

									} else if (filetype.equalsIgnoreCase("GoDaddy email")) {
										connectiontoGoDaddy_output();

									} else if (filetype.equalsIgnoreCase("GMAIL")
											|| filetype.equalsIgnoreCase("G-SUITE")) {
										connectiontogmail_output();

									} else if (filetype.equalsIgnoreCase("Live Exchange")) {
										connectionwithexchangeserver_output();

									} else if (filetype.equalsIgnoreCase("IMAP")) {
										connectiontoimap_output();
									} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {
										connectiontoyahoo_output();

									} else if (filetype.equalsIgnoreCase("AOL")) {
										connectiontoaol_output();

									}
									panel_3_1_2.setVisible(false);
									panel_3_1_2.setVisible(true);
									CardLayout card = (CardLayout) panel_3_.getLayout();
									card.show(panel_3_, "panel_3_1_1");
									btn_signout_p3.setVisible(true);
									btn_converter_1.setEnabled(true);
									btn_converter_1.setVisible(true);
									output = true;
									btn_previous_p3.setEnabled(false);
								} catch (Exception e) {
									btn_previous_p3.setEnabled(true);

									System.out.println("here we riched  4374");

									radioFileFormat.setEnabled(true);
									if (filetype.equalsIgnoreCase("Gmail") || filetype.equalsIgnoreCase("G-SUITE")) {
										if (e.getMessage().contains(
												"AE_1_2_0002 NO [AUTHENTICATIONFAILED] Invalid credentials (Failure)")) {
											JOptionPane.showMessageDialog(mf,
													"Connection Not Estalished with Gmail please check your Credantial OR Otherwise allow 3rd party app to acess your account",
													messageboxtitle, JOptionPane.ERROR_MESSAGE,
													new ImageIcon(Main_Frame.class.getResource("/information.png")));
										} else if (e.getMessage()
												.contains(" Application-specific password required: ")) {
											JOptionPane.showMessageDialog(mf, "Application specific password required",
													messageboxtitle, JOptionPane.ERROR_MESSAGE,
													new ImageIcon(Main_Frame.class.getResource("/information.png")));
										} else {
											System.out.println("line no 4392");

											JOptionPane.showMessageDialog(mf, "Connection not established",
													messageboxtitle, JOptionPane.ERROR_MESSAGE,
													new ImageIcon(Main_Frame.class.getResource("/information.png")));
										}
									} else if (filetype.equalsIgnoreCase("Yahoo Mail")) {
										if (e.getMessage().contains(
												"AE_3_2_0002 NO [AUTHORIZATIONFAILED] LOGIN Invalid credentials")) {
											JOptionPane.showMessageDialog(mf,
													"Connection Not Estalished with Yahoo Mail please check your Credantial Otherwise allow 3rd party app to acess your account",
													messageboxtitle, JOptionPane.ERROR_MESSAGE,
													new ImageIcon(Main_Frame.class.getResource("/information.png")));
										} else if (e.getMessage()
												.contains(" Application-specific password required: ")) {
											JOptionPane.showMessageDialog(mf, "Application specific password required",
													messageboxtitle, JOptionPane.ERROR_MESSAGE,
													new ImageIcon(Main_Frame.class.getResource("/information.png")));
										} else {
											JOptionPane.showMessageDialog(mf, "Connection not established",
													messageboxtitle, JOptionPane.ERROR_MESSAGE,
													new ImageIcon(Main_Frame.class.getResource("/information.png")));
										}
									} else if (e.getMessage().contains(" Application-specific password required: ")) {
										JOptionPane.showMessageDialog(mf, "Application specific password required",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else {
										JOptionPane.showMessageDialog(mf, "Connection not established", messageboxtitle,
												JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									}

								} finally {

									basic_Authentication.setEnabled(true);

									modern_Authentication.setEnabled(true);
									System.out.println("finally combo");
									comboBox_fileDestination_type.setEnabled(true);

									lbl_connecting_p3.setVisible(false);

									tf_portNo_p3.setEnabled(true);

									passwordField_p3.setEnabled(true);
									textField_username_p3.setEnabled(true);
									textField_domain_name_p3.setEnabled(true);
									chckbxShowPassword_p3.setEnabled(true);
//									btn_previous_p3.setEnabled(true);

								}

							}
						});
						th.start();

					}
				}

			}
		});
		btn_Sign_p3.setFont(new Font("Tahoma", Font.BOLD, 14));

		panel_3_1_2_1.setBounds(8, 13, 706, 53);
		panel_3_1_2_1.setBackground(Color.WHITE);
		panel_3_1_2_1.setVisible(false);

		textField_domain_name_p3.setHorizontalAlignment(JTextField.CENTER);
		textField_domain_name_p3.setComponentPopupMenu(jPopupMenu);
		textField_domain_name_p3.setColumns(10);

		lbl_Domain = new JLabel("");
		lbl_Domain.setFont(new Font("Tahoma", Font.BOLD, 14));

		lblPortNo = new JLabel("Port No.");
		lblPortNo.setBounds(29, 142, 158, 32);
		lblPortNo.setFont(new Font("Tahoma", Font.BOLD, 14));

		tf_portNo_p3 = new JTextField();
		tf_portNo_p3.setBounds(332, 142, 375, 25);
		tf_portNo_p3.setHorizontalAlignment(JTextField.CENTER);
		tf_portNo_p3.setText(Integer.toString(993));
		tf_portNo_p3.setColumns(10);

		panel = new JPanel();
		panel.setBounds(29, 272, 1010, 53);
		panel.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel.setBackground(Color.WHITE);

		lblMakeSureYou = new JLabel("Please  Click on The Link");
		lblMakeSureYou.setForeground(Color.BLACK);
		lblMakeSureYou.setFont(new Font("Tahoma", Font.BOLD, 14));

		lblEnableImap_p3 = new JLabel("<HTML><U>To Enable IMAP</U></HTML>");
		lblEnableImap_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				mf.openbrowserenableimap(filetype);
			}
		});
		lblEnableImap_p3.setForeground(Color.BLUE);
		lblEnableImap_p3.setCursor(cursor);
		lblEnableImap_p3.setFont(new Font("Tahoma", Font.PLAIN, 11));

		lblTurnOffTwo_p3 = new JLabel("<HTML><U>Turn Off Two Step Verification</U></HTML>");
		lblTurnOffTwo_p3.setCursor(cursor);
		lblTurnOffTwo_p3.setForeground(Color.BLUE);
		lblTurnOffTwo_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				mf.openbrowserturntwostepoff(filetype);

			}
		});
		lblTurnOffTwo_p3.setFont(new Font("Tahoma", Font.PLAIN, 11));
		lblNewLabel_8 = new JLabel(
				"<html>If you have a third-party app password for your Yahoo Mail account, then please proceed by entering the credentials at the required place. But if<br/>    you don't have the app password or you won't be able to create it, then this software does not offer a facility to complete the conversion process.<html/>");
		lblNewLabel_8.setBounds(134, 236, 905, 38);
		lblNewLabel_8.setFont(new Font("Tahoma", Font.BOLD, 10));
		lblNewLabel_8.setVisible(false);

		lblNewLabel_9 = new JLabel("      Note:");
		lblNewLabel_9.setBounds(60, 248, 64, 13);
		lblNewLabel_9.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNewLabel_9.setForeground(Color.RED);
		lblNewLabel_9.setVisible(false);
		lblNewLabel_5 = new JLabel("(Use third party App Password)");
		lblNewLabel_5.setBounds(28, 116, 200, 18);
		lblNewLabel_5.setForeground(Color.RED);

		lblemailAddress = new JLabel("(Email Address)");
		lblemailAddress.setBounds(118, 74, 158, 18);
		lblemailAddress.setForeground(Color.RED);

		basic_Authentication = new JRadioButton("Basic Authentication");
		basic_Authentication.setBounds(29, 193, 203, 32);
		basic_Authentication.setVisible(false);
		basic_Authentication.setBorderPainted(false);
		basic_Authentication.setContentAreaFilled(false);
		basic_Authentication.setFocusable(false);
		basic_Authentication.setFocusTraversalKeysEnabled(false);
		basic_Authentication.setFocusPainted(false);
		basic_Authentication.setRolloverEnabled(false);
		basic_Authentication.setRequestFocusEnabled(false);
		basic_Authentication.setOpaque(false);
		basic_Authentication.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				textField_username_p3.setEnabled(true);
				passwordField_p3.setEnabled(true);
				tf_portNo_p3.setEnabled(true);
				chckbxShowPassword_p3.setEnabled(true);
				lblPortNo.setEnabled(true);
				lblNewLabel_5.setEnabled(true);
				lblNewLabel_1.setEnabled(true);
				lblNewLabel.setEnabled(true);
				lblemailAddress.setEnabled(true);
//				lblemailAddress.setEnabled(false);
				if (filetype.equalsIgnoreCase("OFFICE 365")) {
					panel.setVisible(false);
				} else {
					panel.setVisible(true);
				}
			}
		});
		basic_Authentication.setBackground(Color.WHITE);
		basic_Authentication.setFont(new Font("Tahoma", Font.BOLD, 12));
		buttonGroup.add(basic_Authentication);

		modern_Authentication = new JRadioButton("Modern Authentication");
		modern_Authentication.setBounds(29, 228, 203, 32);
		buttonGroup.add(modern_Authentication);
		modern_Authentication.setVisible(false);
		modern_Authentication.setBorderPainted(false);
		modern_Authentication.setContentAreaFilled(false);
		modern_Authentication.setFocusable(false);
		modern_Authentication.setFocusTraversalKeysEnabled(false);
		modern_Authentication.setFocusPainted(false);
		modern_Authentication.setRolloverEnabled(false);
		modern_Authentication.setRequestFocusEnabled(false);
		modern_Authentication.setOpaque(false);
		modern_Authentication.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				textField_username_p3.setText("");
				passwordField_p3.setText("");
				textField_username_p3.setEnabled(true);
				passwordField_p3.setEnabled(false);
				tf_portNo_p3.setEnabled(false);
				chckbxShowPassword_p3.setEnabled(false);
				lblPortNo.setEnabled(false);
				lblNewLabel_5.setEnabled(false);
				lblNewLabel_1.setEnabled(false);
				lblNewLabel.setEnabled(true);
				lblemailAddress.setEnabled(true);
//				lblemailAddress.setEnabled(false);
				if (filetype.equalsIgnoreCase("OFFICE 365")) {

					panel.setVisible(false);
				} else {
					panel.setVisible(true);
				}
			}
		});
		modern_Authentication.setBackground(Color.WHITE);
		modern_Authentication.setFont(new Font("Tahoma", Font.BOLD, 12));
		GroupLayout gl_panel = new GroupLayout(panel);
		gl_panel.setHorizontalGroup(gl_panel.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_panel.createSequentialGroup().addGap(8)
						.addComponent(lblMakeSureYou, GroupLayout.PREFERRED_SIZE, 188, GroupLayout.PREFERRED_SIZE)
						.addGap(10).addComponent(lblTurnOffTwo_p3, GroupLayout.DEFAULT_SIZE, 516, Short.MAX_VALUE)
						.addGap(33)
						.addComponent(lblEnableImap_p3, GroupLayout.PREFERRED_SIZE, 77, GroupLayout.PREFERRED_SIZE)
						.addGap(178)));
		gl_panel.setVerticalGroup(gl_panel.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_panel.createSequentialGroup().addGap(9).addComponent(lblMakeSureYou,
						GroupLayout.PREFERRED_SIZE, 32, GroupLayout.PREFERRED_SIZE))
				.addGroup(gl_panel.createSequentialGroup().addGap(6).addComponent(lblTurnOffTwo_p3,
						GroupLayout.PREFERRED_SIZE, 35, GroupLayout.PREFERRED_SIZE))
				.addGroup(gl_panel.createSequentialGroup().addGap(15).addComponent(lblEnableImap_p3,
						GroupLayout.PREFERRED_SIZE, 25, GroupLayout.PREFERRED_SIZE)));
		panel.setLayout(gl_panel);
		panel_3_1_2.setLayout(null);
		GroupLayout gl_panel_3_1_2_1 = new GroupLayout(panel_3_1_2_1);
		gl_panel_3_1_2_1.setHorizontalGroup(gl_panel_3_1_2_1.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_panel_3_1_2_1.createSequentialGroup().addGap(20)
						.addComponent(lbl_Domain, GroupLayout.PREFERRED_SIZE, 271, GroupLayout.PREFERRED_SIZE)
						.addGap(33)
						.addComponent(textField_domain_name_p3, GroupLayout.DEFAULT_SIZE, 379, Short.MAX_VALUE)
						.addGap(7)));
		gl_panel_3_1_2_1
				.setVerticalGroup(gl_panel_3_1_2_1.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_panel_3_1_2_1.createSequentialGroup().addGap(15)
								.addGroup(gl_panel_3_1_2_1.createParallelGroup(Alignment.LEADING)
										.addComponent(lbl_Domain, GroupLayout.PREFERRED_SIZE, 27,
												GroupLayout.PREFERRED_SIZE)
										.addComponent(textField_domain_name_p3, GroupLayout.PREFERRED_SIZE, 26,
												GroupLayout.PREFERRED_SIZE))));
		panel_3_1_2_1.setLayout(gl_panel_3_1_2_1);
		panel_3_1_2.add(panel_3_1_2_1);
		panel_3_1_2.add(lblLive_Chat_p3);
		panel_3_1_2.add(lblNewLabel_9);
		panel_3_1_2.add(lbl_connecting_p3);
		panel_3_1_2.add(btn_Sign_p3);
		panel_3_1_2.add(lblNewLabel_8);
		panel_3_1_2.add(panel);
		panel_3_1_2.add(basic_Authentication);
		panel_3_1_2.add(modern_Authentication);
		panel_3_1_2.add(lblNewLabel);
		panel_3_1_2.add(lblemailAddress);
		panel_3_1_2.add(lblPortNo);
		panel_3_1_2.add(lblNewLabel_1);
		panel_3_1_2.add(lblNewLabel_5);
		panel_3_1_2.add(passwordField_p3);
		panel_3_1_2.add(chckbxShowPassword_p3);
		panel_3_1_2.add(textField_username_p3);
		panel_3_1_2.add(tf_portNo_p3);

		panel_progress = new JPanel();
		panel_progress.setBounds(10, 459, 1061, 74);

		panel_progress.setBackground(Color.WHITE);
		panel_3.add(panel_progress);

		btnStop = new JButton("");
		btnStop.setBounds(867, 3, 133, 30);
		btnStop.setToolTipText("Click here to Stop The Process.");
		btnStop.setContentAreaFilled(false);
		btnStop.setBorderPainted(false);
		btnStop.setBackground(Color.WHITE);
		btnStop.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnStop.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnStop.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
			}
		});

		btnStop.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
		btnStop.setRolloverEnabled(false);
		btnStop.setVisible(false);
		btnStop.setRequestFocusEnabled(false);
		btnStop.setOpaque(false);
		btnStop.setFocusable(false);
		btnStop.setFocusTraversalKeysEnabled(false);
		btnStop.setFocusPainted(false);
		btnStop.setDefaultCapable(false);
		SwingUtilities.invokeLater(new Runnable() {

			@Override
			public void run() {
				comboBox.setSelectedItem("Subject");

				lblnamingconvention = new JLabel("Naming Convention");
				lblnamingconvention.setFont(new Font("Tahoma", Font.BOLD, 11));

			}
		});

		Progressbar = new JLabel("");
		Progressbar.setBounds(8, 3, 835, 33);
		Progressbar.setVisible(false);
		Progressbar.setIcon(new ImageIcon(Main_Frame.class.getResource("/progress-bar.gif")));

		lbl_progressreport = new JLabel("");
		lbl_progressreport.setBounds(8, 46, 835, 24);

		JPanel panel_duplicacy = new JPanel();
		panel_duplicacy.setBounds(10, 7, 527, 320);
		panel_duplicacy.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_duplicacy.setBackground(Color.WHITE);

		chckbxRemoveDuplicacy = new JCheckBox("Remove Duplicate Mail On basis of To, From, Subject, Bcc & Body\r\n");
		chckbxRemoveDuplicacy.setBounds(6, 7, 391, 24);
		chckbxRemoveDuplicacy.setRolloverEnabled(false);
		chckbxRemoveDuplicacy.setRequestFocusEnabled(false);
		chckbxRemoveDuplicacy.setOpaque(false);
		chckbxRemoveDuplicacy.setFocusable(false);
		chckbxRemoveDuplicacy.setFocusPainted(false);
		chckbxRemoveDuplicacy.setContentAreaFilled(false);
		chckbxRemoveDuplicacy.setForeground(Color.RED);
		chckbxRemoveDuplicacy.setFont(new Font("Tahoma", Font.BOLD, 10));
		chckbxRemoveDuplicacy.setBackground(Color.WHITE);

		lbl_splitpst.setBounds(425, 196, 26, 23);
		lbl_splitpst.setVisible(false);
		lbl_splitpst.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		lbl_splitpst.setToolTipText("Split resultant PST file according to size.");

		chckbx_splitpst.setBounds(6, 196, 147, 23);
		chckbx_splitpst.setRolloverEnabled(false);
		chckbx_splitpst.setRequestFocusEnabled(false);
		chckbx_splitpst.setOpaque(false);
		chckbx_splitpst.setFocusable(false);
		chckbx_splitpst.setFocusPainted(false);
		chckbx_splitpst.setContentAreaFilled(false);
		chckbx_splitpst.setFont(new Font("Tahoma", Font.BOLD, 10));
		chckbx_splitpst.setForeground(Color.RED);
		chckbx_splitpst.setBackground(Color.WHITE);
		chckbx_splitpst.setVisible(false);
		chckbx_splitpst.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				if (chckbx_splitpst.isSelected()) {

					spinner_sizespinner.setVisible(true);
					comboBox_setsize.setVisible(true);
					spinner_sizespinner.setEnabled(true);
					comboBox_setsize.setEnabled(true);

				} else {

					spinner_sizespinner.setVisible(false);
					comboBox_setsize.setVisible(false);
					spinner_sizespinner.setEnabled(false);
					comboBox_setsize.setEnabled(false);

				}

			}
		});
		chckbx_splitpst.setForeground(Color.RED);
		chckbx_splitpst.setFont(new Font("Tahoma", Font.BOLD, 10));
		chckbx_splitpst.setBackground(Color.WHITE);

		spinner_sizespinner = new JSpinner();
		spinner_sizespinner.setBounds(175, 199, 52, 22);
		spinner_sizespinner.setVisible(false);
		spinner_sizespinner.setFont(new Font("Calibri", Font.PLAIN, 14));
		spinner_sizespinner.setBackground(Color.WHITE);
		spinner_sizespinner.setFont(new Font("Calibri", Font.PLAIN, 14));
		spinner_sizespinner.setEditor(new JSpinner.DefaultEditor(spinner_sizespinner));
		spinner_sizespinner.setBackground(Color.WHITE);
		SpinnerModel sm = new SpinnerNumberModel(5, 1, 900, 1);

		spinner_sizespinner.setModel(sm);
		spinner_sizespinner.setValue(1);

		spinner_sizespinner.setEditor(new JSpinner.DefaultEditor(spinner_sizespinner));

		comboBox_setsize = new JComboBox();
		comboBox_setsize.setBounds(248, 199, 87, 22);
		comboBox_setsize.setVisible(false);
		comboBox_setsize.setFont(new Font("Calibri", Font.PLAIN, 14));
		comboBox_setsize.setBackground(Color.WHITE);
		comboBox_setsize.setFont(new Font("Calibri", Font.PLAIN, 14));
		comboBox_setsize.setBackground(Color.WHITE);
		comboBox_setsize.addItem("MB");
		comboBox_setsize.addItem("GB");
		comboBox_setsize.setSelectedItem(0);
		comboBox_setsize.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (comboBox_setsize.getSelectedIndex() == 0) {
					spinner_sizespinner.setEditor(new JSpinner.NumberEditor(spinner_sizespinner));
					SpinnerModel sm = new SpinnerNumberModel(1, 1, 900, 1);
					spinner_sizespinner.setModel(sm);

					spinner_sizespinner.setValue(sm.getValue());
					spinner_sizespinner.setEditor(new JSpinner.DefaultEditor(spinner_sizespinner));
				} else if (comboBox_setsize.getSelectedIndex() == 1) {
					spinner_sizespinner.setEditor(new JSpinner.NumberEditor(spinner_sizespinner));
					SpinnerModel sm = new SpinnerNumberModel(1, 1, 20, 1);
					spinner_sizespinner.setModel(sm);
					spinner_sizespinner.setValue(sm.getValue());
					spinner_sizespinner.setEditor(new JSpinner.DefaultEditor(spinner_sizespinner));
				}
			}
		});

		chckbxMaintainFolderStructure = new JCheckBox("Maintain Folder Hierarchy");
		chckbxMaintainFolderStructure.setBounds(6, 60, 344, 25);
		chckbxMaintainFolderStructure.setRolloverEnabled(false);
		chckbxMaintainFolderStructure.setRequestFocusEnabled(false);
		chckbxMaintainFolderStructure.setOpaque(false);
		chckbxMaintainFolderStructure.setFocusable(false);
		chckbxMaintainFolderStructure.setFocusPainted(false);
		chckbxMaintainFolderStructure.setContentAreaFilled(false);
		chckbxMaintainFolderStructure.setForeground(Color.RED);
		chckbxMaintainFolderStructure.setFont(new Font("Tahoma", Font.BOLD, 10));
		chckbxMaintainFolderStructure.setBackground(Color.WHITE);

		chckbxSaveInSame.setBounds(6, 34, 391, 23);

		chckbxSaveInSame.setRolloverEnabled(false);
		chckbxSaveInSame.setRequestFocusEnabled(false);
		chckbxSaveInSame.setOpaque(false);
		chckbxSaveInSame.setFocusable(false);
		chckbxSaveInSame.setFocusPainted(false);
		chckbxSaveInSame.setContentAreaFilled(false);
		chckbxSaveInSame.setFont(new Font("Tahoma", Font.BOLD, 10));
		chckbxSaveInSame.setBackground(Color.WHITE);
		chckbxSaveInSame.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {

				if (arg0.getStateChange() == ItemEvent.SELECTED) {

					destination = tf_Destination_Location.getText();
					btn_Destination.setEnabled(false);
					tf_Destination_Location.setText(filepath(new File(filepath)));
				}

				else {
					btn_Destination.setEnabled(true);
					tf_Destination_Location.setText(destination);

				}
			}
		});
		chckbxSaveInSame.setForeground(new Color(255, 0, 0));

		chckbxSavePdfAttachment.setBounds(6, 88, 365, 23);
		chckbxSavePdfAttachment.setRolloverEnabled(false);
		chckbxSavePdfAttachment.setRequestFocusEnabled(false);
		chckbxSavePdfAttachment.setOpaque(false);
		chckbxSavePdfAttachment.setFocusable(false);
		chckbxSavePdfAttachment.setFocusPainted(false);
		chckbxSavePdfAttachment.setContentAreaFilled(false);
		chckbxSavePdfAttachment.setForeground(Color.RED);
		chckbxSavePdfAttachment.setVisible(false);
		chckbxSavePdfAttachment.setFont(new Font("Tahoma", Font.BOLD, 10));
		chckbxSavePdfAttachment.setBackground(Color.WHITE);
		chckbxSavePdfAttachment.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (chckbxMigrateOrBackup.isSelected() && chckbxSavePdfAttachment.isSelected()) {

					JOptionPane.showMessageDialog(main_multiplefile.this,
							"Please select one of them checkbox before continuing.", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					chckbxSavePdfAttachment.setSelected(false);
				} else if (chckbx_convert_pdf_to_pdf.isSelected() && chckbxSavePdfAttachment.isSelected()) {

					JOptionPane.showMessageDialog(main_multiplefile.this,
							"Please select one of them checkbox before continuing.", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					chckbxSavePdfAttachment.setSelected(false);
				} else if (chckbx_convert_pdf_to_pdf.isSelected() && chckbxMigrateOrBackup.isSelected()) {

					JOptionPane.showMessageDialog(main_multiplefile.this,
							"Please select one of them checkbox before continuing.", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					chckbxSavePdfAttachment.setSelected(false);
				}
			}
		});

		label_12 = new JLabel("");
		label_12.setBounds(425, 7, 26, 23);
		label_12.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_12.setToolTipText("All the replicated or duplicated emails will be removed " + System.lineSeparator()
				+ "on the basis of To, From, Subject, Bcc and Body.");

		label_13 = new JLabel("");
		label_13.setBounds(425, 34, 26, 23);
		label_13.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_13.setToolTipText("All the resultant data will get saved at the " + System.lineSeparator()
				+ "location of the source file.");

		label_14 = new JLabel("");
		label_14.setBounds(425, 60, 26, 23);
		label_14.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_14.setToolTipText("Maintain the folder hierarchy of your mailbox.");

		label_15.setBounds(425, 88, 26, 23);
		label_15.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_15.setToolTipText(
				"Save all the email attachments separately in a  " + System.lineSeparator() + "folder.");

		chckbxSaveMboxIn = new JCheckBox("Save Mbox in same PST/OST");
		chckbxSaveMboxIn.setBounds(6, 114, 271, 23);
		chckbxSaveMboxIn.setRolloverEnabled(false);
		chckbxSaveMboxIn.setRequestFocusEnabled(false);
		chckbxSaveMboxIn.setOpaque(false);
		chckbxSaveMboxIn.setFocusable(false);
		chckbxSaveMboxIn.setFocusPainted(false);
		chckbxSaveMboxIn.setContentAreaFilled(false);
		chckbxSaveMboxIn.setForeground(Color.RED);
		chckbxSaveMboxIn.setVisible(false);
		chckbxSaveMboxIn.setFont(new Font("Tahoma", Font.BOLD, 10));
		chckbxSaveMboxIn.setBackground(Color.WHITE);

		label_16 = new JLabel("");
		label_16.setBounds(425, 114, 26, 23);
		label_16.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_16.setToolTipText("Save all Mbox file in same PST and Ost");
		label_16.setVisible(false);

		chckbxMigrateOrBackup = new JCheckBox("Migrate or Backup Emails Without Attachment files");
		chckbxMigrateOrBackup.setBounds(6, 140, 414, 23);
		chckbxMigrateOrBackup.setRolloverEnabled(false);
		chckbxMigrateOrBackup.setRequestFocusEnabled(false);
		chckbxMigrateOrBackup.setOpaque(false);
		chckbxMigrateOrBackup.setFocusable(false);
		chckbxMigrateOrBackup.setFocusPainted(false);
		chckbxMigrateOrBackup.setContentAreaFilled(false);
		chckbxMigrateOrBackup.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (chckbxMigrateOrBackup.isSelected() && chckbxSavePdfAttachment.isSelected()) {

					JOptionPane.showMessageDialog(main_multiplefile.this,
							"Please select one of them checkbox before continuing.", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					chckbxSavePdfAttachment.setSelected(false);
				} else if (chckbx_convert_pdf_to_pdf.isSelected() && chckbxMigrateOrBackup.isSelected()) {

					JOptionPane.showMessageDialog(main_multiplefile.this,
							"Please select one of them checkbox before continuing.", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					chckbxSavePdfAttachment.setSelected(false);
				} else if (chckbx_convert_pdf_to_pdf.isSelected() && chckbxSavePdfAttachment.isSelected()) {

					JOptionPane.showMessageDialog(main_multiplefile.this,
							"Please select one of them checkbox before continuing.", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					chckbxSavePdfAttachment.setSelected(false);
				}

//				if (chckbxMigrateOrBackup.isSelected()) {
//					label_15.setVisible(false);
//					chckbxSavePdfAttachment.setVisible(false);
//				}
//				else if (!chckbxMigrateOrBackup.isSelected()) {
//					label_15.setVisible(true);
//					chckbxSavePdfAttachment.setVisible(true);
//				} 
				else {
//					if (filetype.equalsIgnoreCase("pdf")) {
//						chckbxSavePdfAttachment.setVisible(true);
//						label_15.setVisible(false);
//					}
				}
			}
		});
		chckbxMigrateOrBackup.setFont(new Font("Tahoma", Font.BOLD, 10));

		chckbxMigrateOrBackup.setForeground(Color.RED);
		chckbxMigrateOrBackup.setBackground(Color.WHITE);

		label_17 = new JLabel("");
		label_17.setBounds(425, 140, 26, 23);
		label_17.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_17.setToolTipText(
				"Check the option, If you want to migrate or backup emails without their attachment files.");

		chckbx_convert_pdf_to_pdf = new JCheckBox("Convert Attachments to PDF Format");
		chckbx_convert_pdf_to_pdf.setBounds(6, 166, 414, 23);
		chckbx_convert_pdf_to_pdf.setRolloverEnabled(false);
		chckbx_convert_pdf_to_pdf.setRequestFocusEnabled(false);
		chckbx_convert_pdf_to_pdf.setOpaque(false);
		chckbx_convert_pdf_to_pdf.setFocusable(false);
		chckbx_convert_pdf_to_pdf.setFocusPainted(false);
		chckbx_convert_pdf_to_pdf.setContentAreaFilled(false);
		chckbx_convert_pdf_to_pdf.setFont(new Font("Tahoma", Font.BOLD, 10));
		chckbx_convert_pdf_to_pdf.setForeground(Color.RED);
		chckbx_convert_pdf_to_pdf.setBackground(Color.WHITE);
		chckbx_convert_pdf_to_pdf.setVisible(false);
		chckbx_convert_pdf_to_pdf.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (chckbxMigrateOrBackup.isSelected() && chckbx_convert_pdf_to_pdf.isSelected()) {

					JOptionPane.showMessageDialog(main_multiplefile.this,
							"Please select one of them checkbox before continuing.", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));

					chckbx_convert_pdf_to_pdf.setSelected(false);
				} else if (chckbxSavePdfAttachment.isSelected() && chckbxMigrateOrBackup.isSelected()) {

					JOptionPane.showMessageDialog(main_multiplefile.this,
							"Please select one of them checkbox before continuing.", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));

					chckbx_convert_pdf_to_pdf.setSelected(false);
				} else if (chckbx_convert_pdf_to_pdf.isSelected() && chckbxSavePdfAttachment.isSelected()) {

					JOptionPane.showMessageDialog(main_multiplefile.this,
							"Please select one of them checkbox before continuing.", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					chckbxSavePdfAttachment.setSelected(false);
				}
			}
		});

		label_pdf_to_pdf = new JLabel("");
		label_pdf_to_pdf.setBounds(425, 170, 41, 25);
		label_pdf_to_pdf.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_pdf_to_pdf.setToolTipText("It helps to convert the contained attachments of PDF files to PDF format.");

		mailbox = new JRadioButton("MailBox", true);
		mailbox.setBounds(20, 268, 76, 21);
		mailbox.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				filterselected = "mailboxsel";
			}
		});
		mailbox.setRolloverEnabled(false);
		mailbox.setRequestFocusEnabled(false);
		mailbox.setOpaque(false);
		mailbox.setFocusable(false);
		mailbox.setFocusPainted(false);
		mailbox.setContentAreaFilled(false);
		mailbox.setFont(new Font("Tahoma", Font.BOLD, 10));
		mailbox.setForeground(Color.RED);
		mailbox.setBackground(Color.WHITE);
		mailbox.setVisible(false);
		publicfolder = new JRadioButton("Public Folder");
		publicfolder.setBounds(114, 268, 90, 21);
		publicfolder.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				filterselected = "publicfoldersel";
			}
		});
		publicfolder.setRolloverEnabled(false);
		publicfolder.setRequestFocusEnabled(false);
		publicfolder.setOpaque(false);
		publicfolder.setFocusable(false);
		publicfolder.setFocusPainted(false);
		publicfolder.setContentAreaFilled(false);
		publicfolder.setFont(new Font("Tahoma", Font.BOLD, 10));
		publicfolder.setForeground(Color.RED);
		publicfolder.setBackground(Color.WHITE);
		publicfolder.setVisible(false);
		archive = new JRadioButton("Archive");
		archive.setBounds(240, 268, 64, 21);
		archive.setRolloverEnabled(false);
		archive.setRequestFocusEnabled(false);
		archive.setOpaque(false);
		archive.setFocusable(false);
		archive.setFocusPainted(false);
		archive.setContentAreaFilled(false);
		archive.setFont(new Font("Tahoma", Font.BOLD, 10));
		archive.setForeground(Color.RED);
		archive.setBackground(Color.WHITE);
		archive.setVisible(false);
		archive.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				filterselected = "archivesel";
			}
		});
		ButtonGroup g5 = new ButtonGroup();
		g5.add(mailbox);
		g5.add(publicfolder);
		g5.add(archive);

		panel_6 = new JPanel();
		panel_6.setBounds(547, 289, 510, 43);
		panel_6.setBackground(Color.WHITE);
		panel_6.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));

		JLabel lblNewLabel_10 = new JLabel("");
		lblNewLabel_10.setBounds(454, 10, 24, 25);
		lblNewLabel_10.setToolTipText("Add the name by which the folder will be created " + "\r\n"
				+ " it must not contain these characters :\\?/|*<>\t");
		lblNewLabel_10.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));

//		textField_customfolder = new JTextField();
		textField_customfolder.setBounds(179, 10, 251, 26);
		textField_customfolder.setEnabled(false);
//		textField_customfolder.setEditable(false);
		textField_customfolder.setBackground(Color.WHITE);
		textField_customfolder.addKeyListener(new KeyAdapter() {
			public void keyReleased(KeyEvent event) {

				String content = textField_customfolder.getText();
				if (content.contains(":") || content.contains(":") || content.contains("\\") || content.contains("?")
						|| content.contains("/") || content.contains("|") || content.contains("*")
						|| content.contains("<") || content.contains(">") || content.contains("\t")
						|| content.contains("//s") || content.contains("\"")) {
					textField_customfolder.setText(getRidOfIllegalFileNameCharacters(content).trim());
				}
			}
		});
		textField_customfolder.setEditable(false);
		textField_customfolder.setColumns(10);

		// chckbxCustomFolderName = new JCheckBox("Custom Folder Name");
		chckbxCustomFolderName.setBounds(6, 10, 167, 26);
		chckbxCustomFolderName.setRolloverEnabled(false);
		chckbxCustomFolderName.setRequestFocusEnabled(false);
		chckbxCustomFolderName.setOpaque(false);
		chckbxCustomFolderName.setSelected(false);
		chckbxCustomFolderName.setFocusable(false);
		chckbxCustomFolderName.setFocusPainted(false);
		chckbxCustomFolderName.setContentAreaFilled(false);
		chckbxCustomFolderName.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxCustomFolderName.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					textField_customfolder.setEditable(true);
					textField_customfolder.setEnabled(true);
				} else {
					textField_customfolder.setEditable(false);
					textField_customfolder.setEnabled(false);
					textField_customfolder.setText("");
				}

//				else {
//					textField_customfolder.setEditable(false);
//				}
			}
		});

		chckbxCustomFolderName.setBackground(Color.WHITE);

		panel_8 = new JPanel();
		panel_8.setBounds(547, 174, 510, 57);
		panel_8.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_8.setBackground(Color.WHITE);
		panel_8.setVisible(false);

		chckbxRestoreToDefault = new JCheckBox("Restore to Default Folder");
		chckbxRestoreToDefault.setBounds(6, 7, 202, 23);
		chckbxRestoreToDefault.setRolloverEnabled(false);
		chckbxRestoreToDefault.setRequestFocusEnabled(false);
		chckbxRestoreToDefault.setOpaque(false);
		chckbxRestoreToDefault.setFocusable(false);
		chckbxRestoreToDefault.setFocusPainted(false);
		chckbxRestoreToDefault.setContentAreaFilled(false);
		chckbxRestoreToDefault.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxRestoreToDefault.setBackground(Color.WHITE);

		chckbx_seperatepst = new JCheckBox("Seperate  PST");
		chckbx_seperatepst.setBounds(261, 7, 117, 23);
		chckbx_seperatepst.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbx_seperatepst.setBackground(Color.WHITE);

		date_filter = new JPanel();
		date_filter.setBounds(547, 11, 510, 157);
		date_filter.setBorder(new LineBorder(new Color(0, 0, 0)));
		date_filter.setBackground(Color.WHITE);
//		date_filter.setBorder(new TitledBorder(
//				new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
//				"                           ", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		date_filter.setBackground(Color.WHITE);

//		checkBox = new JCheckBox("New check box");

		lblNewLabel_6 = new JLabel("Start Date :");
		lblNewLabel_6.setBounds(14, 30, 68, 25);
		lblNewLabel_6.setFont(new Font("Tahoma", Font.BOLD, 11));

		dateChooser_NewFrom = new JDateChooser();
		dateChooser_NewFrom.setBounds(92, 35, 100, 25);
		dateChooser_NewFrom.getCalendarButton().setBackground(Color.WHITE);
		dateChooser_NewFrom.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_NewFrom.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_NewFrom.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_NewFrom.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_NewFrom.setBackground(Color.WHITE);
		dateChooser_NewFrom.setEnabled(false);
		dateChooser_NewFrom.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal2 = Calendar.getInstance();
				cal2.set(Calendar.HOUR_OF_DAY, 00);
				cal2.set(Calendar.MINUTE, 00);
				cal2.set(Calendar.SECOND, 00);
				Date startdate = cal2.getTime();
				dateChooser_NewFrom.setMaxSelectableDate(startdate);
			}
		});
		dateChooser_NewFrom.setDateFormatString("dd-MMM-yyyy");

		JLabel lblNewLabel_7 = new JLabel("End Date : ");
		lblNewLabel_7.setBounds(206, 30, 61, 25);
		lblNewLabel_7.setFont(new Font("Tahoma", Font.BOLD, 11));

		dateChooser_NewTo = new JDateChooser();
		dateChooser_NewTo.setBounds(277, 35, 100, 25);
		dateChooser_NewTo.getCalendarButton().setBackground(Color.WHITE);
		dateChooser_NewTo.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_NewTo.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_NewTo.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_NewTo.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_NewTo.setBackground(Color.WHITE);
		dateChooser_NewTo.setEnabled(false);
		dateChooser_NewTo.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				Calendar cal3 = Calendar.getInstance();
				cal3.set(Calendar.HOUR_OF_DAY, 23);
				cal3.set(Calendar.MINUTE, 59);
				cal3.set(Calendar.SECOND, 59);
				Date enddate = cal3.getTime();
				dateChooser_NewTo.setMaxSelectableDate(enddate);

				Calendar calendarstartdate = dateChooser_NewFrom.getCalendar();
				calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
				calendarstartdate.set(Calendar.MINUTE, 00);
				calendarstartdate.set(Calendar.SECOND, 00);
				dateChooser_NewTo.setMinSelectableDate(calendarstartdate.getTime());

			}
		});
		dateChooser_NewTo.setDateFormatString("dd-MMM-yyyy");

		JScrollPane scrollPane_5 = new JScrollPane();
		scrollPane_5.setBounds(16, 67, 357, 78);

		table_2 = new JTable();
		scrollPane_5.setViewportView(table_2);
		table_2.setModel(new DefaultTableModel(new Object[][] {}, new String[] { "From", "To" }));

		add = new JButton("");
		add.setBounds(391, 75, 90, 30);
		add.setEnabled(false);
		add.setToolTipText("Click here to Add date in table.");
		add.setRolloverEnabled(false);
		add.setRequestFocusEnabled(false);
		add.setOpaque(false);
		add.setFocusable(false);
		add.setFocusTraversalKeysEnabled(false);
		add.setFocusPainted(false);
		add.setDefaultCapable(false);
		add.setContentAreaFilled(false);
		add.setBorderPainted(false);
		add.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				add.setIcon(new ImageIcon(Main_Frame.class.getResource("/add-hvr.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				add.setIcon(new ImageIcon(Main_Frame.class.getResource("/add.png")));
			}
		});

		add.setIcon(new ImageIcon(Main_Frame.class.getResource("/add.png")));
		add.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				new Date();

				if (!fromdateeditor.getText().isEmpty() && !todateeditor.getText().isEmpty()) {
					from = fromdateeditor.getText();
					to = todateeditor.getText();
					String tempFrom, tempTo;
					boolean flag = true;

					if (dateChooser_NewTo.getDate().after(dateChooser_NewFrom.getDate())
							|| dateChooser_NewTo.getDate().equals(dateChooser_NewFrom.getDate())) {
						for (int i = 0; i < table_2.getModel().getRowCount(); i++) {
							tempFrom = table_2.getModel().getValueAt(i, 0) + "";
							tempTo = table_2.getModel().getValueAt(i, 1) + "";
							if (from.equals(tempFrom) && to.equals(tempTo))
								flag = false;
						}
						if (flag)
							((DefaultTableModel) table_2.getModel()).addRow(new Object[] { from, to });
						else
							JOptionPane.showMessageDialog(main_multiplefile.this,
									"Duplicate date cannot be added in the table ,Please add valid date",
									messageboxtitle, JOptionPane.INFORMATION_MESSAGE);
					} else {
						JOptionPane.showMessageDialog(main_multiplefile.this, "Invalid Date, Please add valid date",
								messageboxtitle, JOptionPane.INFORMATION_MESSAGE);
					}

				}

			}
		});
		todateeditor = (JTextFieldDateEditor) dateChooser_NewTo.getDateEditor();
		fromdateeditor = (JTextFieldDateEditor) dateChooser_NewFrom.getDateEditor();

		remove = new JButton("");
		remove.setBounds(391, 112, 90, 30);
		remove.setEnabled(false);
		remove.setToolTipText("Click here to Remove date from table.");
		remove.setRolloverEnabled(false);
		remove.setRequestFocusEnabled(false);
		remove.setOpaque(false);
		remove.setFocusable(false);
		remove.setFocusTraversalKeysEnabled(false);
		remove.setFocusPainted(false);
		remove.setDefaultCapable(false);
		remove.setContentAreaFilled(false);
		remove.setBorderPainted(false);
		remove.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				remove.setIcon(new ImageIcon(Main_Frame.class.getResource("/remov-hvr.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				remove.setIcon(new ImageIcon(Main_Frame.class.getResource("/remov.png")));
			}
		});

		remove.setIcon(new ImageIcon(Main_Frame.class.getResource("/remov.png")));
		remove.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {

				DefaultTableModel model = (DefaultTableModel) table_2.getModel();
				int[] rows = table_2.getSelectedRows();

				if (table_2.getRowCount() < 1) {
					JOptionPane.showMessageDialog(main_multiplefile.this, "Please add date to remove", messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));
					return;
				}
				if (rows.length < 1) {
					JOptionPane.showMessageDialog(main_multiplefile.this, "Please add date to remove", messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));
					return;
				}
				for (int i = 0; i < rows.length; i++) {
					model.removeRow(rows[i] - i);
				}

			}
		});

		lblnamingconvention_1 = new JLabel("Naming Convention");
		lblnamingconvention_1.setBounds(10, 13, 161, 26);
		panel_6.setLayout(null);
		panel_6.add(chckbxCustomFolderName);
		panel_6.add(textField_customfolder);
		panel_6.add(lblNewLabel_10);
		panel_3_1_1.setLayout(null);
		panel_3_1_1.add(panel_duplicacy);
		panel_duplicacy.setLayout(null);
		panel_duplicacy.add(chckbxRemoveDuplicacy);
		panel_duplicacy.add(label_12);
		panel_duplicacy.add(chckbxSaveInSame);
		panel_duplicacy.add(label_13);
		panel_duplicacy.add(chckbxMaintainFolderStructure);
		panel_duplicacy.add(label_14);
		panel_duplicacy.add(chckbxSavePdfAttachment);
		panel_duplicacy.add(label_15);
		panel_duplicacy.add(chckbxSaveMboxIn);
		panel_duplicacy.add(label_16);
		panel_duplicacy.add(chckbxMigrateOrBackup);
		panel_duplicacy.add(label_17);
		panel_duplicacy.add(chckbx_convert_pdf_to_pdf);
		panel_duplicacy.add(label_pdf_to_pdf);
		panel_duplicacy.add(chckbx_splitpst);
		panel_duplicacy.add(spinner_sizespinner);
		panel_duplicacy.add(comboBox_setsize);
		panel_duplicacy.add(lbl_splitpst);
		panel_duplicacy.add(mailbox);
		panel_duplicacy.add(publicfolder);
		panel_duplicacy.add(archive);
		panel_3_1_1.add(datefilter);
		panel_3_1_1.add(date_filter);
		panel_3_1_1.add(panel_8);
		panel_8.setLayout(null);
		panel_8.add(chckbxRestoreToDefault);
		panel_8.add(chckbx_seperatepst);
		panel_3_1_1.add(panel_mailfilter);
		panel_3_1_1.add(panel_taskfilter);
		panel_3_1_1.add(panel_5);
		panel_5.setLayout(null);
		panel_5.add(lblnamingconvention_1);
		panel_5.add(comboBox);
		panel_3_1_1.add(panel_6);
		date_filter.setLayout(null);
		date_filter.add(lblNewLabel_6);
		date_filter.add(dateChooser_NewFrom);
		date_filter.add(lblNewLabel_7);
		date_filter.add(dateChooser_NewTo);
		date_filter.add(scrollPane_5);
		date_filter.add(add);
		date_filter.add(remove);

		label_11 = new JLabel("");
		label_11.setBounds(924, 41, 76, 27);

		JLabel lblSavesbackupmigrateAs = new JLabel(" Saves/Backup/Migrate As ");
		lblSavesbackupmigrateAs.setBounds(10, 13, 189, 29);
		lblSavesbackupmigrateAs.setForeground(Color.BLUE);
		lblSavesbackupmigrateAs.setFont(new Font("Tahoma", Font.BOLD, 13));
		panel_3.add(lblSavesbackupmigrateAs);

		panel_12 = new JPanel();
		panel_12.setBounds(0, 574, 1071, 40);
		panel_12.setBackground(new Color(0, 0, 0));
		//panel_3.add(panel_12);

		btn_previous_p3 = new JButton("");
		btn_previous_p3.setBounds(725, 20, 111, 31);
		btn_previous_p3.setToolTipText("Click here to Go Back.");
		btn_previous_p3.setBorderPainted(false);
		btn_previous_p3.setContentAreaFilled(false);
		btn_previous_p3.setDefaultCapable(false);
		btn_previous_p3.setFocusable(false);
		btn_previous_p3.setFocusTraversalKeysEnabled(false);
		btn_previous_p3.setFocusPainted(false);
		btn_previous_p3.setRolloverEnabled(false);
		btn_previous_p3.setRequestFocusEnabled(false);
		btn_previous_p3.setOpaque(false);
		btn_previous_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_previous_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_previous_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
			}
		});

		btn_previous_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
		btn_previous_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				lblTotalMessageCount.setText("<html><b>" + "  Total Message Count : ");
				hm.clear();
				// lblNew_setemail.setText("");
				dateChooser_NewFrom.setDate(null);
				dateChooser_NewTo.setDate(null);
				datefilter.setSelected(false);
				dateChooser_NewFrom.setEnabled(false);
				dateChooser_NewTo.setEnabled(false);
				remove.setEnabled(false);
				add.setEnabled(false);
//								lblNew_setsubject.setText("");
//								label_date.setText("");
				editorPane.setText("");
				Stoppreview = false;
				try {
					DefaultTableModel model = (DefaultTableModel) table_fileinformation.getModel();

					while (model.getRowCount() > 0) {

						for (int i = 0; i < model.getRowCount(); ++i) {

							model.removeRow(i);
						}
					}
					DefaultTableModel model11 = (DefaultTableModel) table_2.getModel();
					while (model11.getRowCount() > 0) {

						for (int i = 0; i < model11.getRowCount(); ++i) {

							model11.removeRow(i);

						}
					}
					DefaultTableModel model1 = (DefaultTableModel) table_1.getModel();

					while (model1.getRowCount() > 0) {

						for (int i = 0; i < model1.getRowCount(); ++i) {

							model1.removeRow(i);
						}
					}

					if (output) {
						if (filetype.equals("OFFICE 365") || filetype.equals("Live Exchange")
								|| filetype.equals("Hotmail")) {
							clientforexchange_output.dispose();

						} else {

							iconnforimap_output.dispose();
						}
					}
				} catch (Exception e1) {

				}
				comboBox_fileDestination_type.setEnabled(true);
				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_2");
			}
		});
		btn_previous_p3.setFont(new Font("Tahoma", Font.BOLD, 12));

		// btn_converter_1 = new JButton("");
		// btn_converter_1.setBounds(848, 20, 111, 31);
		btn_converter_1 = new GradientButton("Migration");
		btn_converter_1.setEnabled(false);
		btn_converter_1.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {

				if (btn_converter_1.isEnabled()) {
					btn_converter_1.setGradientColor1(new Color(70, 130, 180));
					btn_converter_1.setGradientColor2(new Color(70, 130, 180));
					btn_converter_1.setForeground(new Color(255, 255, 255));
				}
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_converter_1.setGradientColor1(new Color(255, 255, 255));
				btn_converter_1.setGradientColor2(new Color(255, 255, 255));
				btn_converter_1.setForeground(new Color(80, 80, 80));
			}
		});

		btn_converter_1.setToolTipText("Click here to Start Conversion.");

		btn_converter_1.setRequestFocusEnabled(false);
		btn_converter_1.setOpaque(false);
		btn_converter_1.setFocusTraversalKeysEnabled(false);
		btn_converter_1.setFocusable(false);
		btn_converter_1.setFocusPainted(false);
		btn_converter_1.setContentAreaFilled(false);
		btn_converter_1.setBorderPainted(false);
		btn_converter_1.setDefaultCapable(false);
		btn_converter_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				fourth = true;

				Buttonclick("btn_converter_1", fourth);
				if (table.getRowCount() == 0) {
					JOptionPane.showMessageDialog(main_multiplefile.this,
							"Please Add File First then you can Migrate !", messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));
					return;

				}

				panel_progress.setVisible(true);
				panel_progress.setBorder(new LineBorder(new Color(0, 0, 0)));
				radioFileFormat.setEnabled(false);
				rdbtnEmailClients.setEnabled(false);
				count_destination = 0;
				generalFolderMap.clear();

				SimpleDateFormat df2 = new SimpleDateFormat("dd-MMM-yyyy HH:mm");

				fromList.clear();
				toList.clear();
				if (datefilter.isSelected()) {
					if (table_2.getModel().getRowCount() < 1) {
						JOptionPane.showMessageDialog(main_multiplefile.this, "Please Add date in the table!",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));
						return;
					} else {
						for (int i = 0; i < table_2.getModel().getRowCount(); i++) {
							Date frmdate, todate;
							try {

								frmdate = df2.parse(table_2.getValueAt(i, 0).toString() + " 00:00");
								System.out.println(frmdate);
								fromList.add(frmdate);
								todate = df2.parse(table_2.getValueAt(i, 1).toString() + " 23:59");
								System.out.println(todate);
								toList.add(todate);
							} catch (ParseException e1) {
								mf.logger.severe(e1.getMessage());
								mf.logger.warning("Exception : " + e1.getMessage() + System.lineSeparator());
								e1.printStackTrace();
							}
						}
					}
				}

				if (filetype.equalsIgnoreCase("pdf") && chckbxSavePdfAttachment.isSelected()) {
					chckbxSavePdfAttachment.setSelected(true);
					label_15.setVisible(true);
				} else if (filetype.equalsIgnoreCase("pdf") && (!chckbxSavePdfAttachment.isSelected())) {
					chckbxSavePdfAttachment.setSelected(false);
				}

				if (chckbx_calender_box.isSelected()) {
					try {
						Calendar calendarstartdate = dateChooser_calender_start.getCalendar();
						Calendar calendarenddate = dateChooser_calendar_end.getCalendar();
						calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
						calendarstartdate.set(Calendar.MINUTE, 00);
						calendarstartdate.set(Calendar.SECOND, 00);
						calendarenddate.set(Calendar.HOUR_OF_DAY, 23);
						calendarenddate.set(Calendar.MINUTE, 59);
						calendarenddate.set(Calendar.SECOND, 59);

						Calenderfilterstartdate = calendarstartdate.getTime();
						Calenderfilterenddate = calendarenddate.getTime();
					} catch (Exception e1) {
						JOptionPane.showMessageDialog(mf, "Please Enter the date in Calendar filter before Continuing.",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

						throw new NullPointerException();
					}
				}
				if (chckbx_Mail_Filter.isSelected()) {
					try {

						Calendar mailstartdate = dateChooser_mail_fromdate.getCalendar();
						Calendar mailenddate = dateChooser_mail_tilldate.getCalendar();
						mailstartdate.set(Calendar.HOUR_OF_DAY, 00);
						mailstartdate.set(Calendar.MINUTE, 00);
						mailstartdate.set(Calendar.SECOND, 00);
						mailenddate.set(Calendar.HOUR_OF_DAY, 23);
						mailenddate.set(Calendar.MINUTE, 59);
						mailenddate.set(Calendar.SECOND, 59);

						mailfilterstartdate = mailstartdate.getTime();
						mailfilterenddate = mailenddate.getTime();
					} catch (Exception e1) {
						JOptionPane.showMessageDialog(mf, "Please Enter the date in Mail filter before Continuing.",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

						throw new NullPointerException();
					}

				}
				if (task_box.isSelected()) {
					try {

						Calendar taskstartdate = dateChooser_task_start_date.getCalendar();
						Calendar taskenddate = dateChooser_task_end_date.getCalendar();
						taskstartdate.set(Calendar.HOUR_OF_DAY, 00);
						taskstartdate.set(Calendar.MINUTE, 00);
						taskstartdate.set(Calendar.SECOND, 00);
						taskenddate.set(Calendar.HOUR_OF_DAY, 23);
						taskenddate.set(Calendar.MINUTE, 59);
						taskenddate.set(Calendar.SECOND, 59);

						taskfilterstartdate = taskstartdate.getTime();
						taskfilterenddate = taskenddate.getTime();
					} catch (Exception e1) {
						JOptionPane.showMessageDialog(mf, "Please Enter the date in task filter before Continuing.",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

						throw new NullPointerException();
					}

				}
				if (chckbxCustomFolderName.isSelected() && textField_customfolder.getText().isEmpty()) {

					JOptionPane.showMessageDialog(mf, "Please Enter the name of folder before Continuing.",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));

					throw new NullPointerException();

				}
				Desktop desktop = Desktop.getDesktop();
				th = new Thread(new Runnable() {

					@Override
					public void run() {
						try {
							chckbxCustomFolderName.setEnabled(false);
							btn_Destination.setEnabled(false);
							btn_previous_p3.setEnabled(false);
							chckbxSaveInSame.setEnabled(false);
							textField_customfolder.setEditable(false);
							btnStop.setVisible(true);
							chckbxMigrateOrBackup.setEnabled(false);
							comboBox_fileDestination_type.setEnabled(false);
							btn_Destination.setEnabled(false);
							btn_previous_p3.setEnabled(false);
							lbl_progressreport.setText("");
							dateChooser_calender_start.setEnabled(false);
							chckbxRemoveDuplicacy.setEnabled(false);
							dateChooser_calendar_end.setEnabled(false);
							dateChooser_mail_fromdate.setEnabled(false);
							chckbxSaveMboxIn.setEnabled(false);
							dateChooser_mail_tilldate.setEnabled(false);
							dateChooser_task_start_date.setEnabled(false);
							chckbxRestoreToDefault.setEnabled(false);
							dateChooser_task_end_date.setEnabled(false);
							btn_signout_p3.setVisible(false);
							panel_5.setEnabled(false);
							btn_Destination.setEnabled(false);
							btn_previous_p3.setEnabled(false);
							btn_converter_1.setEnabled(false);
							
							btn_previous_p2.setEnabled(false);
							btn_Next.setEnabled(false);
							btn_next_pane2.setEnabled(false);
							btnNewButton_2.setEnabled(false);
							
							
							
							
							btnStop.setVisible(true);
							lbl_progressreport.setText("");
							chckbx_Mail_Filter.setEnabled(false);
							chckbx_calender_box.setEnabled(false);
							chckbx_convert_pdf_to_pdf.setEnabled(false);
							comboBox.setEnabled(false);
							chckbxSavePdfAttachment.setEnabled(false);
							task_box.setEnabled(false);
							long starttime = System.currentTimeMillis();
							chckbxMaintainFolderStructure.setEnabled(false);
							String destinationfile = "";
							String match = "";
							chckbx_splitpst.setEnabled(false);
							spinner_sizespinner.setEnabled(false);
							comboBox_setsize.setEnabled(false);
							spinner_sizespinner.updateUI();
							maxsize = 0;
							datefilter.setEnabled(false);
							textField_customfolder.setEnabled(false);
							archive.setEnabled(false);
							publicfolder.setEnabled(false);
							mailbox.setEnabled(false);
							if (chckbx_splitpst.isSelected()) {
								if (comboBox_setsize.getSelectedIndex() == 0) {
									Object o = spinner_sizespinner.getValue();
									Number n = (Number) o;
									maxsize = (n.longValue()) * (1000 * 1000);

								} else if (comboBox_setsize.getSelectedIndex() == 1) {
									Object o = spinner_sizespinner.getValue();
									Number n = (Number) o;
									maxsize = (n.longValue()) * (1000 * 1000 * 1000);
								}
							}
							Calendar cal = Calendar.getInstance();
//							Main_Frame.comboBox.setSelectedIndex(comboBox.getSelectedIndex());
							if (checkconvertagain) {

								if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")
										|| filetype.equalsIgnoreCase("YAHOO MAIL") || filetype.equalsIgnoreCase("AOL")
										|| filetype.equalsIgnoreCase("Amazon WorkMail")
										|| filetype.equalsIgnoreCase("Live Exchange")
										|| filetype.equalsIgnoreCase("Yandex Mail")
										|| filetype.equalsIgnoreCase("Icloud")
										|| filetype.equalsIgnoreCase("GoDaddy email")
										|| filetype.equalsIgnoreCase("Hostgator email")
										|| filetype.equalsIgnoreCase("Zoho Mail")
										|| filetype.equalsIgnoreCase("OFFICE 365")
										|| filetype.equalsIgnoreCase("Hotmail") || filetype.equalsIgnoreCase("IMAP"))) {

									if (chckbxCustomFolderName.isSelected()) {

										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);

										f = new File(tf_Destination_Location.getText() + File.separator + customerfolder
												+ filetype);
										if (!f.isFile()) {
											if (filetype.equalsIgnoreCase("Thunderbird")) {
												f = new File(tf_Destination_Location.getText() + File.separator
														+ customerfolder + ".sbd" + File.separator + fname + ".sbd");
												new MboxrdStorageWriter(tf_Destination_Location.getText()
														+ File.separator + customerfolder, false);
												f.mkdirs();
												new MboxrdStorageWriter(
														tf_Destination_Location.getText() + File.separator
																+ customerfolder + ".sbd" + File.separator + fname,
														false);

											}
											f.mkdirs();

											destination_path = f.getAbsolutePath();
										} else {

											f = new File(tf_Destination_Location.getText() + File.separator
													+ customerfolder + "(" + calendertime + ")" + filetype + "_"
													+ comboBox.getSelectedItem().toString());

											if (filetype.equalsIgnoreCase("Thunderbird")) {
												f = new File(tf_Destination_Location.getText() + File.separator
														+ customerfolder + "(" + calendertime + ")" + filetype + "_"
														+ comboBox.getSelectedItem().toString() + ".sbd"
														+ File.separator + fname + ".sbd");
												new MboxrdStorageWriter(tf_Destination_Location.getText()
														+ File.separator + customerfolder + "(" + calendertime + ")"
														+ filetype + "_" + comboBox.getSelectedItem().toString(),
														false);

											}

											f.mkdirs();

											destination_path = f.getAbsolutePath();

										}

									} else {
										fname = fileoptionm;
										f = new File(tf_Destination_Location.getText() + File.separator + calendertime
												+ File.separator + fname + filetype + "_"
												+ comboBox.getSelectedItem().toString());
										if (filetype.equalsIgnoreCase("Thunderbird")) {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ calendertime + ".sbd" + File.separator + fname + filetype + "_"
													+ comboBox.getSelectedItem().toString() + ".sbd" + File.separator
													+ fname + ".sbd");
											f.mkdirs();
											new MboxrdStorageWriter(
													tf_Destination_Location.getText() + File.separator + calendertime,
													false);

											new MboxrdStorageWriter(tf_Destination_Location.getText() + File.separator
													+ calendertime + File.separator + fname + filetype + "_"
													+ comboBox.getSelectedItem().toString(), false);

										}
										f.mkdirs();
										destination_path = f.getAbsolutePath();
										destinationfile = f.getAbsolutePath();
									}

								}

							} else {

								if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")
										|| filetype.equalsIgnoreCase("YAHOO MAIL") || filetype.equalsIgnoreCase("AOL")
										|| filetype.equalsIgnoreCase("Icloud")
										|| filetype.equalsIgnoreCase("GoDaddy email")
										|| filetype.equalsIgnoreCase("Hostgator email")
										|| filetype.equalsIgnoreCase("Amazon WorkMail")
										|| filetype.equalsIgnoreCase("Live Exchange")
										|| filetype.equalsIgnoreCase("Yandex Mail")
										|| filetype.equalsIgnoreCase("Zoho Mail")
										|| filetype.equalsIgnoreCase("OFFICE 365")
										|| filetype.equalsIgnoreCase("Hotmail") || filetype.equalsIgnoreCase("IMAP"))) {

									if (chckbxCustomFolderName.isSelected()) {

										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);

										f = new File(
												tf_Destination_Location.getText() + File.separator + customerfolder);
										if (!f.isFile()) {

											if (filetype.equalsIgnoreCase("Thunderbird")) {
												f = new File(tf_Destination_Location.getText() + File.separator
														+ customerfolder + ".sbd");
												new MboxrdStorageWriter(tf_Destination_Location.getText()
														+ File.separator + customerfolder, false);
												;

											}

											f.mkdirs();

											destination_path = f.getAbsolutePath();
										} else {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ customerfolder + "(" + calendertime + ")");
											if (filetype.equalsIgnoreCase("Thunderbird")) {
												f = new File(tf_Destination_Location.getText() + File.separator
														+ customerfolder + "(" + calendertime + ")" + ".sbd");
												new MboxrdStorageWriter(tf_Destination_Location.getText()
														+ File.separator + customerfolder + "(" + calendertime + ")",
														false);
											}
											f.mkdirs();

											destination_path = f.getAbsolutePath();

										}

									} else {

										fname = fileoptionm;

										f = new File(tf_Destination_Location.getText() + File.separator + calendertime
												+ File.separator + fname);
										if (filetype.equalsIgnoreCase("Thunderbird")) {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ calendertime + ".sbd" + File.separator + fname + ".sbd");
											f.mkdirs();
											new MboxrdStorageWriter(
													tf_Destination_Location.getText() + File.separator + calendertime,
													false);

											new MboxrdStorageWriter(tf_Destination_Location.getText() + File.separator
													+ calendertime + ".sbd" + File.separator + fname, false);
										}
										f.mkdirs();
										destination_path = f.getAbsolutePath();
										destinationfile = f.getAbsolutePath();
									}
								}

							}

							if (chckbx_Mail_Filter.isSelected()
									&& (mailfilterenddate == null || mailfilterstartdate == null)) {
								JOptionPane.showMessageDialog(mf, "Please Select Start and End Date.", messageboxtitle,
										JOptionPane.ERROR_MESSAGE,
										new ImageIcon(Main_Frame.class.getResource("/information.png")));
							} else if (chckbx_calender_box.isSelected()
									&& (Calenderfilterenddate == null || Calenderfilterstartdate == null)) {
								JOptionPane.showMessageDialog(mf, "Please Select Start and End Date.", messageboxtitle,
										JOptionPane.ERROR_MESSAGE,
										new ImageIcon(Main_Frame.class.getResource("/information.png")));
							} else {

								if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")
										|| filetype.equalsIgnoreCase("Amazon WorkMail")
										|| filetype.equalsIgnoreCase("YAHOO MAIL") || filetype.equalsIgnoreCase("AOL")
										|| filetype.equalsIgnoreCase("Yandex Mail")
										|| filetype.equalsIgnoreCase("Icloud")
										|| filetype.equalsIgnoreCase("GoDaddy email")
										|| filetype.equalsIgnoreCase("Zoho Mail")) {
									if (filetype.equalsIgnoreCase("GoDaddy email")) {
										calendertime = calendertime.replaceAll("[^a-zA-Z0-9]", "");
										fname = fname.replaceAll("[^a-zA-Z0-9]", "");

									}
									if (!chckbxRestoreToDefault.isSelected()) {
										if (chckbxCustomFolderName.isSelected()) {
											String customerfolder = textField_customfolder.getText().replace("//s", "");

											customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
											if (filetype.equalsIgnoreCase("GoDaddy email")) {
												customerfolder = customerfolder.replaceAll("[^a-zA-Z0-9]", "");

											}
											try {
												try {
													clientforimap_output.createFolder(iconnforimap_output,
															customerfolder);
												} catch (Exception e2) {
													connectionHandle1();
													clientforimap_output.createFolder(iconnforimap_output,
															customerfolder);
												}

												clientforimap_output.selectFolder(iconnforimap_output, customerfolder);
												path = customerfolder;

											} catch (Exception e) {

												clientforimap_output.createFolder(iconnforimap_output,
														customerfolder + "(" + calendertime + ")");

												clientforimap_output.selectFolder(iconnforimap_output,
														customerfolder + "(" + calendertime + ")");
												path = customerfolder + "(" + calendertime + ")";

											}
										} else {

											calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
											try {
												clientforimap_output.createFolder(iconnforimap_output, calendertime);
											} catch (Exception e2) {
												connectionHandle1();
												clientforimap_output.createFolder(iconnforimap_output, calendertime);
											}

											clientforimap_output.selectFolder(iconnforimap_output, calendertime);

											path = calendertime;

										}
									}
								} else if (filetype.equalsIgnoreCase("IMAP")
										|| filetype.equalsIgnoreCase("Hostgator email")) {
									System.out.println("bb");
									String sepretor = clientforimap_output.getDelimiter();
									if (chckbxCustomFolderName.isSelected()) {
										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);

										try {

											try {
//												clientforimap_output.createFolder(iconnforimap_output,
//														"INBOX" + "." + customerfolder);
												clientforimap_output.createFolder(iconnforimap_output,
														"INBOX" + sepretor + customerfolder);
											} catch (Exception e2) {
												connectionHandle1();
//												clientforimap_output.createFolder(iconnforimap_output,
//														"INBOX" + "." + customerfolder);
												clientforimap_output.createFolder(iconnforimap_output,
														"INBOX" + sepretor + customerfolder);
											}

//											clientforimap_output.selectFolder(iconnforimap_output,
//													"INBOX" + "." + customerfolder);
//											path = "INBOX" + "." + customerfolder;
											clientforimap_output.selectFolder(iconnforimap_output,
													"INBOX" + sepretor + customerfolder);
											path = "INBOX" + sepretor + customerfolder;

										} catch (Exception e) {

//											clientforimap_output.createFolder(iconnforimap_output,
//													"INBOX" + "." + customerfolder + "(" + calendertime + ")");
//											clientforimap_output.selectFolder(iconnforimap_output,
//													"INBOX" + "." + customerfolder + "(" + calendertime + ")");
//											path = "INBOX" + "." + customerfolder + "(" + calendertime + ")";
											clientforimap_output.createFolder(iconnforimap_output,
													"INBOX" + sepretor + customerfolder + "(" + calendertime + ")");
											clientforimap_output.selectFolder(iconnforimap_output,
													"INBOX" + sepretor + customerfolder + "(" + calendertime + ")");
											path = "INBOX" + sepretor + customerfolder + "(" + calendertime + ")";
										}
									} else {

										calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());

										try {
//											clientforimap_output.createFolder(iconnforimap_output,
//													"INBOX" + "." + calendertime);
											clientforimap_output.createFolder(iconnforimap_output,
													"INBOX" + sepretor + calendertime);
										} catch (Exception e2) {
											connectionHandle1();
//											clientforimap_output.createFolder(iconnforimap_output,
//													"INBOX" + "." + calendertime);
											clientforimap_output.createFolder(iconnforimap_output,
													"INBOX" + sepretor + calendertime);
										}

//										clientforimap_output.selectFolder(iconnforimap_output,
//												"INBOX" + "." + calendertime);
//										path = "INBOX" + "." + calendertime;
										clientforimap_output.selectFolder(iconnforimap_output,
												"INBOX" + sepretor + calendertime);
										path = "INBOX" + sepretor + calendertime;

									}

								} else if (filetype.equalsIgnoreCase("Live Exchange")
										|| filetype.equalsIgnoreCase("Hotmail")) {
									if (!chckbxRestoreToDefault.isSelected()) {
										if (chckbxCustomFolderName.isSelected()) {
											String customerfolder = textField_customfolder.getText().replace("//s", "");

											customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
											try {
//
												Folderuri = clientforexchange_output.createFolder(customerfolder)
														.getUri();
												fa = Folderuri;

											} catch (Exception e) {

												Folderuri = clientforexchange_output
														.createFolder(customerfolder + "(" + calendertime + ")")
														.getUri();
												fa = Folderuri;
											}
										} else {
											Folderuri = clientforexchange_output.createFolder(calendertime).getUri();
											fa = Folderuri;

										}
									}
								} else if (filetype.equalsIgnoreCase("Office 365")) {
									String customerfolder = textField_customfolder.getText().replace("//s", "");

									customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
								}
								String finalpath = destination_path;
								Progressbar.setVisible(true);
								long totalcount = 0;
								mf.logger.info("Convertion into " + fileoptionm + "Start Time : " + cal.getTime()
										+ System.lineSeparator());
								Main_Frame.count_destination = 0;
								int filecounter = 1;
								for (int i = 0; i < filesfin.length; i++) {

									if (stop) {
										break;
									}
									count_destination = 0;
									file = new File(filesfin[i].replace("<html><b>", ""));
									destination_path = finalpath;

									fname = file.getName();

									mf.logger.info("file name " + fname + "Start Time : " + cal.getTime()
											+ System.lineSeparator());
									if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")
											|| filetype.equalsIgnoreCase("Yandex Mail")
											|| filetype.equalsIgnoreCase("Zoho Mail")
											|| filetype.equalsIgnoreCase("Icloud")
											|| filetype.equalsIgnoreCase("GoDaddy email")
											|| filetype.equalsIgnoreCase("Hostgator email")
											|| filetype.equalsIgnoreCase("Amazon WorkMail")
											|| filetype.equalsIgnoreCase("YAHOO MAIL")
											|| filetype.equalsIgnoreCase("AOL") || filetype.equalsIgnoreCase("IMAP"))) {
										path = "";
									}

									if (table.getValueAt(i, 3).toString().replace("<html><b>", "")
											.equalsIgnoreCase("File")) {
										int fileCount = 1;
										for (Map.Entry<String, List<String>> entry : hm.entrySet()) {
											if (stop) {
												break;
											}
											if (entry.getKey().trim()
													.equalsIgnoreCase(filesfin[i].replace("<html><b>", ""))) {

												if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
														|| fileoptionm
																.equalsIgnoreCase("Exchange Offline Storage (.ost)")
														|| fileoptionm.equalsIgnoreCase("OLM File (.olm)")) {
													pstfolderlist = new ArrayList<String>();

													pstfolderlist.addAll(entry.getValue());
												}

												for (int k = 0; k < pstfolderlist.size(); k++) {

													System.out.println(pstfolderlist.get(k));
												}

												filepath = file.getAbsolutePath();

												fname = file.getName().replace(".mbx", "").replace(".mbox", "")
														.replace(".pst", "").replace(".ost", "").replace(".nsf", "")
														.replace(".eml", "").replace(".olm", "");
												fname = getRidOfIllegalFileNameCharacters(fname);
												fname = fname.trim();
												destination_path = destination_path + File.separator + fname;

												if (filetype.equalsIgnoreCase("Thunderbird")) {
													destination_path = destination_path + ".sbd";

												}

												if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")
														|| fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
													splitcount = 0;
													if (filetype.equalsIgnoreCase("EML")
															|| filetype.equalsIgnoreCase("MSG")
															|| filetype.equalsIgnoreCase("EMLX")
															|| filetype.equalsIgnoreCase("HTML")
															|| filetype.equalsIgnoreCase("MHTML")) {
														file = new File(destination_path);
														if (file.exists()) {
															destination_path = destination_path + "_" + fileCount;
															fileCount++;
														}
														file = new File(destination_path);
														file.mkdirs();
														ConvertPSTOST_file mf1 = new ConvertPSTOST_file(mf, filetype,
																destination_path, count_destination, filepath,
																main_multiplefile.this, pstfolderlist, fromList,
																toList);
														Thread saveTh = new Thread(mf1);
														saveTh.start();
														saveTh.join();
													} else if (filetype.equalsIgnoreCase("VCF")
															|| filetype.equalsIgnoreCase("ICS")) {
														file = new File(destination_path);
														if (file.exists()) {
															destination_path = destination_path + "_" + fileCount;
															fileCount++;
														}
														file = new File(destination_path);
														file.mkdirs();
														ConvertPSTOST_vcfics mf5 = new ConvertPSTOST_vcfics(mf,
																filetype, destination_path, count_destination, filepath,
																main_multiplefile.this, pstfolderlist, fromList, toList,
																temppathm);

														Thread saveTh = new Thread(mf5);
														saveTh.start();
														saveTh.join();

													} else if (filetype.equalsIgnoreCase("Live Exchange")
															|| filetype.equalsIgnoreCase("HOTMAIL")) {
														label_11.setIcon(new ImageIcon(
																Main_Frame.class.getResource("/download.png")));

														if (chckbxRestoreToDefault.isSelected()) {
															ConvertPST_defaultrestore mf1 = new ConvertPST_defaultrestore(
																	mf, filetype, destination_path, count_destination,
																	filepath, main_multiplefile.this, pstfolderlist,
																	fromList, toList, temppathm);

															Thread saveTh = new Thread(mf1);
															saveTh.start();
															saveTh.join();

//																									
														} else {
															ConvertPSTOST_365 mf6 = new ConvertPSTOST_365(mf, filetype,
																	destination_path, count_destination, filepath,
																	main_multiplefile.this, pstfolderlist, fromList,
																	toList, temppathm, username_p3, password_p3,
																	Folderuri, clientforexchange_output, fa, fname);

															Thread saveTh = new Thread(mf6);
															saveTh.start();
															saveTh.join();
//                                                                                                    
														}

													} else if (filetype.equalsIgnoreCase("OFFICE 365")) {
														officeEws365(filecounter);
														filecounter++;
														map.clear();
														path = "";
													} else if (filetype.equalsIgnoreCase("IMAP")
															|| filetype.equalsIgnoreCase("Hostgator email")) {
														label_11.setIcon(new ImageIcon(
																Main_Frame.class.getResource("/download.png")));
														match = path;
														path4 = path;
														path = path + "." + fname;

														clientforimap_output.createFolder(iconnforimap_output, path);
														clientforimap_output.selectFolder(iconnforimap_output, path);
														ConvertPSTOST_imap();
														path = match;

//														ConvertPSTOST_imap mf7 = new ConvertPSTOST_imap(mf, filetype,
//																destination_path, count_destination, filepath,
//																main_multiplefile.this, pstfolderlist, fromList, toList,
//																temppathm, username_p3, password_p3, Folderuri,
//																clientforimap_output, iconnforimap_output, path,
//																domain_p3, portnofiletype, fa, fname, match);
//														Thread saveTh = new Thread(mf7);
//														saveTh.start();
//														saveTh.join();
//														path = main_multiplefile.match;

													} else if (filetype.equalsIgnoreCase("YAHOO MAIL")
															|| filetype.equalsIgnoreCase("AOL")
															|| filetype.equalsIgnoreCase("Amazon WorkMail")
															|| filetype.equalsIgnoreCase("GMAIL")
															|| filetype.equalsIgnoreCase("G-SUITE")
															|| filetype.equalsIgnoreCase("Icloud")
															|| filetype.equalsIgnoreCase("GoDaddy email")
															|| filetype.equalsIgnoreCase("Yandex Mail")
															|| filetype.equalsIgnoreCase("Zoho Mail")) {
														label_11.setIcon(new ImageIcon(
																Main_Frame.class.getResource("/download.png")));

														System.out.println("Start from here .");
//														Gmail_Folder mf8 = new Gmail_Folder(mf, filetype,
//																destination_path, count_destination, filepath,
//																main_multiplefile.this, pstfolderlist, fromList, toList,
//																temppathm, username_p3, password_p3, Folderuri,
//																clientforimap_output, iconnforimap_output, path, fname,
//																match, domain_p3, portnofiletype);
//														Thread saveTh = new Thread(mf8);
//														saveTh.start();
//														saveTh.join();
//														path = path.replace("/" + fname, "");

														// Its working normally
														match = path;
														if (filetype.equalsIgnoreCase("GoDaddy email")) {
															fname = fname.replaceAll("[^a-zA-Z0-9]", "");

														}
														path4 = path;
														path = path + "/" + fname;

														if (clientforimap_output.existFolder(path)) {
															clientforimap_output.selectFolder(iconnforimap_output,
																	path);
														} else {
															clientforimap_output.createFolder(iconnforimap_output,
																	path);
															clientforimap_output.selectFolder(iconnforimap_output,
																	path);
														}

														ConvertPSTOST_gmail();
														path = path.replace("/" + fname, "");
														path = match;

														//
													} else if (filetype.equalsIgnoreCase("CSV")) {
														file = new File(destination_path);
														if (file.exists()) {
															destination_path = destination_path + "_" + fileCount;
															fileCount++;
														}
														file = new File(destination_path);
														file.mkdirs();
														ConvertPSTOST_csv mf5 = new ConvertPSTOST_csv(mf, filetype,
																destination_path, count_destination, filepath,
																main_multiplefile.this, pstfolderlist, fromList, toList,
																temppathm);
														Thread saveTh = new Thread(mf5);
														saveTh.start();
														saveTh.join();

													} else if (filetype.equalsIgnoreCase("PST")) {
														file = new File(destination_path);
														if (file.exists()) {
															destination_path = destination_path + "_" + fileCount;
															fileCount++;
														}
														file = new File(destination_path);
														file.mkdirs();
														ConvertOST_PST mf2 = new ConvertOST_PST(mf, filetype,
																destination_path, count_destination, filepath,
																main_multiplefile.this, pstfolderlist, fromList, toList,
																maxsize);
														Thread saveTh = new Thread(mf2);
														saveTh.start();
														saveTh.join();
													} else if (filetype.equalsIgnoreCase("MBOX")
															|| filetype.equalsIgnoreCase("Thunderbird")
															|| filetype.equalsIgnoreCase("Opera Mail")) {
//														new File(destination_path).mkdirs();
														file = new File(destination_path);
														if (file.exists()) {
															destination_path = destination_path + "_" + fileCount;
															fileCount++;
														}
														file = new File(destination_path);
														file.mkdirs();
														if (filetype.equalsIgnoreCase("Thunderbird")) {

															new MboxrdStorageWriter(
																	destination_path.replace(fname + ".sbd", "")
																			+ fname,
																	false);
														}
														ConvertPSTOST_mbox mf4 = new ConvertPSTOST_mbox(mf, filetype,
																destination_path, count_destination, filepath,
																main_multiplefile.this, pstfolderlist, fromList, toList,
																temppathm);
														Thread saveTh = new Thread(mf4);
														saveTh.start();
														saveTh.join();
													} else {
														file = new File(destination_path);
														if (file.exists()) {
															destination_path = destination_path + "_" + fileCount;
															fileCount++;
														}
														file = new File(destination_path);
														file.mkdirs();
														ConvertPSTOST_word mf3 = new ConvertPSTOST_word(mf, filetype,
																destination_path, count_destination, filepath,
																main_multiplefile.this, pstfolderlist, fromList,
																toList);
														Thread saveTh = new Thread(mf3);
														saveTh.start();
														saveTh.join();
													}

												}
//												destination_path = f.getAbsolutePath();
												destination_path = destinationfile;
											}
										}
										//
										if (filetype.equalsIgnoreCase("YAHOO MAIL")
												|| filetype.equalsIgnoreCase("GMAIL")
												|| filetype.equalsIgnoreCase("OFFICE 365") || filetype.equals("AOL")
												|| filetype.equalsIgnoreCase("Hotmail")
												|| filetype.equalsIgnoreCase("Icloud")
												|| filetype.equalsIgnoreCase("GoDaddy email")
												|| filetype.equalsIgnoreCase("Hostgator email")
												|| filetype.equalsIgnoreCase("Yandex Mail")
												|| filetype.equalsIgnoreCase("Amazon WorkMail")
												|| filetype.equalsIgnoreCase("Zoho Mail")
												|| filetype.equalsIgnoreCase("Live Exchange")
												|| filetype.equalsIgnoreCase("IMAP")
												|| filetype.equalsIgnoreCase("G-SUITE")) {

											if (filetype.equalsIgnoreCase("YAHOO MAIL")) {
												destination_path = "http://login.yahoo.com";
											} else if (filetype.equalsIgnoreCase("GMAIL")
													|| filetype.equalsIgnoreCase("G-SUITE")) {
												destination_path = "https://mail.google.com";
											} else if (filetype.equals("AOL")) {
												destination_path = "https://login.aol.com";
											} else if (filetype.equalsIgnoreCase("Zoho Mail")) {

												destination_path = "https://accounts.zoho.in/signin?servicename=VirtualOffice&signupurl=https://www.zoho.in/mail/zohomail-pricing.html&serviceurl=https://mail.zoho.in";

											} else if (filetype.equalsIgnoreCase("Yandex Mail")) {

												destination_path = "https://mail.yandex.com/?uid=1213147137#tabs/relevant";

											} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {

												destination_path = "Amazon WorkMail";
											} else if (filetype.equalsIgnoreCase("GoDaddy email")) {

												destination_path = "https://sso.godaddy.com/login?app=email&realm=pass";

											} else if (filetype.equalsIgnoreCase("Icloud")) {

												destination_path = "https://www.icloud.com/mail";

											} else if (filetype.equalsIgnoreCase("Hostgator email")) {

												destination_path = "https://www.hostgator.in/login.php";

											} else if (filetype.equals("IMAP")) {
												destination_path = "IMAP";
											} else if (filetype.equals("Hotmail")) {
												destination_path = "https://outlook.live.com";
											} else if (filetype.equals("Live Exchange")) {
												destination_path = "Live Exchange";
											} else {
												destination_path = "https://outlook.office365.com";
											}
										} else {
											destination_path = f.getAbsolutePath();
										}
									}
									String duration = Duration(starttime).toString();
									mode = (DefaultTableModel) table_fileConvertionreport_panel4.getModel();
									mode.addRow(new Object[] { fileoptionm, filetype, fname, Status, duration,
											count_destination, destination_path });
									destination_path = destination_path.replace(File.separator + fname, "");
									totalcount = totalcount + count_destination;
									mf.logger.info("File Saved " + count_destination + System.lineSeparator()
											+ "End Time : " + cal.getTime() + System.lineSeparator()
											+ "**********************************************************");

								}

								mode = (DefaultTableModel) table_fileConvertionreport_panel4.getModel();

								mode.addRow(new Object[] { "Total Message", "", "", "", "", totalcount, "" });
								mf.logger.info("File Saved " + totalcount + System.lineSeparator() + "End Time : "
										+ cal.getTime() + System.lineSeparator()
										+ "*****************************END*****************************");

							}

							destination_path = destination_path.replace(File.separator + fname, "");
							if (filetype.equalsIgnoreCase("THUNDERBIRD")) {
								JOptionPane.showMessageDialog(mf,
										"Please open the converted file from " + destination_path + " Thunderbird",
										messageboxtitle, JOptionPane.INFORMATION_MESSAGE);
							}
							if (filetype.equalsIgnoreCase("YAHOO MAIL") || filetype.equalsIgnoreCase("GMAIL")
									|| filetype.equalsIgnoreCase("G-SUITE") || filetype.equalsIgnoreCase("OFFICE 365")
									|| filetype.equals("AOL") || filetype.equalsIgnoreCase("Live Exchange")
									|| filetype.equalsIgnoreCase("Zoho Mail") || filetype.equalsIgnoreCase("Icloud")
									|| filetype.equalsIgnoreCase("GoDaddy email")
									|| filetype.equalsIgnoreCase("Hostgator email")
									|| filetype.equalsIgnoreCase("Amazon WorkMail")
									|| filetype.equalsIgnoreCase("Yandex Mail") || filetype.equalsIgnoreCase("hotmail")
									|| filetype.equalsIgnoreCase("IMAP")) {
								if (Desktop.isDesktopSupported()
										&& Desktop.getDesktop().isSupported(Desktop.Action.BROWSE)) {

									if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

										openBrowser(
												"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 ");
										reportpath = "http://login.yahoo.com";
									} else if (filetype.equalsIgnoreCase("Zoho Mail")) {

										reportpath = "https://accounts.zoho.in/signin?servicename=VirtualOffice&signupurl=https://www.zoho.in/mail/zohomail-pricing.html&serviceurl=https://mail.zoho.in";
										openBrowser(reportpath);
									} else if (filetype.equalsIgnoreCase("GoDaddy email")) {

										reportpath = "https://sso.godaddy.com/login?app=email&realm=pass";
										openBrowser(reportpath);
									} else if (filetype.equalsIgnoreCase("Icloud")) {

										reportpath = "https://www.icloud.com/mail";
										openBrowser(reportpath);
									} else if (filetype.equalsIgnoreCase("Hostgator email")) {

										JOptionPane.showMessageDialog(mf,
												"Please open the converted file from Hostgator", messageboxtitle,
												JOptionPane.INFORMATION_MESSAGE);
									} else if (filetype.equalsIgnoreCase("GMAIL")
											|| filetype.equalsIgnoreCase("G-SUITE")) {

										openBrowser("https://mail.google.com");
										reportpath = "https://mail.google.com";
									} else if (filetype.equalsIgnoreCase("Yandex Mail")) {

										reportpath = "https://mail.yandex.com/?uid=1213147137#tabs/relevant";
										openBrowser(reportpath);
									} else if (filetype.equals("AOL")) {

										openBrowser("https://login.aol.com");
										reportpath = "https://login.aol.com";
									} else if (filetype.equalsIgnoreCase("Live Exchange")) {

										JOptionPane.showMessageDialog(mf,
												"Please open the converted file from Live Exchange", messageboxtitle,
												JOptionPane.INFORMATION_MESSAGE);
									} else if (filetype.equalsIgnoreCase("Hotmail")) {

										openBrowser("https://outlook.live.com");
										reportpath = "https://outlook.live.com";
									} else if (filetype.equalsIgnoreCase("IMAP")) {

										JOptionPane.showMessageDialog(mf, "Please open the converted file from  IMAP",
												messageboxtitle, JOptionPane.INFORMATION_MESSAGE);
									} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {

										JOptionPane.showMessageDialog(mf,
												"Please open the converted file from  Amazon WorkMail", messageboxtitle,
												JOptionPane.INFORMATION_MESSAGE);
									} else {

										openBrowser("https://outlook.office365.com");
										reportpath = "https://outlook.office365.com";
									}
								}

							} else {

								desktop.open(f);
								reportpath = f.getAbsolutePath();

							}

						} catch (Exception e) {
							e.printStackTrace();
						} finally {
							Progressbar.setVisible(false);
							CardLayout card = (CardLayout) Cardlayout.getLayout();
							card.show(Cardlayout, "panel_4");
							
//							btn_previous_p2.setEnabled(true);
//							btn_Next.setEnabled(true);
//							btn_next_pane2.setEnabled(false);
//							btn_converter_1.setEnabled(false);
							btnNewButton_2.setEnabled(true);
							count_eml_msg_emlx = 0;

							JOptionPane.showMessageDialog(main_multiplefile.this,
									"Process has been successfully completed.", messageboxtitle,
									JOptionPane.ERROR_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/information.png")));
							// progressBar_message_p3.setVisible(false);
							try {
								if (filetype.equalsIgnoreCase("MBOX")) {
									if (!(wr == null)) {
										wr.dispose();
									}
								} else if (filetype.equalsIgnoreCase("OST") || filetype.equalsIgnoreCase("PST")) {
									if (!(pst == null)) {
										pst.dispose();

									}

								} else if (filetype.equalsIgnoreCase("CSV")) {
									if (!(writer == null)) {
										try {
											writer.close();
										} catch (IOException e) {

											e.printStackTrace();
										}
									}
								}
								
								/////////////////////
								if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")
										|| filetype.equalsIgnoreCase("OFFICE 365"))) {

									System.out.println("here we riched");
									textField_username_p3.setEnabled(true);
									passwordField_p3.setEnabled(true);
									tf_portNo_p3.setEnabled(true);
									chckbxShowPassword_p3.setEnabled(true);
									lblPortNo.setEnabled(true);
									lblNewLabel_5.setEnabled(true);
									lblNewLabel_1.setEnabled(true);
									lblNewLabel.setEnabled(true);
									lblemailAddress.setEnabled(true);
									lblemailAddress.setEnabled(false);

								}
								textField_customfolder.setEnabled(true);
								archive.setEnabled(true);
								publicfolder.setEnabled(true);
								mailbox.setEnabled(true);
								modern_Authentication.setEnabled(true);
								basic_Authentication.setEnabled(true);

//				                           basic_Authentication.setSelected(true);
								filetype = "";
								path = "";
								checkconvertagain = true;
								stop = false;
								btnStop.setVisible(false);
								chckbxSavePdfAttachment.setEnabled(true);
								btn_Destination.setEnabled(true);
								btn_previous_p3.setEnabled(true);
								chckbxSaveInSame.setEnabled(true);
								comboBox_fileDestination_type.setEnabled(true);
								btn_Destination.setEnabled(true);
								btn_previous_p3.setEnabled(true);
								lbl_progressreport.setText("");
								comboBox.setEnabled(true);
								chckbx_splitpst.setEnabled(true);
								textField_customfolder.setEditable(true);
								chckbxMigrateOrBackup.setEnabled(true);
								dateChooser_calender_start.setEnabled(true);
								chckbx_convert_pdf_to_pdf.setEnabled(true);
								chckbxRemoveDuplicacy.setEnabled(true);
								datefilter.setEnabled(true);
								dateChooser_calendar_end.setEnabled(true);
								chckbxCustomFolderName.setEnabled(true);
								btn_Destination.setEnabled(true);
								btn_previous_p3.setEnabled(true);
								checkmboxpstost = true;
								chckbxRestoreToDefault.setEnabled(true);
								chckbxSaveMboxIn.setEnabled(true);
								btn_signout_p3.setVisible(false);
								label_11.setVisible(false);
								panel_5.setEnabled(true);
								chckbx_Mail_Filter.setSelected(false);
								chckbx_Mail_Filter.setEnabled(true);
								chckbx_calender_box.setEnabled(true);
								task_box.setEnabled(true);
								chckbxMaintainFolderStructure.setEnabled(true);
								textField_domain_name_p3.setEnabled(true);
								textField_username_p3.setEnabled(true);
								passwordField_p3.setEnabled(true);
								tf_portNo_p3.setEnabled(true);
								lblNewLabel.setEnabled(true);
								lblemailAddress.setEnabled(true);
								lblNewLabel_1.setEnabled(true);
								lblNewLabel_5.setEnabled(true);
								lblPortNo.setEnabled(true);
								
								btn_previous_p2.setEnabled(true);
								btn_Next.setEnabled(true);
								
//								CardLayout card = (CardLayout) Cardlayout.getLayout();
//								card.show(Cardlayout, "panel_2");
								
								
								
								
								
							} catch (Exception e) {

							}
							checky = true;
							radioFileFormat.setEnabled(true);
							rdbtnEmailClients.setEnabled(true);
						}
					}

				});

				th.start();

			}
		});
		btn_converter_1.setFont(new Font("Tahoma", Font.BOLD, 12));

		radioFileFormat = new JRadioButton("File Formats");
		radioFileFormat.setBounds(205, 14, 116, 25);
		buttonGroup_1.add(radioFileFormat);
		radioFileFormat.setSelected(true);
//												 panel_3.add(radioFileFormat);

		radioFileFormat.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				comboBox_fileDestination_type.removeAllItems();

				for (int i = 0; i < Main_Frame.file_sfd.length; i++) {
					comboBox_fileDestination_type.addItem(Main_Frame.file_sfd[i]);
				}
			}
		});
		radioFileFormat.setFont(new Font("Tahoma", Font.BOLD, 14));
		radioFileFormat.setBackground(new Color(255, 255, 255));
		panel_3.add(radioFileFormat);

		rdbtnEmailClients = new JRadioButton("Email Clients");
		rdbtnEmailClients.setBounds(342, 13, 136, 25);
		buttonGroup_1.add(rdbtnEmailClients);
		rdbtnEmailClients.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				comboBox_fileDestination_type.removeAllItems();
				rdbtnEmailClients.setSelected(true);

				for (int i = 0; i < Main_Frame.email_sfd.length; i++) {
					comboBox_fileDestination_type.addItem(Main_Frame.email_sfd[i]);
				}
			}
		});
		rdbtnEmailClients.setFont(new Font("Tahoma", Font.BOLD, 14));
		rdbtnEmailClients.setBackground(new Color(255, 255, 255));
		panel_3.add(rdbtnEmailClients);
		panel_2.setLayout(null);
		panel_2.add(scrollPane_1);
		panel_2.add(scrollPane_2);
		panel_2.add(lblTotalMessageCount);
		// panel_2.add(btnViewer);
		// panel_2.add(btnAttachment);
		panel_2.add(innercardlayout);
		
		searchField = new JTextField();
		searchField.setBackground(UIManager.getColor("Button.light"));
		searchField.setFont(new Font("Tahoma", Font.BOLD, 12));
		searchField.setBounds(721, 239, 252, 23);
		panel_2.add(searchField);
		searchField.setColumns(10);
		
		JLabel lblNewLabel_11 = new JLabel("Search on table  :-");
		lblNewLabel_11.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblNewLabel_11.setBounds(583, 240, 132, 20);
		panel_2.add(lblNewLabel_11);

		btn_cancel = new JButton("");
		btn_cancel.setBounds(813, 560, 111, 31);
		// panel_2.add(btn_cancel);
		btn_cancel.setToolTipText("Click here to Stop the Scanning process.");
		btn_cancel.setRolloverEnabled(false);
		btn_cancel.setRequestFocusEnabled(false);
		btn_cancel.setOpaque(false);
		btn_cancel.setFocusable(false);
		btn_cancel.setFocusPainted(false);
		btn_cancel.setFocusTraversalKeysEnabled(false);
		btn_cancel.setContentAreaFilled(false);
		btn_cancel.setDefaultCapable(false);
		btn_cancel.setBorderPainted(false);
		btn_cancel.setVisible(false);
		btn_cancel.setEnabled(false);
		btn_cancel.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_cancel.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_cancel.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
			}
		});

		btn_cancel.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));

		lblLoadingPleaseWait = new JLabel("Please Wait...");
		lblLoadingPleaseWait.setBounds(36, 548, 134, 20);
		// panel_2.add(lblLoadingPleaseWait);
		lblLoadingPleaseWait.setForeground(UIManager.getColor("Button.highlight"));
		lblLoadingPleaseWait.setFont(new Font("Tahoma", Font.BOLD | Font.ITALIC, 9));
		lblLoadingPleaseWait.setVisible(false);
		btn_cancel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				String warn = "Do you want to stop the process?";
				int ans = JOptionPane.showConfirmDialog(main_multiplefile.this, warn, messageboxtitle,
						JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {
					Stoppreview = true;
					lblLoadingPleaseWait.setVisible(false);
					label_10.setVisible(false);
					btnAttachment.setEnabled(true);
					table_fileinformation.setEnabled(true);
					btnViewer.setEnabled(true);
					btn_next_pane2.setEnabled(true);
					btn_previous_p2.setEnabled(true);
				}

			}
		});

		JButton btn_previous = new JButton("");
		btn_previous.setBounds(10, 556, 8, 31);
		// panel_1.add(btn_previous);
		btn_previous.setToolTipText("Click here to Go Back.");
		btn_previous.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_previous.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_previous.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
			}
		});
		btn_previous.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
		btn_previous.setRolloverEnabled(false);
		btn_previous.setRequestFocusEnabled(false);
		btn_previous.setOpaque(false);
		btn_previous.setFocusable(false);
		btn_previous.setFocusTraversalKeysEnabled(false);
		btn_previous.setFocusPainted(false);
		btn_previous.setDefaultCapable(false);
		btn_previous.setContentAreaFilled(false);
		btn_previous.setBorderPainted(false);
		btn_previous.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				Main_Frame main_m = new Main_Frame(demo, Main_Frame.versiontype);
				main_m.setVisible(true);
				main_m.setLocationRelativeTo(null);
				main_m.setResizable(true);
				main_m.rdbtnSingleFile.setSelected(true);
				dispose();
			}
		});
		btn_previous.setFont(new Font("Tahoma", Font.BOLD, 12));

		JLabel lblNewLabel_2 = new JLabel("New label");
		lblNewLabel_2.setBackground(UIManager.getColor("Button.light"));

		lblNewLabel_2.setOpaque(true);
		lblNewLabel_2.setBounds(-200, 545, 1173, 61);
		panel_1.add(lblNewLabel_2);
		panel_3.setLayout(null);
		panel_3.add(comboBox_fileDestination_type);
		panel_3.add(btn_signout_p3);
		panel_3.add(panel_3_2);
		panel_3_2.setLayout(null);
		panel_3_2.add(tf_Destination_Location);
		panel_3_2.add(btn_Destination);
		panel_3.add(panel_3_);
		panel_3.add(panel_progress);
		panel_progress.setLayout(null);
		panel_progress.add(Progressbar);
		panel_progress.add(btnStop);
		panel_progress.add(lbl_progressreport);
		panel_progress.add(label_11);
		panel_3.add(lblSavesbackupmigrateAs);
		panel_3.add(panel_12);
		panel_12.setLayout(null);
		panel_12.add(btn_previous_p3);
		// panel_12.add(btn_converter_1);
		panel_3.add(radioFileFormat);
		panel_3.add(rdbtnEmailClients);

		btnStop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				String warn = "Do you want to stop the process?";
				int ans = JOptionPane.showConfirmDialog(main_multiplefile.this, warn, messageboxtitle,
						JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {
					// th.interrupt();
					stop = true;
				}

			}
		});

		panel_4 = new JPanel();
		panel_4.setBackground(Color.WHITE);
		Cardlayout.add(panel_4, "panel_4");

		JScrollPane scrollPane_table_panel4 = new JScrollPane();

		table_fileConvertionreport_panel4 = new JTable() {
			/**
			 *
			 */
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {

				return false;
			}
		};

		table_fileConvertionreport_panel4.setModel(new DefaultTableModel(new Object[][] {},
				new String[] { "<html><b>" + "From", "<html><b>" + "To", "<html><b>" + "File Name",
						"<html><b>" + "Status", "<html><b>" + "Duration", "<html><b>" + "Message count",
						"<html><b>" + "Path" }));
		scrollPane_table_panel4.setViewportView(table_fileConvertionreport_panel4);
		table_fileConvertionreport_panel4.getColumnModel().getColumn(2).setPreferredWidth(126);
		btnDowloadReport = new JButton("");
		btnDowloadReport_1 = new JButton("");
		btnDowloadReport_1.setToolTipText("Click here to Download the Report.");
		btnDowloadReport_1.setRolloverEnabled(false);
		btnDowloadReport_1.setRequestFocusEnabled(false);
		btnDowloadReport_1.setOpaque(false);
		btnDowloadReport_1.setFocusable(false);
		btnDowloadReport_1.setFocusTraversalKeysEnabled(false);
		btnDowloadReport_1.setFocusPainted(false);
		btnDowloadReport_1.setDefaultCapable(false);
		btnDowloadReport_1.setContentAreaFilled(false);
		btnDowloadReport_1.setBorderPainted(false);
		btnDowloadReport_1.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnDowloadReport_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/download-report-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnDowloadReport_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/download-report-btn.png")));
			}
		});

		btnDowloadReport_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/download-report-btn.png")));

		btnDowloadReport_1.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				cal = Calendar.getInstance();
				calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
				reportpath = logpathm;
				new File(reportpath + File.separator + messageboxtitle + " report").mkdirs();

				File file = new File(reportpath + File.separator + messageboxtitle + " report" + File.separator
						+ calendertime + "report.csv");

				try {
					FileWriter outputfile = new FileWriter(file);

					CSVWriter writer = new CSVWriter(outputfile);

					String[] header = { "From", "To", "File Name", "Status", "Duration", "Message Count", "Path" };

					writer.writeNext(header);

					for (int i = 0; i < table_fileConvertionreport_panel4.getRowCount(); i++) {
						String g1 = "";
						try {
							g1 = table_fileConvertionreport_panel4.getValueAt(i, 0).toString();
						} catch (Exception e) {

						}
						String g2 = "";
						try {
							g2 = table_fileConvertionreport_panel4.getValueAt(i, 1).toString();
						} catch (Exception e) {

						}
						String g3 = "";
						try {
							g3 = table_fileConvertionreport_panel4.getValueAt(i, 2).toString();
						} catch (Exception e) {

						}
						String g4 = "";
						try {
							g4 = table_fileConvertionreport_panel4.getValueAt(i, 3).toString();
						} catch (Exception e) {

						}
						String g5 = "";
						try {
							g5 = table_fileConvertionreport_panel4.getValueAt(i, 4).toString();
						} catch (Exception e) {

						}
						String g6 = "";
						try {
							g6 = table_fileConvertionreport_panel4.getValueAt(i, 5).toString();
						} catch (Exception e) {

						}

						String g7 = "";
						try {
							g7 = table_fileConvertionreport_panel4.getValueAt(i, 6).toString();
						} catch (Exception e) {

						}

						String[] data1 = { g1, g2, g3, g4, g5, g6, g7 };

						writer.writeNext(data1);
					}

					writer.close();
					file.setReadOnly();
					Desktop desktop = Desktop.getDesktop();
					desktop.open(file);

				} catch (Exception e) {

					e.printStackTrace();
				}
			}
		});
		btnDowloadReport_1.setFont(new Font("Tahoma", Font.BOLD, 15));

		panel_13 = new JPanel();
		panel_13.setBackground(new Color(0, 0, 0));

		JButton btnConvertAgain = new JButton("");
		btnConvertAgain.setToolTipText("Click here to Convert Again.");
		btnConvertAgain.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnConvertAgain.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-again-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnConvertAgain.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-again-btn.png")));
			}
		});
		btnConvertAgain.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-again-btn.png")));
		btnConvertAgain.setBorderPainted(false);
		btnConvertAgain.setContentAreaFilled(false);
		btnConvertAgain.setDefaultCapable(false);
		btnConvertAgain.setFocusTraversalKeysEnabled(false);
		btnConvertAgain.setFocusable(false);
		btnConvertAgain.setOpaque(false);
		btnConvertAgain.setRolloverEnabled(false);
		btnConvertAgain.setRequestFocusEnabled(false);
		btnConvertAgain.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")
						|| filetype.equalsIgnoreCase("OFFICE 365"))) {

					System.out.println("here we riched");
					textField_username_p3.setEnabled(true);
					passwordField_p3.setEnabled(true);
					tf_portNo_p3.setEnabled(true);
					chckbxShowPassword_p3.setEnabled(true);
					lblPortNo.setEnabled(true);
					lblNewLabel_5.setEnabled(true);
					lblNewLabel_1.setEnabled(true);
					lblNewLabel.setEnabled(true);
					lblemailAddress.setEnabled(true);
					lblemailAddress.setEnabled(false);

				}
				textField_customfolder.setEnabled(true);
				archive.setEnabled(true);
				publicfolder.setEnabled(true);
				mailbox.setEnabled(true);
				modern_Authentication.setEnabled(true);
				basic_Authentication.setEnabled(true);

//                           basic_Authentication.setSelected(true);
				filetype = "";
				path = "";
				checkconvertagain = true;
				stop = false;
				btnStop.setVisible(false);
				chckbxSavePdfAttachment.setEnabled(true);
				btn_Destination.setEnabled(true);
				btn_previous_p3.setEnabled(true);
				chckbxSaveInSame.setEnabled(true);
				comboBox_fileDestination_type.setEnabled(true);
				btn_Destination.setEnabled(true);
				btn_previous_p3.setEnabled(true);
				lbl_progressreport.setText("");
				comboBox.setEnabled(true);
				chckbx_splitpst.setEnabled(true);
				textField_customfolder.setEditable(true);
				chckbxMigrateOrBackup.setEnabled(true);
				dateChooser_calender_start.setEnabled(true);
				chckbx_convert_pdf_to_pdf.setEnabled(true);
				chckbxRemoveDuplicacy.setEnabled(true);
				datefilter.setEnabled(true);
				dateChooser_calendar_end.setEnabled(true);
				chckbxCustomFolderName.setEnabled(true);
				btn_Destination.setEnabled(true);
				btn_previous_p3.setEnabled(true);
				checkmboxpstost = true;
				chckbxRestoreToDefault.setEnabled(true);
				chckbxSaveMboxIn.setEnabled(true);
				btn_signout_p3.setVisible(false);
				label_11.setVisible(false);
				panel_5.setEnabled(true);
				chckbx_Mail_Filter.setSelected(false);
				chckbx_Mail_Filter.setEnabled(true);
				chckbx_calender_box.setEnabled(true);
				task_box.setEnabled(true);
				chckbxMaintainFolderStructure.setEnabled(true);
				textField_domain_name_p3.setEnabled(true);
				textField_username_p3.setEnabled(true);
				passwordField_p3.setEnabled(true);
				tf_portNo_p3.setEnabled(true);
				lblNewLabel.setEnabled(true);
				lblemailAddress.setEnabled(true);
				lblNewLabel_1.setEnabled(true);
				lblNewLabel_5.setEnabled(true);
				lblPortNo.setEnabled(true);
				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_2");

			}
		});
		btnConvertAgain.setFont(new Font("Tahoma", Font.BOLD, 14));
		GroupLayout gl_panel_4 = new GroupLayout(panel_4);
		gl_panel_4
				.setHorizontalGroup(
						gl_panel_4.createParallelGroup(Alignment.LEADING)
								.addGroup(gl_panel_4.createSequentialGroup().addGap(7)
										.addComponent(scrollPane_table_panel4, GroupLayout.DEFAULT_SIZE, 1063,
												Short.MAX_VALUE)
										.addGap(5))
								.addGroup(gl_panel_4.createSequentialGroup().addGap(490)
										.addComponent(btnDowloadReport_1, GroupLayout.PREFERRED_SIZE, 141,
												Short.MAX_VALUE)
										.addGap(444))
								.addComponent(panel_13, GroupLayout.DEFAULT_SIZE, 1075, Short.MAX_VALUE));
		gl_panel_4.setVerticalGroup(gl_panel_4.createParallelGroup(Alignment.LEADING).addGroup(gl_panel_4
				.createSequentialGroup().addGap(13)
				.addComponent(scrollPane_table_panel4, GroupLayout.DEFAULT_SIZE, 474, Short.MAX_VALUE).addGap(23)
				.addComponent(btnDowloadReport_1, GroupLayout.PREFERRED_SIZE, 31, GroupLayout.PREFERRED_SIZE).addGap(1)
				.addComponent(panel_13, GroupLayout.PREFERRED_SIZE, 70, GroupLayout.PREFERRED_SIZE).addGap(4)));
		GroupLayout gl_panel_13 = new GroupLayout(panel_13);
		gl_panel_13.setHorizontalGroup(gl_panel_13.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_panel_13.createSequentialGroup().addGap(489)
						.addComponent(btnConvertAgain, GroupLayout.PREFERRED_SIZE, 141, Short.MAX_VALUE).addGap(445)));
		gl_panel_13.setVerticalGroup(gl_panel_13.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_panel_13.createSequentialGroup().addGap(20).addComponent(btnConvertAgain,
						GroupLayout.PREFERRED_SIZE, 31, GroupLayout.PREFERRED_SIZE)));
		panel_13.setLayout(gl_panel_13);
		panel_4.setLayout(gl_panel_4);

		panel_10 = new JPanel();
		panel_10.setForeground(Color.WHITE);
		panel_10.setBounds(0, 0, 1174, 71);
		panel_10.setBackground(Color.WHITE);

		Label label_4 = new Label("");
		label_4.setGradientColor2(new Color(255, 204, 51));
		label_4.setGradientColor1(new Color(153, 204, 153));
		label_4.setBounds(0, 0, 1174, 78);
		contentPane.setLayout(null);
		contentPane.add(panel_10);
		panel_10.setLayout(null);

		lblNewLabel_4 = new JLabel("New label");
		lblNewLabel_4.setIcon(new ImageIcon(Main_Frame.class.getResource("/topbard.png")));
		lblNewLabel_4.setBounds(10, 5, 72, 60);
		panel_10.add(lblNewLabel_4);

		JLabel lblNewLabel_3_1 = new JLabel(
				"Perfect tool to convert Pst file into various file format like PDF,CSV,Gmail,Office-365,+12(other format).");
		lblNewLabel_3_1.setForeground(Color.BLACK);
		lblNewLabel_3_1.setFont(new Font("Tahoma", Font.BOLD, 15));
		lblNewLabel_3_1.setBounds(208, 34, 862, 20);
		panel_10.add(lblNewLabel_3_1);

		JLabel lblNewLabel_3 = new JLabel("<html><u>DevopixTech Software Solution");
		lblNewLabel_3.setForeground(Color.BLACK);
		lblNewLabel_3.setFont(new Font("Tahoma", Font.BOLD, 20));
		lblNewLabel_3.setBounds(410, 0, 375, 42);
		panel_10.add(lblNewLabel_3);
		panel_10.add(label_4);
		contentPane.add(Cardlayout);
		
		JPanel panel_9 = new JPanel();
		panel_9.setBackground(Color.WHITE);
		Cardlayout.add(panel_9, "panel_9");
		panel_9.setLayout(null);
		
		JLabel lblNewLabel_12 = new JLabel("New label");
		lblNewLabel_12.setBounds(46, 66, 73, 59);
		panel_9.add(lblNewLabel_12);
		
		JLabel lblNewLabel_12_1 = new JLabel("New label");
		lblNewLabel_12_1.setBounds(169, 66, 73, 59);
		panel_9.add(lblNewLabel_12_1);
		
		JLabel lblNewLabel_12_2 = new JLabel("New label");
		lblNewLabel_12_2.setBounds(307, 66, 73, 59);
		panel_9.add(lblNewLabel_12_2);
		
		JLabel lblNewLabel_12_3 = new JLabel("New label");
		lblNewLabel_12_3.setBounds(436, 66, 73, 59);
		panel_9.add(lblNewLabel_12_3);
		
		JLabel lblNewLabel_12_4 = new JLabel("New label");
		lblNewLabel_12_4.setBounds(566, 66, 73, 59);
		panel_9.add(lblNewLabel_12_4);
		
		JLabel lblNewLabel_12_5 = new JLabel("New label");
		lblNewLabel_12_5.setBounds(683, 66, 73, 59);
		panel_9.add(lblNewLabel_12_5);
		
		JLabel lblNewLabel_12_6 = new JLabel("New label");
		lblNewLabel_12_6.setBounds(818, 66, 73, 59);
		panel_9.add(lblNewLabel_12_6);
		
		JLabel lblNewLabel_12_6_1 = new JLabel("New label");
		lblNewLabel_12_6_1.setBounds(818, 174, 73, 59);
		panel_9.add(lblNewLabel_12_6_1);
		
		JLabel lblNewLabel_12_5_1 = new JLabel("New label");
		lblNewLabel_12_5_1.setBounds(683, 174, 73, 59);
		panel_9.add(lblNewLabel_12_5_1);
		
		JLabel lblNewLabel_12_4_1 = new JLabel("New label");
		lblNewLabel_12_4_1.setBounds(566, 174, 73, 59);
		panel_9.add(lblNewLabel_12_4_1);
		
		JLabel lblNewLabel_12_3_1 = new JLabel("New label");
		lblNewLabel_12_3_1.setBounds(436, 174, 73, 59);
		panel_9.add(lblNewLabel_12_3_1);
		
		JLabel lblNewLabel_12_2_1 = new JLabel("New label");
		lblNewLabel_12_2_1.setBounds(307, 174, 73, 59);
		panel_9.add(lblNewLabel_12_2_1);
		
		JLabel lblNewLabel_12_1_1 = new JLabel("New label");
		lblNewLabel_12_1_1.setBounds(169, 174, 73, 59);
		panel_9.add(lblNewLabel_12_1_1);
		
		JLabel lblNewLabel_12_7 = new JLabel("New label");
		lblNewLabel_12_7.setBounds(46, 174, 73, 59);
		panel_9.add(lblNewLabel_12_7);
		
		JLabel lblNewLabel_12_6_2 = new JLabel("New label");
		lblNewLabel_12_6_2.setBounds(818, 310, 73, 59);
		panel_9.add(lblNewLabel_12_6_2);
		
		JLabel lblNewLabel_12_5_2 = new JLabel("New label");
		lblNewLabel_12_5_2.setBounds(683, 310, 73, 59);
		panel_9.add(lblNewLabel_12_5_2);
		
		JLabel lblNewLabel_12_4_2 = new JLabel("New label");
		lblNewLabel_12_4_2.setBounds(566, 310, 73, 59);
		panel_9.add(lblNewLabel_12_4_2);
		
		JLabel lblNewLabel_12_3_2 = new JLabel("New label");
		lblNewLabel_12_3_2.setBounds(436, 310, 73, 59);
		panel_9.add(lblNewLabel_12_3_2);
		
		JLabel lblNewLabel_12_2_2 = new JLabel("New label");
		lblNewLabel_12_2_2.setBounds(307, 310, 73, 59);
		panel_9.add(lblNewLabel_12_2_2);
		
		JLabel lblNewLabel_12_1_2 = new JLabel("New label");
		lblNewLabel_12_1_2.setBounds(169, 310, 73, 59);
		panel_9.add(lblNewLabel_12_1_2);
		
		JLabel lblNewLabel_12_8 = new JLabel("New label");
		lblNewLabel_12_8.setBounds(46, 310, 73, 59);
		panel_9.add(lblNewLabel_12_8);
		
		JLabel lblNewLabel_12_6_3 = new JLabel("New label");
		lblNewLabel_12_6_3.setBounds(818, 469, 73, 59);
		panel_9.add(lblNewLabel_12_6_3);
		
		JLabel lblNewLabel_12_5_3 = new JLabel("New label");
		lblNewLabel_12_5_3.setBounds(683, 469, 73, 59);
		panel_9.add(lblNewLabel_12_5_3);
		
		JLabel lblNewLabel_12_4_3 = new JLabel("New label");
		lblNewLabel_12_4_3.setBounds(566, 469, 73, 59);
		panel_9.add(lblNewLabel_12_4_3);
		
		JLabel lblNewLabel_12_3_3 = new JLabel("New label");
		lblNewLabel_12_3_3.setBounds(436, 469, 73, 59);
		panel_9.add(lblNewLabel_12_3_3);
		
		JLabel lblNewLabel_12_2_3 = new JLabel("New label");
		lblNewLabel_12_2_3.setBounds(307, 469, 73, 59);
		panel_9.add(lblNewLabel_12_2_3);
		
		JLabel lblNewLabel_12_1_3 = new JLabel("New label");
		lblNewLabel_12_1_3.setBounds(169, 469, 73, 59);
		panel_9.add(lblNewLabel_12_1_3);
		
		JLabel lblNewLabel_12_9 = new JLabel("New label");
		lblNewLabel_12_9.setBounds(46, 469, 73, 59);
		panel_9.add(lblNewLabel_12_9);
		
		JRadioButton rdbtnNewRadioButton_yahoo = new JRadioButton("YAHOO");
		rdbtnNewRadioButton_yahoo.setBounds(27, 553, 109, 23);
		panel_9.add(rdbtnNewRadioButton_yahoo);
		
		JRadioButton rdbtnNewRadioButton_thunderword = new JRadioButton("THUNDERWORD");
		rdbtnNewRadioButton_thunderword.setBounds(150, 553, 109, 23);
		panel_9.add(rdbtnNewRadioButton_thunderword);
		
		JRadioButton rdbtnNewRadioButton_aol = new JRadioButton("AOL");
		rdbtnNewRadioButton_aol.setBounds(292, 553, 109, 23);
		panel_9.add(rdbtnNewRadioButton_aol);
		
		JRadioButton rdbtnNewRadioButton_hotmail = new JRadioButton("HOTMAIL");
		rdbtnNewRadioButton_hotmail.setBounds(422, 553, 109, 23);
		panel_9.add(rdbtnNewRadioButton_hotmail);
		
		JRadioButton rdbtnNewRadioButton_imap = new JRadioButton("IMAP");
		rdbtnNewRadioButton_imap.setBounds(555, 553, 109, 23);
		panel_9.add(rdbtnNewRadioButton_imap);
		
		JRadioButton rdbtnNewRadioButton_zoho = new JRadioButton("ZOHO");
		rdbtnNewRadioButton_zoho.setBounds(683, 553, 109, 23);
		panel_9.add(rdbtnNewRadioButton_zoho);
		
		JRadioButton rdbtnNewRadioButton_icloud = new JRadioButton("ICLOUD");
		rdbtnNewRadioButton_icloud.setBounds(815, 553, 109, 23);
		panel_9.add(rdbtnNewRadioButton_icloud);
		
		JRadioButton rdbtnNewRadioButton_gsuit = new JRadioButton("G_SUITE");
		rdbtnNewRadioButton_gsuit.setBounds(782, 402, 109, 23);
		panel_9.add(rdbtnNewRadioButton_gsuit);
		
		JRadioButton rdbtnNewRadioButton_gmail = new JRadioButton("GMAIL");
		rdbtnNewRadioButton_gmail.setBounds(650, 402, 109, 23);
		panel_9.add(rdbtnNewRadioButton_gmail);
		
		JRadioButton rdbtnNewRadioButton_office = new JRadioButton("OFFICE_365");
		rdbtnNewRadioButton_office.setBounds(522, 402, 109, 23);
		panel_9.add(rdbtnNewRadioButton_office);
		
		JRadioButton rdbtnNewRadioButton_ics = new JRadioButton("ICS");
		rdbtnNewRadioButton_ics.setBounds(389, 402, 109, 23);
		panel_9.add(rdbtnNewRadioButton_ics);
		
		JRadioButton rdbtnNewRadioButton_vcf = new JRadioButton("VCF");
		rdbtnNewRadioButton_vcf.setBounds(259, 402, 109, 23);
		panel_9.add(rdbtnNewRadioButton_vcf);
		
		JRadioButton rdbtnNewRadioButton_docm = new JRadioButton("DOCM");
		rdbtnNewRadioButton_docm.setBounds(117, 402, 109, 23);
		panel_9.add(rdbtnNewRadioButton_docm);
		
		JRadioButton rdbtnNewRadioButton_docx = new JRadioButton("DOCX");
		rdbtnNewRadioButton_docx.setBounds(-6, 402, 109, 23);
		panel_9.add(rdbtnNewRadioButton_docx);
		
		JRadioButton rdbtnNewRadioButton_doc = new JRadioButton("DOC");
		rdbtnNewRadioButton_doc.setBounds(794, 265, 109, 23);
		panel_9.add(rdbtnNewRadioButton_doc);
		
		JRadioButton rdbtnNewRadioButton_png = new JRadioButton("PNG");
		rdbtnNewRadioButton_png.setBounds(662, 265, 109, 23);
		panel_9.add(rdbtnNewRadioButton_png);
		
		JRadioButton rdbtnNewRadioButton_mhtml = new JRadioButton("MHTML");
		rdbtnNewRadioButton_mhtml.setBounds(534, 265, 109, 23);
		panel_9.add(rdbtnNewRadioButton_mhtml);
		
		JRadioButton rdbtnNewRadioButton_html = new JRadioButton("HTML");
		rdbtnNewRadioButton_html.setBounds(401, 265, 109, 23);
		panel_9.add(rdbtnNewRadioButton_html);
		
		JRadioButton rdbtnNewRadioButton_tiff = new JRadioButton("TIFF");
		rdbtnNewRadioButton_tiff.setBounds(271, 265, 109, 23);
		panel_9.add(rdbtnNewRadioButton_tiff);
		
		JRadioButton rdbtnNewRadioButton_jpg = new JRadioButton("JPG");
		rdbtnNewRadioButton_jpg.setBounds(129, 265, 109, 23);
		panel_9.add(rdbtnNewRadioButton_jpg);
		
		JRadioButton rdbtnNewRadioButton_gif = new JRadioButton("GIF");
		rdbtnNewRadioButton_gif.setBounds(6, 265, 109, 23);
		panel_9.add(rdbtnNewRadioButton_gif);
		
		JRadioButton rdbtnNewRadioButton_CSV = new JRadioButton("CSV");
		rdbtnNewRadioButton_CSV.setBounds(815, 132, 109, 23);
		panel_9.add(rdbtnNewRadioButton_CSV);
		
		JRadioButton rdbtnNewRadioButton_pdf = new JRadioButton("PDF");
		rdbtnNewRadioButton_pdf.setBounds(683, 132, 109, 23);
		panel_9.add(rdbtnNewRadioButton_pdf);
		
		JRadioButton rdbtnNewRadioButton_MSG = new JRadioButton("MSG");
		rdbtnNewRadioButton_MSG.setBounds(555, 132, 109, 23);
		panel_9.add(rdbtnNewRadioButton_MSG);
		
		JRadioButton rdbtnNewRadioButton_EMLX = new JRadioButton("EMLX");
		rdbtnNewRadioButton_EMLX.setBounds(422, 132, 109, 23);
		panel_9.add(rdbtnNewRadioButton_EMLX);
		
		JRadioButton rdbtnNewRadioButton_eml = new JRadioButton("EML");
		rdbtnNewRadioButton_eml.setBounds(292, 132, 109, 23);
		panel_9.add(rdbtnNewRadioButton_eml);
		
		JRadioButton rdbtnNewRadioButton_mbox = new JRadioButton("MBOX");
		rdbtnNewRadioButton_mbox.setBounds(150, 132, 109, 23);
		panel_9.add(rdbtnNewRadioButton_mbox);
		
		JRadioButton rdbtnNewRadioButton_pst = new JRadioButton("Pst");
		rdbtnNewRadioButton_pst.setBounds(27, 132, 109, 23);
		panel_9.add(rdbtnNewRadioButton_pst);

		panel_7 = new JPanel();
		panel_7.setBounds(0, 71, 199, 608);
		contentPane.add(panel_7);
		panel_7.setBackground(UIManager.getColor("Button.light"));
		panel_7.setBorder(null);
		panel_7.setLayout(null);

		btn_Next = new GradientButton("Tree");
		btn_Next.setEnabled(false);
		btn_Next.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btn_Next.isEnabled()) {
					btn_Next.setGradientColor1(new Color(70, 130, 180));
					btn_Next.setGradientColor2(new Color(70, 130, 180));
					btn_Next.setForeground(new Color(255, 255, 255));
				}
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_Next.setGradientColor1(new Color(255, 255, 255));
				btn_Next.setGradientColor2(new Color(255, 255, 255));
				btn_Next.setForeground(new Color(80, 80, 80));
			}
		});

		//
		updateBtn = new JButton("");
		updateBtn.setBounds(101, 521, 80, 79);
		panel_7.add(updateBtn);
		updateBtn.setFont(new Font("Tahoma", Font.BOLD, 11));
		// updateBtn.addMouseListener(new MouseAdapter() {
		// public void mouseEntered(MouseEvent arg0) {
		// updateBtn.setIcon(new
		// ImageIcon(Main_Frame.class.getResource("/update-hvr-btn.png")));
		// }
		//
		// public void mouseExited(MouseEvent e) {
		// updateBtn.setIcon(new
		// ImageIcon(Main_Frame.class.getResource("/update-btn.png")));
		// }
		// });
		updateBtn.setIcon(new ImageIcon(Main_Frame.class.getResource("/update-btn.png")));
		updateBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					Desktop.getDesktop().browse(new URL(All_Data.updateSoftware).toURI());
				} catch (IOException | URISyntaxException e1) {

					e1.printStackTrace();
					StringWriter sw = new StringWriter();
					PrintWriter pw = new PrintWriter(sw);
					e1.printStackTrace(pw);

				}
			}
		});

		updateBtn.setFocusTraversalKeysEnabled(false);
		updateBtn.setFocusable(false);
		updateBtn.setOpaque(false);
		updateBtn.setRolloverEnabled(false);
		updateBtn.setRequestFocusEnabled(false);
		updateBtn.setFocusPainted(false);
		updateBtn.setDefaultCapable(false);
		updateBtn.setContentAreaFilled(false);
		updateBtn.setVisible(false);
		updateBtn.setBorderPainted(false);
		updateBtn.setVisible(false);
		updateBtn.setToolTipText("Click here to download the latest version of the software.");

		btn_Next.setShadowColor(new Color(0, 0, 204));
		btn_Next.setGradientColor1(Color.WHITE);
		btn_Next.setBounds(5, 74, 190, 50);

		panel_7.add(btn_Next);

		btn_Next.setToolTipText("Click here to Go Forward.");

		btn_Next.setRolloverEnabled(false);
		btn_Next.setRequestFocusEnabled(false);
		btn_Next.setOpaque(false);
		btn_Next.setFocusable(false);
		btn_Next.setFocusTraversalKeysEnabled(false);
		btn_Next.setFocusPainted(false);
		btn_Next.setDefaultCapable(false);
		btn_Next.setContentAreaFilled(false);
		btn_Next.setBorderPainted(false);
		btn_Next.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				second = true;

				Buttonclick("btn_Next", second);

				if (table.getRowCount() == 0) {

					JOptionPane.showMessageDialog(mf, "Choose file First then Procedure", messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));
					return;

				}

				Thread thread = new Thread(new Runnable() {

					@Override
					public void run() {
						lists.clear();
						listst.clear();

						btn_next_pane2.setEnabled(true);
						btn_converter_1.setEnabled(true);
						btnNewButton_2.setEnabled(true);
						btn_Next_1.setEnabled(true);

						Object[] s = new Object[table.getRowCount()];
						filesfin = new String[table.getRowCount()];

						for (int i = 0; i < table.getRowCount(); i++) {
							s[i] = table.getValueAt(i, 2);
						}

						for (int i = 0; i < s.length; i++) {

							filesfin[i] = (String) s[i];

						}
						parent = files[0].getParent();
						mf.logger = mf.logFile();
						mf.logger.info("Start Time : " + calendertime + System.lineSeparator() + "File Type : "
								+ fileoptionm + "                         " + "File filetype" + "    " + fileoptionm
								+ System.lineSeparator()
								+ "======================================================================");

//				SwingWorker sw1 = new SwingWorker() {
//					@Override
//					protected Object doInBackground() {
//						obTh = new LoadingThreadclass(mf);
//						obTh.start();
						loadingDialog = new LoadingDialog(mf, true);
						loadingDialog.setLocationRelativeTo(mf);

						Thread innerthread = new Thread(new Runnable() {
							public void run() {
								loadingDialog.setVisible(true);
							}
						});
						innerthread.start();

						InetAddress addr;
						String hostName = "";
						try {
							addr = InetAddress.getLocalHost();
							hostName = addr.getHostName();

						} catch (UnknownHostException e1) {

							e1.printStackTrace();
						}

						model = (DefaultTreeModel) tree.getModel();

						root = new DefaultMutableTreeNode("<html><b>" + hostName);

						model.setRoot(root);

						mainnode = new DefaultMutableTreeNode("<html><b>" + fileoptionm);

						root.add(mainnode);

						CardLayout card = (CardLayout) Cardlayout.getLayout();
						card.show(Cardlayout, "panel_2");
						foldercountcheck = 0;
						for (int i = 0; i < filesfin.length; i++) {

							String filetype = table.getValueAt(i, 3).toString();
							filetype = filetype.replace("<html><b>", "");
							if (filetype.equalsIgnoreCase("file")) {

								filepath = filesfin[i].replace("<html><b>", "");
								path2 = filepath;

								if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
										|| fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

									try {
										String extension = getFileExtension(new File(path2));
										if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
											if (extension.equalsIgnoreCase("pst")) {

												try {
													readAnOST_PstFile();
												} catch (Exception e) {
													i--;
													mainnode.removeAllChildren();
													continue;
												}
											}

										} else if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

											if (extension.equalsIgnoreCase("ost")) {

												try {
													readAnOST_PstFile();
												} catch (Exception e) {
													i--;
													mainnode.removeAllChildren();
													continue;
												}
											}
										}

									} catch (Exception e1) {

										e1.printStackTrace();
									}

								} else if (fileoptionm.equalsIgnoreCase("MBOX")) {
									try {
										readMboxFile();
									} catch (Exception e) {
										i--;
										mainnode.removeAllChildren();
										continue;
									}

								} else if (fileoptionm.equalsIgnoreCase("DBX")) {
									try {
										readMboxFile();
									} catch (Exception e) {
										i--;
										mainnode.removeAllChildren();
										continue;
									}
								} else if (fileoptionm.equalsIgnoreCase("OLM File (.olm)")) {
									try {
										readolmFile();
									} catch (Exception e) {
										i--;
										mainnode.removeAllChildren();
										continue;
									}
								} else {
									String extension = getFileExtension(new File(path2));
									if (fileoptionm.equalsIgnoreCase("EML File (.eml)")) {
										if (extension.equals("eml")) {
											try {
												readmailFile();
											} catch (Exception e) {
												i--;
												mainnode.removeAllChildren();
												continue;
											}

										}
									} else if (fileoptionm.equalsIgnoreCase("EMLX File (.emlx)")) {
										if (extension.equalsIgnoreCase("emlx")) {

											try {
												readmailFile();
											} catch (Exception e) {
												i--;
												mainnode.removeAllChildren();
												continue;
											}

										}
									} else if (fileoptionm.equalsIgnoreCase("OFT File (.oft)")) {
										if (extension.equalsIgnoreCase("oft")) {

											try {
												readmailFile();
											} catch (Exception e) {
												i--;
												mainnode.removeAllChildren();
												continue;
											}

										}
									} else if (fileoptionm.equalsIgnoreCase("Message File (.msg)")) {
										if (extension.equalsIgnoreCase("msg")) {

											try {
												readmailFile();
											} catch (Exception e) {
												i--;
												mainnode.removeAllChildren();
												continue;
											}
										}
									} else if (fileoptionm.equalsIgnoreCase("Maildir")) {

										try {
											readmailFile();
										} catch (Exception e) {
											i--;
											mainnode.removeAllChildren();
											continue;
										}

									}

								}
							} else {

								if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
										|| fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

									try {

										read_PSTOST_folder(new File(filesfin[i].replace("<html><b>", "")));

									} catch (Exception e1) {

										i--;
										mainnode.removeAllChildren();
										continue;
									}

								} else if (fileoptionm.equalsIgnoreCase("MBOX")) {
									try {

										read_mbox_folder(new File(filesfin[i].replace("<html><b>", "")));

									} catch (Exception e1) {

										i--;
										mainnode.removeAllChildren();
										continue;
									}

								} else if (fileoptionm.equalsIgnoreCase("OLM File (.olm)")) {

									try {

										read_olm_folder(new File(filesfin[i].replace("<html><b>", "")));

									} catch (Exception e1) {

										i--;
										mainnode.removeAllChildren();
										continue;
									}

								} else {

									try {
										File fol = new File(filesfin[i].replace("<html><b>", ""));
										String s1 = filepath(fol);
										CustomTreeNode main = new CustomTreeNode("<html><b>" + s1);
										mainnode.add(main);

										testemldd(fol, main);
									} catch (Exception e1) {
										i--;
										mainnode.removeAllChildren();
										continue;
									}

								}

							}
						}
//						return null;
//					}

//					@Override
//					protected void done() {
//
						Icon open = new ImageIcon(backup.class.getResource("/Open-folder-accept-icon.png"));
						Icon close = new ImageIcon(backup.class.getResource("/closed-folder-add-icon.png"));
						Icon Ram = new ImageIcon(backup.class.getResource("/leaf-icon.png"));
						DefaultCheckboxTreeCellRenderer render = (DefaultCheckboxTreeCellRenderer) tree
								.getCellRenderer();
						render.setClosedIcon(close);
						render.setOpenIcon(open);
						render.setLeafIcon(Ram);

						tree.expandRow(0);
						tree.expandAll();
						loadingDialog.dispose();

						if (!cropted) {
							JOptionPane.showMessageDialog(mf,
									"Your File is Imported Successfully.Please Expand the tree Hierarchy",
									All_Data.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/information.png")));

						}

						tree.getCheckingModel().setCheckingPath(tree.getPathForRow(0));

						// obTh.close();
						if (fileoptionm.equalsIgnoreCase("EML File (.eml)")
								|| fileoption.equalsIgnoreCase("EMLX File (.emlx)")
								|| fileoption.equalsIgnoreCase("OFT File (.oft)")
								|| fileoptionm.equalsIgnoreCase("Message File (.msg)")
								|| fileoptionm.equalsIgnoreCase("Maildir")) {

							comboBox.addItem("Original File Name");

						}

						if (foldercountcheck == 0) {
							JOptionPane.showMessageDialog(mf, "No " + fileoptionm + " file Can be found",
									messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/information.png")));
							CardLayout card1 = (CardLayout) Cardlayout.getLayout();
							card1.show(Cardlayout, "panel_1");
						}
//					}
//				};
//
//				sw1.execute();

					}

				});

				thread.start();

			}

		});
		btn_Next.setFont(new Font("Verdana", Font.BOLD, 12));

		btn_previous_p2 = new GradientButton("Home");
		btn_previous_p2.setEnabled(false);
		btn_previous_p2.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btn_previous_p2.isEnabled()) {
					btn_previous_p2.setGradientColor1(new Color(70, 130, 180));
					btn_previous_p2.setGradientColor2(new Color(70, 130, 180));
					btn_previous_p2.setForeground(new Color(255, 255, 255));
				}
			}

			@Override
			public void mouseExited(MouseEvent e) {

				btn_previous_p2.setGradientColor1(new Color(255, 255, 255));
				btn_previous_p2.setGradientColor2(new Color(255, 255, 255));
				btn_previous_p2.setForeground(new Color(80, 80, 80));

			}

		});
		btn_previous_p2.setShadowColor(new Color(0, 0, 255));
		btn_previous_p2.setForeground(UIManager.getColor("CheckBox.shadow"));
		// btn_previous_p2.setBounds(798, 21, 111, 31);
		btn_previous_p2.setToolTipText("Click here to Go Back.\r\n");
		btn_previous_p2.setRolloverEnabled(false);
		btn_previous_p2.setRequestFocusEnabled(false);
		btn_previous_p2.setOpaque(false);
		btn_previous_p2.setFocusable(false);
		btn_previous_p2.setFocusTraversalKeysEnabled(false);
		btn_previous_p2.setFocusPainted(false);
		btn_previous_p2.setDefaultCapable(false);
		btn_previous_p2.setContentAreaFilled(false);
		btn_previous_p2.setBorderPainted(false);
		btn_previous_p2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if (table.getRowCount() == 0) {
					JOptionPane.showMessageDialog(null, "Please add file or Folder in the table and then Procude");
				}
				fir = true;

				Buttonclick("btn_previous", fir);

				lblTotalMessageCount.setText("<html><b>" + "  Total Message Count : ");
				editorPane.setText("");
				Stoppreview = true;
				CardLayout card1 = (CardLayout) innercardlayout.getLayout();
				card1.show(innercardlayout, "panel_Contact");
				label_contactfullname.setText("");
				label_contactemail.setText("");
				label_contactcompany.setText("");
				label_contactphonenumber.setText("");
				textArea_contact.setText("");
				model = (DefaultTreeModel) tree.getModel();
				DefaultMutableTreeNode root = (DefaultMutableTreeNode) model.getRoot();
				root.removeAllChildren();
				model.reload();
				TreePath[] ac = new TreePath[0];
				tree.setCheckingPaths(ac);
				DefaultTableModel model = (DefaultTableModel) table_fileinformation.getModel();

				while (model.getRowCount() > 0) {

					for (int i = 0; i < model.getRowCount(); ++i) {

						model.removeRow(i);
					}
				}

				DefaultTableModel model1 = (DefaultTableModel) table_1.getModel();

				while (model1.getRowCount() > 0) {

					for (int i = 0; i < model1.getRowCount(); ++i) {

						model1.removeRow(i);
					}
				}

				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_1");
			}
		});
		btn_previous_p2.setFont(new Font("Verdana", Font.BOLD, 13));
		// panel_9.add(btn_previous_p2);

		btn_previous_p2.setBounds(5, 20, 190, 50);
		panel_7.add(btn_previous_p2);

		btn_next_pane2.setRolloverEnabled(false);
		btn_next_pane2.setRequestFocusEnabled(false);
		btn_next_pane2.setOpaque(false);
		btn_next_pane2.setFocusable(false);
		btn_next_pane2.setFocusTraversalKeysEnabled(false);
		btn_next_pane2.setFocusPainted(false);
		btn_next_pane2.setDefaultCapable(false);
		btn_next_pane2.setContentAreaFilled(false);
		btn_next_pane2.setBorderPainted(false);

		btn_next_pane2.setShadowColor(new Color(0, 0, 204));
		btn_next_pane2.setFont(new Font("Verdana", Font.BOLD, 12));
		btn_next_pane2.setBounds(4, 178, 190, 50);
		panel_7.add(btn_next_pane2);

		btnNewButton_2 = new GradientButton("Migration Repoart");
		btnNewButton_2.setEnabled(false);
		btnNewButton_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				fifth = true;

				Buttonclick("btnNewButton_2", fifth);
				
				btn_previous_p2.setEnabled(true);
				btn_Next.setEnabled(true);

				if (table_fileConvertionreport_panel4.getRowCount() > 0) {

					CardLayout card = (CardLayout) Cardlayout.getLayout();
					card.show(Cardlayout, "panel_4");
					
					
					
					if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")
							|| filetype.equalsIgnoreCase("OFFICE 365"))) {

						System.out.println("here we riched");
						textField_username_p3.setEnabled(true);
						passwordField_p3.setEnabled(true);
						tf_portNo_p3.setEnabled(true);
						chckbxShowPassword_p3.setEnabled(true);
						lblPortNo.setEnabled(true);
						lblNewLabel_5.setEnabled(true);
						lblNewLabel_1.setEnabled(true);
						lblNewLabel.setEnabled(true);
						lblemailAddress.setEnabled(true);
						lblemailAddress.setEnabled(false);

					}
					textField_customfolder.setEnabled(true);
					archive.setEnabled(true);
					publicfolder.setEnabled(true);
					mailbox.setEnabled(true);
					modern_Authentication.setEnabled(true);
					basic_Authentication.setEnabled(true);

//	                           basic_Authentication.setSelected(true);
					filetype = "";
					path = "";
					checkconvertagain = true;
					stop = false;
					btnStop.setVisible(false);
					chckbxSavePdfAttachment.setEnabled(true);
					btn_Destination.setEnabled(true);
					btn_previous_p3.setEnabled(true);
					chckbxSaveInSame.setEnabled(true);
					comboBox_fileDestination_type.setEnabled(true);
					btn_Destination.setEnabled(true);
					btn_previous_p3.setEnabled(true);
					lbl_progressreport.setText("");
					comboBox.setEnabled(true);
					chckbx_splitpst.setEnabled(true);
					textField_customfolder.setEditable(true);
					chckbxMigrateOrBackup.setEnabled(true);
					dateChooser_calender_start.setEnabled(true);
					chckbx_convert_pdf_to_pdf.setEnabled(true);
					chckbxRemoveDuplicacy.setEnabled(true);
					datefilter.setEnabled(true);
					dateChooser_calendar_end.setEnabled(true);
					chckbxCustomFolderName.setEnabled(true);
					btn_Destination.setEnabled(true);
					btn_previous_p3.setEnabled(true);
					checkmboxpstost = true;
					chckbxRestoreToDefault.setEnabled(true);
					chckbxSaveMboxIn.setEnabled(true);
					btn_signout_p3.setVisible(false);
					label_11.setVisible(false);
					panel_5.setEnabled(true);
					chckbx_Mail_Filter.setSelected(false);
					chckbx_Mail_Filter.setEnabled(true);
					chckbx_calender_box.setEnabled(true);
					task_box.setEnabled(true);
					chckbxMaintainFolderStructure.setEnabled(true);
					textField_domain_name_p3.setEnabled(true);
					textField_username_p3.setEnabled(true);
					passwordField_p3.setEnabled(true);
					tf_portNo_p3.setEnabled(true);
					lblNewLabel.setEnabled(true);
					lblemailAddress.setEnabled(true);
					lblNewLabel_1.setEnabled(true);
					lblNewLabel_5.setEnabled(true);
					lblPortNo.setEnabled(true);
//					CardLayout card1 = (CardLayout) Cardlayout.getLayout();
//					card1.show(Cardlayout, "panel_2");
					
					
					
					
				} else {

					JOptionPane.showMessageDialog(mf,
							"No Any repoart available to view please migrate first after then you can view repoart ",
							messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));

				}

			}
		});
		btnNewButton_2.setShadowColor(new Color(0, 0, 204));

		btnNewButton_2.setDefaultCapable(false);
		btnNewButton_2.setContentAreaFilled(false);
		btnNewButton_2.setBorderPainted(false);
		btnNewButton_2.setRequestFocusEnabled(false);
		btnNewButton_2.setOpaque(false);
		btnNewButton_2.setRolloverEnabled(false);
		btnNewButton_2.setFocusable(false);
		btnNewButton_2.setFocusTraversalKeysEnabled(false);
		btnNewButton_2.setFocusPainted(false);

		btnNewButton_2.setFont(new Font("Verdana", Font.BOLD, 12));
		btnNewButton_2.setBounds(5, 285, 190, 50);

		btnNewButton_2.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btnNewButton_2.isEnabled()) {

					btnNewButton_2.setGradientColor1(new Color(70, 130, 180));
					btnNewButton_2.setGradientColor2(new Color(70, 130, 180));
					btnNewButton_2.setForeground(new Color(255, 255, 255));
				}
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnNewButton_2.setGradientColor1(new Color(255, 255, 255));
				btnNewButton_2.setGradientColor2(new Color(255, 255, 255));
				btnNewButton_2.setForeground(new Color(80, 80, 80));
			}
		});

		panel_7.add(btnNewButton_2);

		btn_converter_1.setShadowColor(new Color(0, 0, 204));
		btn_converter_1.setRolloverEnabled(false);
		btn_converter_1.setRequestFocusEnabled(false);
		btn_converter_1.setOpaque(false);
		btn_converter_1.setFont(new Font("Verdana", Font.BOLD, 12));
		btn_converter_1.setFocusable(false);
		btn_converter_1.setFocusTraversalKeysEnabled(false);
		btn_converter_1.setFocusPainted(false);
		btn_converter_1.setDefaultCapable(false);
		btn_converter_1.setContentAreaFilled(false);
		btn_converter_1.setBorderPainted(false);
		btn_converter_1.setBounds(5, 231, 190, 50);
		panel_7.add(btn_converter_1);
		// label_4.setIcon(new ImageIcon(Main_Frame.class.getResource("/topbar.png")));

		techHelp = new JButton("");
		techHelp.setBackground(new Color(189, 93, 232));
		techHelp.setBounds(8, 338, 80, 80);
		panel_7.add(techHelp);
		techHelp.setToolTipText("Click here for Technical Support");
		techHelp.setRolloverEnabled(false);
		techHelp.setRequestFocusEnabled(false);
		techHelp.setOpaque(false);
		techHelp.setFocusable(false);
		techHelp.setFocusTraversalKeysEnabled(false);
		techHelp.setFocusPainted(false);
		techHelp.setDefaultCapable(false);
		techHelp.setBorderPainted(false);
		techHelp.setContentAreaFilled(false);

		techHelp.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				openBrowser(All_Data.helpuri);
			}
		});
		techHelp.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				techHelp.setIcon(new ImageIcon(Main_Frame.class.getResource("/livechat_hrv.png")));

			}

			public void mouseExited(MouseEvent e) {
				techHelp.setIcon(new ImageIcon(Main_Frame.class.getResource("/livechat.png")));
			}
		});
		techHelp.setIcon(new ImageIcon(Main_Frame.class.getResource("/livechat.png")));
		// techHelp.setIcon(new
		// ImageIcon(Main_Frame.class.getResource("/live-chat-btn1.png")));

		panel_7.add(btn_buy);
		btn_buy.setToolTipText("Click here to buy the Software.");
		btn_buy.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				openBrowser(buyurl);

			}
		});

		btn_buy.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_buy.setIcon(new ImageIcon(Main_Frame.class.getResource("/buy-btn_hrv.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_buy.setIcon(new ImageIcon(Main_Frame.class.getResource("/buy-btn.png")));
			}
		});

		btn_buy.setIcon(new ImageIcon(Main_Frame.class.getResource("/buy-btn.png")));

		btn_buy.setOpaque(false);
		btn_buy.setRolloverEnabled(false);
		btn_buy.setRequestFocusEnabled(false);
		btn_buy.setFocusTraversalKeysEnabled(false);
		btn_buy.setFocusable(false);
		btn_buy.setFocusPainted(false);
		btn_buy.setDefaultCapable(false);
		btn_buy.setContentAreaFilled(false);
		btn_buy.setBorderPainted(false);

		panel_7.add(btnActivate);
//		btnActivate.addMouseListener(new MouseAdapter() {
//			@Override
//			public void mouseEntered(MouseEvent arg0) {
//				btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/key-act-btn_hrv.png")));
//			}
//
//			@Override
//			public void mouseExited(MouseEvent e) {
//				btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/key-act-btn.png")));
//			}
//		});

		
		btnActivate.setRolloverEnabled(false);
		btnActivate.setRequestFocusEnabled(false);
		btnActivate.setOpaque(false);
		btnActivate.setFocusable(false);
		btnActivate.setFocusTraversalKeysEnabled(false);
		btnActivate.setFocusPainted(false);
		btnActivate.setDefaultCapable(false);
		btnActivate.setContentAreaFilled(false);
		btnActivate.setBorderPainted(false);

		JButton btn_help = new JButton("");
		btn_help.setBounds(8, 521, 77, 80);
		panel_7.add(btn_help);
		btn_help.setToolTipText("Click here For Software Guide.");
		btn_help.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				openBrowser(helpurl);
			}
		});
		btn_help.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_help.setIcon(new ImageIcon(Main_Frame.class.getResource("/about-btn_hrv.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_help.setIcon(new ImageIcon(Main_Frame.class.getResource("/about-btn.png")));
			}
		});

		btn_help.setIcon(new ImageIcon(Main_Frame.class.getResource("/about-btn.png")));

		btn_help.setRolloverEnabled(false);
		btn_help.setRequestFocusEnabled(false);
		btn_help.setOpaque(false);
		btn_help.setFocusable(false);
		btn_help.setFocusTraversalKeysEnabled(false);
		btn_help.setFocusPainted(false);
		btn_help.setDefaultCapable(false);
		btn_help.setContentAreaFilled(false);
		btn_help.setBorderPainted(false);

		JButton btn_info = new JButton("");
		btn_info.setBounds(5, 429, 80, 80);
		panel_7.add(btn_info);
		btn_info.setToolTipText("Click here For Software Information.");
		btn_info.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				AboutDialog ab;
				if (demo) {
					ab = new AboutDialog(mf, true, "Demo");

				} else {
					ab = new AboutDialog(mf, true, "full");
				}
				ab.setLocationRelativeTo(mf);
				ab.setVisible(true);

			}
		});
		btn_info.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_info.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-btn_hrv.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_info.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-btn.png")));
			}
		});

		btn_info.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-btn.png")));
		btn_info.setRolloverEnabled(false);
		btn_info.setRequestFocusEnabled(false);
		btn_info.setOpaque(false);
		btn_info.setFocusable(false);
		btn_info.setFocusTraversalKeysEnabled(false);
		btn_info.setFocusPainted(false);
		btn_info.setDefaultCapable(false);
		btn_info.setContentAreaFilled(false);
		btn_info.setBorderPainted(false);
		
		 btn_Next_1 = new GradientButton("Destination Type");
		btn_Next_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				
				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_9");
				
			}
		});
		btn_Next_1.setText("Destination Type");
		btn_Next_1.setToolTipText("select Destination Type.");
		btn_Next_1.setShadowColor(new Color(0, 0, 204));
		btn_Next_1.setRolloverEnabled(false);
		btn_Next_1.setRequestFocusEnabled(false);
		btn_Next_1.setOpaque(false);
		btn_Next_1.setGradientColor1(Color.WHITE);
		btn_Next_1.setFont(new Font("Verdana", Font.BOLD, 12));
		btn_Next_1.setFocusable(false);
		btn_Next_1.setFocusTraversalKeysEnabled(false);
		btn_Next_1.setFocusPainted(false);
		btn_Next_1.setEnabled(false);
		btn_Next_1.setDefaultCapable(false);
		btn_Next_1.setContentAreaFilled(false);
		btn_Next_1.setBorderPainted(false);
		btn_Next_1.setBounds(5, 129, 190, 50);
		panel_7.add(btn_Next_1);
		btnActivate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				try {
					if (demo) {
						if (System.getProperty("os.name").toLowerCase().contains("windows")) {
							licFileon = new File(System.getenv("APPDATA") + File.separator + All_Data.messageboxtitle
									+ File.separator + "licenseOnline.lic");
						} else {
							licFileon = new File(System.getProperty("user.home") + File.separator + "Library"
									+ File.separator + "Application Support" + File.separator + All_Data.messageboxtitle
									+ File.separator + "licenseOnline.lic");
						}
						boolean activatefromdemo = true;
						//new Starting_Frame();
						
						
						ActivationFrame mf = new ActivationFrame();

						mf.setLocationRelativeTo(null);
						mf.setVisible(true);
						setVisible(false);
						
						
						
//						OnlineActivation mf = new OnlineActivation(Starting_Frame.mf, licFileon, activatefromdemo);
//						mf.setLocationRelativeTo(null);
//						mf.setVisible(true);
//						mf.btnBack.setVisible(false);
//						mf.addWindowListener(new WindowAdapter() {
//							@Override
//							public void windowClosing(WindowEvent arg0) {
//								String warn = "Do you want to close?";
//								int ans = JOptionPane.showConfirmDialog(mf, warn, All_Data.messageboxtitle,
//										JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE,
//										new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
//								if (ans == JOptionPane.YES_OPTION) {
//									setEnabled(true);
//									mf.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
//								} else {
//									mf.setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
//								}
//							}
//						});
						All_Data.messageboxtitle = All_Data.messageboxbackup;
//						All_Data.messageboxtitle = projectTitle;
					} else {
						String warn = "Do you want to Deactivate the Software?";
						int ans = JOptionPane.showConfirmDialog(main_multiplefile.this, warn, All_Data.messageboxtitle,
								JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
						if (ans == JOptionPane.YES_OPTION) {
							Starting_Frame sf = new Starting_Frame();
							new Uninstall(Starting_Frame.ToolUri, All_Data.messageboxtitle,
									Starting_Frame.activationKey, Starting_Frame.orderId);
							dispose();

							Main_Frame frame = new Main_Frame(true, 4);
							frame.setLocationRelativeTo(null);
							frame.temppath = frame.textField_1.getText();
							frame.cal = Calendar.getInstance();
							frame.calendertime = frame
									.getRidOfIllegalFileNameCharacters(frame.cal.getTime().toString());

							main_multiplefile multi = new main_multiplefile(frame, frame.demo, frame.messageboxtitle);
							multi.setLocationRelativeTo(null);
							multi.setVisible(true);

//							sf.setLocationRelativeTo(null);
//							sf.setResizable(false);
//							sf.setVisible(true);
						}
					}
				} catch (Exception e1) {
					mf.logger.warning(e1.getMessage() + System.lineSeparator());
				} catch (Error e1) {
					mf.logger.warning(e1.getMessage() + System.lineSeparator());
				}

			}

		});

	}

	protected void lisstOfFiles(File[] files) {

		for (File b : files) {
			if (b.getName().endsWith(".pst") || b.getName().endsWith("PST")) {
				hashset.add(b);
			}
			if (b.listFiles() != null) {
				files = b.listFiles();
				lisstOfFiles(files);
			}
		}

	}

	void filter_file() throws Exception {

		jFileChooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");
		files = null;

		jFileChooser.setMultiSelectionEnabled(true);
		jFileChooser.setAcceptAllFileFilterUsed(false);
		// jFileChooser.setAcceptAllFileFilterUsed(false);
		FileNameExtensionFilter filter;
		if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

			filter = new FileNameExtensionFilter(".ost", "ost");

			jFileChooser.addChoosableFileFilter(filter);

		} else if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {

			filter = new FileNameExtensionFilter(".pst", "pst");

			jFileChooser.addChoosableFileFilter(filter);

		} else if (fileoptionm.equalsIgnoreCase("DBX")) {

			filter = new FileNameExtensionFilter(".dbx", "DBX");

			jFileChooser.addChoosableFileFilter(filter);

		} else if (fileoptionm.equalsIgnoreCase("EML File (.eml)")) {
			filter = new FileNameExtensionFilter(".eml", "eml");

			jFileChooser.addChoosableFileFilter(filter);

		} else if (fileoptionm.equalsIgnoreCase("OFT File (.oft)")) {
			filter = new FileNameExtensionFilter(".oft", "oft");

			jFileChooser.addChoosableFileFilter(filter);

		} else if (fileoptionm.equalsIgnoreCase("EMLX File (.emlx)")) {
			filter = new FileNameExtensionFilter(".emlx", "emlx");

			jFileChooser.addChoosableFileFilter(filter);

		} else if (fileoptionm.equalsIgnoreCase("Message File (.msg)")) {
			filter = new FileNameExtensionFilter(".msg", "msg");

			jFileChooser.addChoosableFileFilter(filter);

		} else if (fileoptionm.equalsIgnoreCase("MBOX")) {

			jFileChooser.setFileFilter(new FileNameExtensionFilter(".mbox", "mbx", "mbox"));
			jFileChooser.setAcceptAllFileFilterUsed(true);

		} else if (fileoptionm.equalsIgnoreCase("Maildir")) {

			jFileChooser.setAcceptAllFileFilterUsed(true);

		} else if (fileoptionm.equalsIgnoreCase("OLM File (.olm)")) {
			filter = new FileNameExtensionFilter(".olm", "olm");

			jFileChooser.addChoosableFileFilter(filter);

		}

		if (jFileChooser.showOpenDialog(main_multiplefile.this) == JFileChooser.APPROVE_OPTION) {

			files = jFileChooser.getSelectedFiles();

			for (int i = 0; i < files.length; i++) {
				String extension = getFileExtension(files[i]);
				if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
						|| fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

					try {

						if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
							if (extension.equalsIgnoreCase("pst")) {

								hashset.add(files[i]);

							}
						} else if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

							if (extension.equalsIgnoreCase("ost")) {
								hashset.add(files[i]);
							}
						}

					} catch (Exception e1) {

						e1.printStackTrace();
					}

				} else if (fileoptionm.equalsIgnoreCase("MBOX")) {
					hashset.add(files[i]);
				} else if (fileoptionm.equalsIgnoreCase("DBX")) {
					hashset.add(files[i]);
				} else if (fileoptionm.equalsIgnoreCase("OLM File (.olm)")) {

					hashset.add(files[i]);
				} else {
					if (fileoptionm.equalsIgnoreCase("EML File (.eml)")) {
						if (extension.equals("eml")) {

							hashset.add(files[i]);

						}
					} else if (fileoptionm.equalsIgnoreCase("EMLX File (.emlx)")) {
						if (extension.equalsIgnoreCase("emlx")) {

							hashset.add(files[i]);
						}
					} else if (fileoptionm.equalsIgnoreCase("OFT File (.oft)")) {
						if (extension.equalsIgnoreCase("oft")) {

							hashset.add(files[i]);
						}
					} else if (fileoptionm.equalsIgnoreCase("Message File (.msg)")) {
						if (extension.equalsIgnoreCase("msg")) {

							hashset.add(files[i]);
						}
					} else if (fileoptionm.equalsIgnoreCase("Maildir")) {

						hashset.add(files[i]);

					}

				}

			}
			DefaultTableModel model = (DefaultTableModel) table.getModel();

			while (model.getRowCount() > 0) {

				for (int i = 0; i < model.getRowCount(); ++i) {

					model.removeRow(i);
					filesno--;
				}
			}

			Iterator<File> itr = hashset.iterator();
			while (itr.hasNext()) {

				modeli = (DefaultTableModel) table.getModel();
				File fo = itr.next();
				String filet = "";
				if (fo.isFile()) {
					filet = "File";
				} else {
					filet = "Folder";
				}
				long sizeInBytes = fo.length();
				modeli.addRow(
						new Object[] { filesno, fo.getName(), fo.getAbsolutePath(), filet, bytes2String(sizeInBytes) });
				filesno++;
				countforfile++;
			}
		}

	}

	void destinationPath() throws Exception {
		jFileChooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");

		jFileChooser.setMultiSelectionEnabled(true);

		jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

		jFileChooser.showOpenDialog(main_multiplefile.this);
		checkdestination = false;
		File file = jFileChooser.getSelectedFile();

		String destination = file.getAbsolutePath();

		tf_Destination_Location.setText(destination);

	}

	private static void expandAllNodes() {
		int j = tree.getRowCount();
		int i = 0;
		while (i < j) {
			tree.expandRow(i);
			i += 1;

		}
	}

	public void readAnOST_PstFile() {

		try {
			try {

				pst = PersonalStorage.fromFile(filepath);
			} catch (Exception e1) {
				cropted = false;
				if (e1.getMessage().contains("File not found File")) {
					JOptionPane.showMessageDialog(mf, "File is in use ", messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));
					return;
				} else {
					cropted = true;
					loadingDialog.dispose();
					JLabel lb = new JLabel(
							"<html><b>Your File is Corrupted !<br/><h4> You need to use:-<u><html><b style=\"color:red;\">Aryson Outlook PST Repair</u></h4><html><b> Either Contact Support !</html><b>");
					lb.setToolTipText("click here to go Aryson Outlook PST Repair");
					lb.addMouseListener(new MouseAdapter() {
						@Override
						public void mouseClicked(MouseEvent e) {
							openBrowser(All_Data.pstreccoverylink);
						}
					});
					String[] buttons = { "Support", "Home" };
					int response = JOptionPane.showOptionDialog(ld, lb, Main_Frame.messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE, 0,
							new ImageIcon(Main_Frame.class.getResource("/information.png")), buttons, buttons[0]);
					if (response == 0) {
						System.out.println("No button clicked");
						openBrowser(All_Data.helpuri);
					} else {
						openBrowser(All_Data.pstreccoverylink);
					}
					return;
				}

			}

			FolderInfoCollection folderInfoCollection = pst.getRootFolder().getSubFolders();
			foldercountcheck++;

			String filepat = filepath.replace(",", "");
			CustomTreeNode e = new CustomTreeNode("<html><b>" + filepat);

			e.filepath = filepath;
			mainnode.add(e);

			FolderInfo folderInfo1 = pst.getRootFolder();
//			String rootname = folderInfo1.getDisplayName().replaceAll("[\\[\\]],", "").replace(".", "");
			String rootname = folderInfo1.getDisplayName().replaceAll("[\\[\\]],", "");
			if (rootname.equalsIgnoreCase("")) {
				rootname = "Root Folder";
			}

			DefaultMutableTreeNode node1 = new DefaultMutableTreeNode("<html><b>" + rootname);
			e.add(node1);

			for (int i = 0; i < folderInfoCollection.size(); i++) {
				if (mf.stop_tree) {
					break;
				}
				FolderInfo folderInfo = (FolderInfo) folderInfoCollection.get_Item(i);

				String foldername = folderInfo.getDisplayName();
				foldername = foldername.replace(",", "").replace(".", "");
//				.replace(".", "") this has been deleted from above line;
//				foldername = foldername.replace(",", "").replace(".", "");
				foldername = getRidOfIllegalFileNameCharacters(foldername);
				foldername = foldername.replaceAll("[\\[\\]]", "");
				foldername = foldername.trim();
				if (!foldername.equals("Root - Public")) {
					DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + foldername);
					node1.add(node);
//					obTh.ob.MessageLabel.setText(foldername);

					loadingDialog.MessageLabel.setText(foldername);

					if (folderInfo.hasSubFolders()) {

						readOstpstsubfolder(folderInfo, node);

					}
				}
			}
		} catch (Exception e) {
			if (e.getMessage().contains("File not found File")) {
				JOptionPane.showMessageDialog(mf, "File is in use ", messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));
			}

			else {
				JOptionPane.showMessageDialog(mf, "File is Currupted  Please Choose another file  ", messageboxtitle,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));
			}
		}
	}

	public void readOstpstsubfolder(FolderInfo f, DefaultMutableTreeNode node) {

		FolderInfoCollection folderCollection = f.getSubFolders();

		for (int i = 0; i < folderCollection.size(); i++) {

			if (mf.stop_tree) {
				break;
			}
			FolderInfo folderInfo = (FolderInfo) folderCollection.get_Item(i);
			String foldername = folderInfo.getDisplayName();
			foldername = foldername.replace(",", "").replace(".", "");

			foldername = getRidOfIllegalFileNameCharacters(foldername);
			foldername = foldername.replaceAll("[\\[\\]]", "");
			foldername = foldername.trim();

			DefaultMutableTreeNode nod1 = new DefaultMutableTreeNode(

					"<html><b>" + foldername);

			node.add(nod1);

//			obTh.ob.MessageLabel.setText(foldername);

			loadingDialog.MessageLabel.setText(foldername);
			if (folderInfo.hasSubFolders()) {

				readOstpstsubfolder(folderInfo, nod1);

			}

		}
	}

	public void readolmFile() {
		OlmStorage storage = null;

		try {
			storage = new OlmStorage(filepath);
			String filepat = filepath.replace(",", "");
			CustomTreeNode e = new CustomTreeNode("<html><b>" + filepat);
			e.filepath = filepath;
			mainnode.add(e);
			foldercountcheck++;
			try {
				for (OlmFolder folder : storage.getFolderHierarchy()) {

					if (mf.stop_tree) {
						break;
					}
					String foldername = folder.getName().replaceAll("[\\[\\]],", "");
					DefaultMutableTreeNode c = new DefaultMutableTreeNode("<html><b>" + foldername);

					e.add(c);
//					obTh.ob.MessageLabel.setText(foldername);
					loadingDialog.MessageLabel.setText(foldername);
					if (folder.getSubFolders().size() > 0) {

						getFolder(folder, c);

					}

				}

			} catch (Exception e1) {

				return;
			} finally {
				storage.dispose();

			}
		} catch (Exception e) {
			if (e.getMessage().contains("File not found File")) {
				JOptionPane.showMessageDialog(mf, "File is in use ", messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));
			}

			else {
				JOptionPane.showMessageDialog(mf, "File is Currupted  Please Choose another file  ", messageboxtitle,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));
			}
		}

	}

	private void getFolder(OlmFolder folder, DefaultMutableTreeNode node1) {

		for (OlmFolder subFolder : folder.getSubFolders()) {

			if (mf.stop_tree) {
				break;
			}
			String foldername = subFolder.getName().replaceAll("[\\[\\]],", "").replaceAll("[\\[\\]]", "");
			DefaultMutableTreeNode nd = new DefaultMutableTreeNode("<html><b>" + foldername);
			node1.add(nd);
//			obTh.ob.MessageLabel.setText(foldername);
			loadingDialog.MessageLabel.setText(foldername);
			if (subFolder.getSubFolders().size() > 0) {

				getFolder(subFolder, nd);

			}

		}
	}

	public void readMboxFile() {

		file = new File(filepath);

		String filepath = filepath(file);

		visitAllNodes(mainnode);
		String filenamembox = file.getName();
		foldercountcheck++;
		if (listst.contains(filepath)) {
			DefaultMutableTreeNode nd = null;
			for (int k = 0; k < lists.size(); k++) {

				if (listst.get(k).equalsIgnoreCase(filepath)) {
					nd = lists.get(k);
					break;
				}

			}

			CustomTreeNode child = new CustomTreeNode("<html><b>" + filenamembox);
			child.filepath = file.getAbsolutePath();
			nd.add(child);

		} else {
			DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + filepath);
			mainnode.add(node);

			CustomTreeNode child = new CustomTreeNode("<html><b>" + filenamembox);
			child.filepath = file.getAbsolutePath();
			node.add(child);
		}

	}

	public void readmailFile() {

		file = new File(filepath);

		String filepath = filepath(file);
		foldercountcheck++;
		visitAllNodes(mainnode);
		String filenamemail = file.getName().replaceAll("[\\[\\]]", "");
		if (listst.contains(filepath)) {
			DefaultMutableTreeNode nd = null;
			for (int k = 0; k < lists.size(); k++) {

				if (listst.get(k).equalsIgnoreCase(filepath)) {
					nd = lists.get(k);
					break;
				}

			}

			CustomTreeNode child = new CustomTreeNode("<html><b>" + filenamemail);
			child.filepath = file.getAbsolutePath();
			nd.add(child);

		} else {

			DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + filepath);

			mainnode.add(node);

			CustomTreeNode child = new CustomTreeNode("<html><b>" + filenamemail);
			child.filepath = file.getAbsolutePath();
			node.add(child);
		}
//		obTh.ob.MessageLabel.setText(filenamemail);
		loadingDialog.MessageLabel.setText(foldername);
	}

	void readewwds(File c) {

		File[] files = c.listFiles();

		int messagesize = files.length;

		for (int j = 0; j < messagesize; j++) {
			if (Stoppreview) {
				break;
			}
			if (files[j].isDirectory()) {

			} else {

				String extension = getFileExtension(files[j]);
				if (fileoptionm.equalsIgnoreCase("EML File (.eml)")) {
					if (extension.equals("eml")) {

						path2 = files[j].getAbsolutePath();

						fileInformation_on_mail();

					}
				} else if (fileoptionm.equalsIgnoreCase("EMLX File (.emlx)")) {
					if (extension.equalsIgnoreCase("emlx")) {

						path2 = files[j].getAbsolutePath();

						fileInformation_on_mail();

					}
				} else if (fileoptionm.equalsIgnoreCase("OFT File (.oft)")) {
					if (extension.equalsIgnoreCase("oft")) {

						path2 = files[j].getAbsolutePath();

						fileInformation_on_mail();

					}
				} else if (fileoptionm.equalsIgnoreCase("Apple Mail")) {
					if (extension.equalsIgnoreCase("emlx")) {

						path2 = files[j].getAbsolutePath();

						fileInformation_on_mail();

					}
				} else if (fileoptionm.equalsIgnoreCase("Message File (.msg)")) {
					if (extension.equalsIgnoreCase("msg")) {

						path2 = files[j].getAbsolutePath();

						fileInformation_on_mail();
					}
				} else if (fileoptionm.equalsIgnoreCase("Maildir")) {

					path2 = files[j].getAbsolutePath();

					fileInformation_on_mail();

				}

			}

		}
	}

	public void fileInformation_on_mail() {
		MailMessage message = null;

		try {

			MailMessage message1 = MailMessage.load(path2);

			MailConversionOptions option = new MailConversionOptions();
			MapiMessage msg = MapiMessage.fromMailMessage(message1, MapiConversionOptions.getASCIIFormat());
			message = msg.toMailMessage(option);

			String from = "";
			String Subject = "";
			String Date = "";

			listmail.add(message1);

			try {
				from = message.getFrom().toString();
			} catch (Exception e) {

			}
			try {
				Subject = message.getSubject();
			} catch (Exception e) {

			}
			try {
				Date = message1.getDate().toString();
			} catch (Exception e) {

			}
			lblTotalMessageCount.setText("Total Message Count : " + ids);
			ids++;

			if (message.getAttachments().size() > 0) {
				ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
				JLabel imagelabl = new JLabel();
				imagelabl.setIcon(icon);
				mode = (DefaultTableModel) table_fileinformation.getModel();

				mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date, imagelabl });
			} else {
				mode = (DefaultTableModel) table_fileinformation.getModel();

				mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date });
			}
		} catch (Exception e) {
			JOptionPane.showMessageDialog(mf, "File is Currupted  Please Choose another file  " + filepath,
					messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
					new ImageIcon(Main_Frame.class.getResource("/information.png")));

		}
	}

	public void fileInhformation_Ost_Pst() throws Exception {
		FolderInfo f1 = pst.getRootFolder();
		String f1nmae = f1.getDisplayName().replaceAll("[\\[\\]]", "");
		if (f1nmae.equalsIgnoreCase("")) {
			f1nmae = "Root Folder";
		}
		if (foldername.equals(f1nmae)) {
			MessageInfoCollection messageInfoCollection = f1.getContents();

			lblTotalMessageCount.setText("<html><b>" + "  Total Message Count : " + messageInfoCollection.size());

			for (int j = 0; j < messageInfoCollection.size(); j++)

			{

				try {
					if (Stoppreview) {
						break;
					}

					MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(j);
					MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
					MailConversionOptions de = new MailConversionOptions();
					MapiMessage message1 = (MapiMessage) pst.extractMessage(messageInfo);
					MailMessage mess = message1.toMailMessage(de);
					MapiMessage message = MapiMessage.fromMailMessage(mess, d);

					listmapi.add(message1);
					String from = "";
					String Subject = "";
					String DeliveryTime = "";
					try {
						from = message1.getSenderEmailAddress();
					} catch (Exception a) {
						from = "NA";
					}
					try {
						Subject = message1.getSubject();
					} catch (Exception a) {
						Subject = "NA";
					}
					try {
						DeliveryTime = message1.getDeliveryTime().toString();
					} catch (Exception a) {
						DeliveryTime = "NA";
					}

					if (message1.getAttachments().size() > 0) {
						ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
						JLabel imagelabl = new JLabel();
						imagelabl.setIcon(icon);
						mode = (DefaultTableModel) table_fileinformation.getModel();

						mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
								"<html><b>" + DeliveryTime, message1.getAttachments().size(), imagelabl });
					} else {
						mode = (DefaultTableModel) table_fileinformation.getModel();

						mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
								"<html><b>" + DeliveryTime, 0 });
					}

				} catch (Exception e) {
					continue;
				}

			}

			path = "";

		}

		FolderInfoCollection folderInfoCollection = pst.getRootFolder().getSubFolders();

		for (int i = 0; i < folderInfoCollection.size(); i++) {
			try {
				if (Stoppreview) {
					break;
				}
				FolderInfo f = folderInfoCollection.get_Item(i);
				String Folder = f.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();
				path = f1nmae + File.separator + Folder;
				int size = f.getContentCount();

				foldername = foldername.replace("[" + size + "]", "");

				if (foldername.equals(path)) {
					folderInfo = folderInfoCollection.get_Item(i);
					MessageInfoCollection messageInfoCollection = f.getContents();

					lblTotalMessageCount
							.setText("<html><b>" + "  Total Message Count : " + messageInfoCollection.size());

					for (int j = 0; j < messageInfoCollection.size(); j++)

					{
						try {

							if (Stoppreview) {
								break;
							}
							MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(j);

							MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
							MailConversionOptions de = new MailConversionOptions();
							MapiMessage message1 = (MapiMessage) pst.extractMessage(messageInfo);
							MailMessage mess = message1.toMailMessage(de);
							MapiMessage message = MapiMessage.fromMailMessage(mess, d);

							listmapi.add(message1);
							listPSTOSTgemesingo.add(messageInfo);

							String from = "";
							String Subject = "";
							String DeliveryTime = "";
							try {
								from = message1.getSenderEmailAddress();
							} catch (Exception a) {
								from = "NA";
							}
							try {
								Subject = message1.getSubject();
							} catch (Exception a) {
								Subject = "NA";
							}
							try {
								DeliveryTime = message1.getDeliveryTime().toString();
							} catch (Exception a) {
								DeliveryTime = "NA";
							}

							if (message1.getAttachments().size() > 0) {
								ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime, message1.getAttachments().size(), imagelabl });
							} else {
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime, 0 });
							}
						} catch (Exception e) {
							continue;
						}

					}

					path = "";
					break;

				}
				if (f.hasSubFolders()) {
					fileInhformationsubfolder_Ost_Pst(f);

				}
			} catch (Exception e) {
				continue;
			}

		}

	}

	public void fileInhformationsubfolder_Ost_Pst(FolderInfo folder) {

		FolderInfoCollection folderInfoCollection = folder.getSubFolders();

		for (int i = 0; i < folderInfoCollection.size(); i++) {
			try {

				FolderInfo folderInf = folderInfoCollection.get_Item(i);
				int size = folderInf.getContentCount();
				foldername = foldername.replace("[" + size + "]", "");

				String Folder = folderInf.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();

				path = path + File.separator + Folder;

				if (foldername.equals(path)) {
					folderInfo = folderInfoCollection.get_Item(i);
					MessageInfoCollection messageInfoCollection = folderInf.getContents();
					lblTotalMessageCount
							.setText("<html><b>" + "  Total Message Count : " + messageInfoCollection.size());
					for (int j = 0; j < messageInfoCollection.size(); j++)

					{

						try {
							if (Stoppreview) {
								break;
							}

							MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(j);

							MapiMessage message1 = (MapiMessage) pst.extractMessage(messageInfo);
							MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
							MailConversionOptions de = new MailConversionOptions();
							MailMessage mess = message1.toMailMessage(de);
							MapiMessage message = MapiMessage.fromMailMessage(mess, d);
							listmapi.add(message1);
							String from = "NA";
							String Subject = "NA";
							Date DeliveryTime = null;
							try {
								from = message1.getSenderEmailAddress();
							} catch (Exception a) {
								from = "";
							}
							try {
								Subject = message1.getSubject();
							} catch (Exception a) {
								Subject = "";
							}
							try {
								DeliveryTime = message1.getDeliveryTime();
							} catch (Exception a) {

							}
							if (message1.getAttachments().size() > 0) {
								ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);
								mode = (DefaultTableModel) table_fileinformation.getModel();
//								lblTotalMessageCount.setText("<html><b>" + "  Total Message Count : " + (j + 1));
								mode.addRow(new Object[] { from, Subject, DeliveryTime,
										message1.getAttachments().size(), imagelabl });
							} else {
								mode = (DefaultTableModel) table_fileinformation.getModel();
//								lblTotalMessageCount.setText("<html><b>" + "  Total Message Count : " + (j + 1));
								mode.addRow(new Object[] { from, Subject, DeliveryTime, 0 });
							}

						} catch (Exception e) {
							continue;
						}

					}
					
					try {

						rowSorter = new TableRowSorter<>(table_fileinformation.getModel());
						table_fileinformation.setRowSorter(rowSorter);

						searchField.getDocument().addDocumentListener(new DocumentListener() {

							@Override
							public void insertUpdate(DocumentEvent e) {
								String text = searchField.getText();

								if (text.trim().length() == 0) {
									rowSorter.setRowFilter(null);
								} else {
									rowSorter.setRowFilter(RowFilter.regexFilter("(?i)" + text));

								}
							}

							@Override
							public void removeUpdate(DocumentEvent e) {
								String text = searchField.getText();

								if (text.trim().length() == 0) {
									rowSorter.setRowFilter(null);
								} else {
									rowSorter.setRowFilter(RowFilter.regexFilter("(?i)" + text));
								}
							}

							@Override
							public void changedUpdate(DocumentEvent e) {
								throw new UnsupportedOperationException("Not supported yet."); // To change body of
																								// generated methods,
																								// choose Tools |
																								// Templates.

							}

						});
						
						
						table_fileinformation.addMouseListener(new MouseAdapter() {
					            @Override
					            public void mouseClicked(MouseEvent e) {
					                int viewRow = table_fileinformation.getSelectedRow(); // Get view index
					                System.out.println(viewRow);
					                if (viewRow >= 0) {
					                     modelRow = table_fileinformation.convertRowIndexToModel(viewRow); // Convert to model index
					                    System.out.println("Selected model row: " + modelRow);
					                    
					                   //MapiMessage msg= listmapi.get(modelRow);				                    
					                }
					            }
					        });
						
						
					} catch (Exception e) {
						System.out.println("error during searching 9074");
						// TODO Auto-generated catch block
						// e.printStackTrace();
					}
					
					
					
					
					
					

					path = "";
					break;

				}
				if (folderInf.hasSubFolders()) {
					fileInhformationsubfolder_Ost_Pst(folderInf);

				}
				path = removefolder(path);

			} catch (Exception e) {
				continue;
			}
		}

	}

	public void fileInformation_on_mbox() {
		MboxrdStorageReader mbox = null;

		try {
			FileStream stream = new FileStream(path2, FileMode.OpenOrCreate, FileAccess.Read);

			mbox = new MboxrdStorageReader(stream.toInputStream(), false);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(mf, "File is Currupted  Please Choose another file  " + filepath,
					messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
					new ImageIcon(Main_Frame.class.getResource("/information.png")));

		}

		MailMessage message1 = mbox.readNextMessage();
		int i2 = 1;
		while (message1 != null) {

			if (Stoppreview) {
				break;
			}
			MailConversionOptions option = new MailConversionOptions();
			MapiMessage msg = MapiMessage.fromMailMessage(message1, MapiConversionOptions.getASCIIFormat());
			MailMessage message = msg.toMailMessage(option);

			String from = message.getFrom().toString();

			String Subject = message.getSubject();

			String Date = message.getDate().toString();
			lblTotalMessageCount.setText("Total Message Count :" + i2);
			i2++;
			listmail.add(message);
			if (message.getAttachments().size() > 0) {
				ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
				JLabel imagelabl = new JLabel();
				imagelabl.setIcon(icon);
				mode = (DefaultTableModel) table_fileinformation.getModel();

				mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date, imagelabl });
			} else {
				mode = (DefaultTableModel) table_fileinformation.getModel();

				mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date });
			}

			try {
				message1 = mbox.readNextMessage();

			} catch (Exception e) {
				continue;
			}

		}

	}

	void readapple_mail(File filearray, DefaultMutableTreeNode node) throws Exception {

		File[] files = filearray.listFiles();

		for (int i = 0; i < files.length; i++) {

			if (files[i].isFile()) {
				String extension = "";
				try {
					extension = getFileExtension(files[i]);
				} catch (Exception e) {
					extension = "";
				}
				if (files[i].length() > 0) {
					if (extension.equalsIgnoreCase("emlx")) {

						DefaultMutableTreeNode t = new DefaultMutableTreeNode("<html><b>" + files[i].getName());

						node.add(t);
					}
				}
			} else {

				File[] fo = files[i].listFiles();

				if (fo.length > 0) {
					if (!files[i].getName().equalsIgnoreCase("MailData")) {
						foldername = files[i].getName();

						DefaultMutableTreeNode t = new DefaultMutableTreeNode(foldername);

						node.add(t);

						readapple_mail(files[i], t);
					}
				}
			}

		}

	}

	public void fileinformation_olm() {
		OlmStorage storage = new OlmStorage(path2);
		// System.out.println("hello");

		for (OlmFolder folder : storage.getFolderHierarchy()) {

			String pa1 = folder.getName().replaceAll("[\\[\\]]", "");

			if (folder.getName().equals(foldername)) {

				Iterator<MapiMessage> it = storage.enumerateMessages(folder).iterator();
				if (Stoppreview) {
					break;
				}
				if (folder.hasMessages()) {
					while (it.hasNext()) {
						MapiMessage msg1 = it.next();
						MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
						MailConversionOptions de = new MailConversionOptions();

						MailMessage mess = msg1.toMailMessage(de);
						MapiMessage msg = MapiMessage.fromMailMessage(mess, d);

						listmapi.add(msg);
						String from = msg.getSenderEmailAddress();

						Date DeliveryTime = msg.getDeliveryTime();

						String Subject = msg.getSubject();

						if (msg.getAttachments().size() > 0) {
							ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
							JLabel imagelabl = new JLabel();
							imagelabl.setIcon(icon);
							mode = (DefaultTableModel) table_fileinformation.getModel();

							mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
									"<html><b>" + DeliveryTime, imagelabl });
						} else {
							mode = (DefaultTableModel) table_fileinformation.getModel();

							mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
									"<html><b>" + DeliveryTime });
						}

					}
				}
			}

			if (folder.getSubFolders().size() > 0) {

				getFolderolminfo(folder, pa1, storage);

			}

		}

	}

	private void getFolderolminfo(OlmFolder folder, String rootFolder, OlmStorage storage) {

		for (OlmFolder subFolder : folder.getSubFolders()) {

			String curpath = rootFolder + File.separator + subFolder.getName().replaceAll("[\\[\\]]", "");

			if (curpath.equals(foldername)) {

				if (subFolder.hasMessages()) {
					for (MapiMessage msg1 : storage.enumerateMessages(subFolder)) {
						try {
							if (Stoppreview) {
								break;
							}

							MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
							MailConversionOptions de = new MailConversionOptions();

							MailMessage mess = msg1.toMailMessage(de);
							MapiMessage msg = MapiMessage.fromMailMessage(mess, d);

							listmapi.add(msg);
							String from = msg.getSenderEmailAddress();

							Date DeliveryTime = msg.getDeliveryTime();

							String Subject = msg.getSubject();

							if (msg.getAttachments().size() > 0) {
								ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime, imagelabl });
							} else {
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime });
							}
						} catch (Exception e) {
							mf.logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
							continue;
						}
					}
				}
				
				break;

			}

			if (subFolder.getSubFolders().size() > 0) {

				getFolderolminfo(subFolder, curpath, storage);

			}

			curpath = removefolder(curpath);
		}

	}

	public static void getTree(String folderName) {

		DefaultMutableTreeNode node = null;

		if (folderName.contains(File.separator)) {

			String parts[] = folderName.split(Matcher.quoteReplacement(File.separator));

			DefaultMutableTreeNode par = new DefaultMutableTreeNode(parts[parts.length - 2]);
			{
				DefaultMutableTreeNode child = new DefaultMutableTreeNode(parts[parts.length - 1]);

				search(root, par);

				lastNode.add(child);

				expandAllNodes();
			}
		}

		else {
			node = new DefaultMutableTreeNode(folderName);
			model.insertNodeInto(node, root, root.getChildCount());

			lastNode = node;
			expandAllNodes();
		}

	}

	void read_mbox_folder(File filearray) {
		File[] files = filearray.listFiles();

		for (int j = 0; j < files.length; j++) {

			if (files[j].isDirectory()) {

				read_mbox_folder(files[j]);

			} else {

				String extension = getFileExtension(files[j]);

				if (extension.equalsIgnoreCase("mbox") || extension.equalsIgnoreCase("mbx") || extension.equals("")) {
					filepath = files[j].getAbsolutePath();

					readMboxFile();

				}

			}

		}

	}

	void reademl_emlx_msg_folder(File filearray) {
		File[] files = filearray.listFiles();

		int messagesize = files.length;

		for (int j = 0; j < messagesize; j++) {

			if (files[j].isDirectory()) {

				reademl_emlx_msg_folder(files[j]);

			} else {

				String extension = getFileExtension(files[j]);
				if (fileoptionm.equalsIgnoreCase("EML File (.eml)")) {
					if (extension.equals("eml")) {

						filepath = files[j].getAbsolutePath();

						readmailFile();

					}
				} else if (fileoptionm.equalsIgnoreCase("EMLX File (.emlx)")) {
					if (extension.equalsIgnoreCase("emlx")) {

						filepath = files[j].getAbsolutePath();

						readmailFile();

					}
				} else if (fileoptionm.equalsIgnoreCase("OFT File (.oft)")) {
					if (extension.equalsIgnoreCase("oft")) {

						filepath = files[j].getAbsolutePath();

						readmailFile();

					}
				} else if (fileoptionm.equalsIgnoreCase("Apple Mail")) {
					if (extension.equalsIgnoreCase("emlx")) {

						filepath = files[j].getAbsolutePath();

						readmailFile();

					}
				} else if (fileoptionm.equalsIgnoreCase("Message File (.msg)")) {
					if (extension.equalsIgnoreCase("msg")) {

						filepath = files[j].getAbsolutePath();

						readmailFile();
					}
				} else if (fileoptionm.equalsIgnoreCase("Maildir")) {

					filepath = files[j].getAbsolutePath();

					readmailFile();

				}
			}
		}
	}

	void read_olm_folder(File filearray) {
		File[] files = filearray.listFiles();

		for (int j = 0; j < files.length; j++) {

			if (files[j].isDirectory()) {

				read_olm_folder(files[j]);

			} else {

				String extension = getFileExtension(files[j]);

				if (extension.equalsIgnoreCase("olm")) {
					filepath = files[j].getAbsolutePath();

					readolmFile();

				}
			}
		}
	}

	void read_PSTOST_folder(File filearray) {
		File[] files = filearray.listFiles();

		for (int j = 0; j < files.length; j++) {

			if (files[j].isDirectory()) {

				read_PSTOST_folder(files[j]);

			} else {

				String extension = getFileExtension(files[j]);
				if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
					if (extension.equalsIgnoreCase("pst")) {
						filepath = files[j].getAbsolutePath();

						readAnOST_PstFile();

					}

				} else if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

					if (extension.equalsIgnoreCase("ost")) {
						filepath = files[j].getAbsolutePath();

						readAnOST_PstFile();

					}
				}
			}

		}

	}

	public void Mapimess_CSV(MapiMessage message, CSVWriter writer) {

		String subname = getRidOfIllegalFileNameCharacters(mf.namingconventionmapi(message));
		if (message.getMessageClass().equals("IPM.Contact")) {
			String Email1addresstype = null;
			String Email1displayname = null;
			String Email1address = null;
			String Email1fax = null;
			String Email2addresstype = null;
			String Email2displayname = null;
			String Email2address = null;
			String Email2fax = null;
			String Email3addresstype = null;
			String Email3displayname = null;
			String Email3address = null;
			String Email3fax = null;
			String homefaxaddresstype = null;
			String homefaxdisplayname = null;
			String homefaxaddress = null;
			String homefaxno = null;

			String primaryfaxdisplayname = null;
			String primaryfaxaddress = null;
			String primaryfaxno = null;
			String primaryfaxaddresstype = null;
			String bussinessfaxaddresstype = null;
			String bussinessfaxaddress = null;
			String bussinessfaxno = null;
			String bussinessfaxdisplayname = null;
			String birthday = null;
			String WeddingAnniversary = null;
			String firstName = null;
			String middleName = null;
			String lastname = null;
			String prefix = null;
			String fileunder = null;
			String fileunderid = null;
			String generation = null;
			String title = null;
			String account = null;
			String BusinessHomePage = null;
			String ComputerNetworkName = null;
			String CustomerId = null;
			String FreeBusyLocation = null;
			String FtpSite = null;
			String Gender = null;
			String GovernmentIdNumber = null;
			String Hobbies = null;
			String Html = null;
			String InstantMessagingAddress = null;
			String Language = null;
			String Location = null;
			String Notes = null;
			String OrganizationalIdNumber = null;
			String PersonalHomePage = null;
			String ReferredByName = null;
			String SpouseName = null;
			String homeAddress = null;
			String homeCity = null;
			String homeCountry = null;
			String homeCountryCode = null;
			String homePostalCode = null;
			String homegetPostOfficeBox = null;
			String homeStateOrProvince = null;
			String homegetStreet = null;
			String otherAddress = null;
			String otherCity = null;
			String otherCountry = null;
			String otherCountryCode = null;
			String otherPostalCode = null;
			String othergetPostOfficeBox = null;
			String otherStateOrProvince = null;
			String othergetStreet = null;
			String workAddress = null;
			String workCity = null;
			String workCountry = null;
			String workCountryCode = null;
			String workPostalCode = null;
			String workgetPostOfficeBox = null;
			String workStateOrProvince = null;
			String workgetStreet = null;
			String Assistant = null;
			String CompanyName = null;
			String DepartmentName = null;
			String ManagerName = null;
			String OfficeLocation = null;
			String Profession = null;
			String getTitle = null;
			String AssistantTelephoneNumber = null;
			String Business2TelephoneNumber = null;
			String BusinessTelephoneNumber = null;
			String CallbackTelephoneNumber = null;
			String CarTelephoneNumber = null;
			String CompanyMainTelephoneNumber = null;
			String Home2TelephoneNumber = null;
			String HomeTelephoneNumber = null;
			String IsdnNumber = null;
			String MobileTelephoneNumber = null;
			String OtherTelephoneNumber = null;
			String PagerTelephoneNumber = null;
			String PrimaryTelephoneNumber = null;
			String RadioTelephoneNumber = null;
			String TelexNumber = null;
			String TtyTddPhoneNumber = null;

			MapiContact con = (MapiContact) message.toMapiMessageItem();
			MapiContactProfessionalPropertySet ProfPropSet = null;
			try {
				ProfPropSet = con.getProfessionalInfo();
			} catch (Exception e) {

			}
			MapiContactElectronicAddress email1 = null;
			try {
				email1 = con.getElectronicAddresses().getEmail1();
			} catch (Exception e) {

			}
			MapiContactElectronicAddress email2 = null;
			try {
				email2 = con.getElectronicAddresses().getEmail2();
			} catch (Exception e) {

			}
			MapiContactElectronicAddress email3 = null;
			try {
				email3 = con.getElectronicAddresses().getEmail3();
			} catch (Exception e) {

			}
			MapiContactElectronicAddress homefax = null;
			try {
				homefax = con.getElectronicAddresses().getHomeFax();
			} catch (Exception e) {

			}
			MapiContactElectronicAddress primaryfax = null;
			try {
				primaryfax = con.getElectronicAddresses().getPrimaryFax();
			} catch (Exception e) {

			}
			MapiContactElectronicAddress bussinessfax = null;
			try {
				bussinessfax = con.getElectronicAddresses().getBusinessFax();
			} catch (Exception e) {

			}
			MapiContactEventPropertySet event = null;
			try {
				event = con.getEvents();
			} catch (Exception e) {

			}
			MapiContactPersonalInfoPropertySet personfPropSet = null;
			try {
				personfPropSet = con.getPersonalInfo();
			} catch (Exception e) {

			}
			MapiContactProfessionalPropertySet personPropSet = null;
			try {
				personPropSet = con.getProfessionalInfo();
			} catch (Exception e) {

			}
			MapiContactPhysicalAddressPropertySet mapipcs = null;
			try {
				mapipcs = con.getPhysicalAddresses();
			} catch (Exception e) {

			}
			MapiContactTelephonePropertySet mapitelephone = null;
			try {
				mapitelephone = con.getTelephones();
			} catch (Exception e) {

			}
			MapiContactPhysicalAddress contacthomephys = null;
			try {
				contacthomephys = mapipcs.getHomeAddress();
			} catch (Exception e) {

			}
			MapiContactPhysicalAddress contactotherphys = null;
			try {
				contactotherphys = mapipcs.getOtherAddress();
			} catch (Exception e) {

			}
			MapiContactPhysicalAddress contactworkphys = null;
			try {
				contactworkphys = mapipcs.getWorkAddress();
			} catch (Exception e) {

			}
			MapiContactNamePropertySet NamePropSet = null;
			try {
				NamePropSet = con.getNameInfo();
			} catch (Exception e) {

			}
			try {

				title = ProfPropSet.getTitle();
			} catch (Exception ep) {
				title = "";
			}

			try {
				if (title.equalsIgnoreCase("null") || title.contains("meta") || title.contains("aspose")) {
					title = "NA";
				}
			} catch (Exception e1) {
				title = "NA";
			}

			try {

				firstName = NamePropSet.getGivenName();
			} catch (Exception ep) {
				firstName = "";
			}
			try {
				if (firstName.equalsIgnoreCase("null") || firstName.contains("meta") || firstName.contains("aspose")) {
					firstName = "NA";
				}
			} catch (Exception e1) {
				firstName = "NA";
			}

			try {

				middleName = NamePropSet.getMiddleName();
			} catch (Exception ep) {
				middleName = "";
			}
			try {
				if (middleName.equalsIgnoreCase("null") || middleName.contains("meta")
						|| middleName.contains("aspose")) {
					middleName = "NA";
				}
			} catch (Exception e1) {
				middleName = "NA";
			}
			try {

				lastname = NamePropSet.getSurname();
			} catch (Exception ep) {
				lastname = "";
			}
			try {
				if (lastname.equalsIgnoreCase("null") || lastname.contains("meta") || lastname.contains("aspose")) {
					lastname = "NA";
				}
			} catch (Exception e1) {
				lastname = "NA";
			}
			try {

				Email1addresstype = email1.getAddressType();
			} catch (Exception ep) {
				Email1addresstype = "";
			}
			try {
				if (Email1addresstype.equalsIgnoreCase("null") || Email1addresstype.contains("meta")
						|| Email1addresstype.contains("aspose")) {
					Email1addresstype = "NA";
				}
			} catch (Exception e1) {
				Email1addresstype = "";
			}
			try {

				Email1displayname = email1.getDisplayName();
			} catch (Exception ep) {
				Email1displayname = "";
			}
			try {
				if (Email1displayname.equalsIgnoreCase("null") || Email1displayname.contains("meta")
						|| Email1displayname.contains("aspose")) {
					Email1displayname = "NA";
				}
			} catch (Exception e1) {
				Email1displayname = "";
			}
			try {

				Email1address = email1.getEmailAddress();
			} catch (Exception ep) {
				Email1address = "";
			}
			try {
				if (Email1address.equalsIgnoreCase("null") || Email1address.contains("meta")
						|| Email1address.contains("aspose")) {

				}
			} catch (Exception e1) {
				Email1address = "NA";
			}
			try {

				Email1fax = email1.getFaxNumber();
			} catch (Exception ep) {
				Email1fax = "";
			}
			try {
				if (Email1fax.equalsIgnoreCase("null") || Email1fax.contains("meta") || Email1fax.contains("aspose")) {
					Email1fax = "NA";
				}
			} catch (Exception e1) {
				Email1fax = "NA";
			}
			try {

				Email2addresstype = email2.getAddressType();
			} catch (Exception ep) {
				Email2addresstype = "";
			}
			try {
				if (Email2addresstype.equalsIgnoreCase("null") || Email2addresstype.contains("meta")
						|| Email2addresstype.contains("aspose")) {
					Email2addresstype = "NA";
				}
			} catch (Exception e1) {
				Email2addresstype = "NA";
			}
			try {

				Email2displayname = email2.getDisplayName();
			} catch (Exception ep) {
				Email2displayname = "";
			}
			try {
				if (Email2displayname.equalsIgnoreCase("null") || Email2displayname.contains("meta")
						|| Email2displayname.contains("aspose")) {
					Email2displayname = "NA";
				}
			} catch (Exception e1) {
				Email2displayname = "NA";
			}
			try {

				Email2address = email2.getEmailAddress();
			} catch (Exception ep) {
				Email1address = "";
			}
			try {
				if (Email1address.equalsIgnoreCase("null") || Email1address.contains("meta")
						|| Email1address.contains("aspose")) {
					Email1address = "NA";
				}
			} catch (Exception e1) {
				Email1address = "NA";
			}
			try {

				Email2fax = email2.getFaxNumber();
			} catch (Exception ep) {
				Email2fax = "";
			}
			try {
				if (Email2fax.equalsIgnoreCase("null") || Email2fax.contains("meta") || Email2fax.contains("aspose")) {
					Email2fax = "NA";
				}
			} catch (Exception e1) {
				Email2fax = "NA";
			}
			try {

				Email3addresstype = email3.getAddressType();
			} catch (Exception ep) {
				Email3addresstype = "";
			}
			try {
				if (Email3addresstype.equalsIgnoreCase("null") || Email3addresstype.contains("meta")
						|| Email3addresstype.contains("aspose")) {
					Email3addresstype = "NA";
				}
			} catch (Exception e1) {
				Email3addresstype = "NA";
			}
			try {

				Email3displayname = email3.getDisplayName();
			} catch (Exception ep) {
				Email3displayname = "";
			}
			try {
				if (Email3displayname.equalsIgnoreCase("null") || Email3displayname.contains("meta")
						|| Email3displayname.contains("aspose")) {
					Email3displayname = "NA";
				}
			} catch (Exception e1) {
				Email3displayname = "NA";
			}
			try {

				Email3address = email3.getEmailAddress();
			} catch (Exception ep) {
				Email3address = "";
			}
			try {
				if (Email3address.equalsIgnoreCase("null") || Email3address.contains("meta")
						|| Email3address.contains("aspose")) {
					Email3address = "NA";
				}
			} catch (Exception e1) {
				Email3address = "NA";
			}
			try {

				Email3fax = email3.getFaxNumber();
			} catch (Exception ep) {
				Email3fax = "";
			}
			try {
				if (Email3fax.equalsIgnoreCase("null") || Email3fax.contains("meta") || Email3fax.contains("aspose")) {
					Email3fax = "NA";
				}
			} catch (Exception e1) {
				Email3fax = "NA";
			}
			try {

				homefaxaddresstype = homefax.getAddressType();
			} catch (Exception ep) {
				homefaxaddresstype = "";
			}
			try {
				if (homefaxaddresstype.equalsIgnoreCase("null") || homefaxaddresstype.contains("meta")
						|| homefaxaddresstype.contains("aspose")) {
					homefaxaddresstype = "NA";
				}
			} catch (Exception e1) {
				homefaxaddresstype = "NA";
			}
			try {

				homefaxdisplayname = homefax.getDisplayName();
			} catch (Exception ep) {
				homefaxdisplayname = "";
			}
			try {
				if (homefaxdisplayname.equalsIgnoreCase("null") || homefaxdisplayname.contains("meta")
						|| homefaxdisplayname.contains("aspose")) {
					homefaxdisplayname = "NA";
				}
			} catch (Exception e1) {
				homefaxdisplayname = "NA";
			}
			try {

				homefaxaddress = homefax.getEmailAddress();
			} catch (Exception ep) {
				homefaxaddress = "";
			}
			try {
				if (homefaxaddress.equalsIgnoreCase("null") || homefaxaddress.contains("meta")
						|| homefaxaddress.contains("aspose")) {
					homefaxaddress = "NA";
				}
			} catch (Exception e1) {
				homefaxaddress = "NA";
			}
			try {

				homefaxno = homefax.getFaxNumber();
			} catch (Exception ep) {
				homefaxno = "";
			}
			try {
				if (homefaxno.equalsIgnoreCase("null") || homefaxno.contains("meta") || homefaxno.contains("aspose")) {
					homefaxno = "NA";
				}
			} catch (Exception e1) {
				homefaxno = "NA";
			}
			try {

				primaryfaxaddresstype = primaryfax.getAddressType();
			} catch (Exception ep) {
				primaryfaxaddresstype = "";
			}
			try {
				if (primaryfaxaddresstype.equalsIgnoreCase("null") || primaryfaxaddresstype.contains("meta")
						|| primaryfaxaddresstype.contains("aspose")) {
					primaryfaxaddresstype = "NA";
				}
			} catch (Exception e1) {
				primaryfaxaddresstype = "NA";
			}
			try {

				primaryfaxdisplayname = primaryfax.getDisplayName();
			} catch (Exception ep) {
				primaryfaxdisplayname = "";
			}
			try {
				if (primaryfaxdisplayname.equalsIgnoreCase("null") || primaryfaxdisplayname.contains("meta")
						|| primaryfaxdisplayname.contains("aspose")) {
					primaryfaxdisplayname = "NA";
				}
			} catch (Exception e1) {
				primaryfaxdisplayname = "NA";
			}
			try {

				primaryfaxaddress = primaryfax.getEmailAddress();
			} catch (Exception ep) {
				primaryfaxaddress = "";
			}
			try {
				if (primaryfaxaddress.equalsIgnoreCase("null") || primaryfaxaddress.contains("meta")
						|| primaryfaxaddress.contains("aspose")) {
					primaryfaxaddress = "NA";
				}
			} catch (Exception e1) {
				primaryfaxaddress = "NA";
			}
			try {

				primaryfaxno = primaryfax.getFaxNumber();
			} catch (Exception ep) {
				primaryfaxno = "";
			}
			try {
				if (primaryfaxno.equalsIgnoreCase("null") || primaryfaxno.contains("meta")
						|| primaryfaxno.contains("aspose")) {
					primaryfaxno = "NA";
				}
			} catch (Exception e1) {
				primaryfaxno = "NA";
			}
			try {

				bussinessfaxaddresstype = bussinessfax.getAddressType();
			} catch (Exception ep) {
				bussinessfaxaddresstype = "";
			}
			try {
				if (bussinessfaxaddresstype.equalsIgnoreCase("null") || bussinessfaxaddresstype.contains("meta")
						|| bussinessfaxaddresstype.contains("aspose")) {
					bussinessfaxaddresstype = "NA";
				}
			} catch (Exception e1) {
				bussinessfaxaddresstype = "NA";
			}
			try {

				bussinessfaxdisplayname = bussinessfax.getDisplayName();
			} catch (Exception ep) {
				bussinessfaxdisplayname = "";
			}
			try {
				if (bussinessfaxdisplayname.equalsIgnoreCase("null") || bussinessfaxdisplayname.contains("meta")
						|| bussinessfaxdisplayname.contains("aspose")) {
					bussinessfaxdisplayname = "NA";
				}
			} catch (Exception e1) {
				bussinessfaxdisplayname = "NA";
			}
			try {

				bussinessfaxaddress = bussinessfax.getEmailAddress();
			} catch (Exception ep) {
				bussinessfaxaddress = "";
			}
			try {
				if (bussinessfaxaddress.equalsIgnoreCase("null") || bussinessfaxaddress.contains("meta")
						|| bussinessfaxaddress.contains("aspose")) {
					bussinessfaxaddress = "NA";
				}
			} catch (Exception e1) {
				bussinessfaxaddress = "NA";
			}
			try {

				bussinessfaxno = bussinessfax.getFaxNumber();
			} catch (Exception ep) {
				bussinessfaxno = "";
			}
			try {
				if (bussinessfaxno.equalsIgnoreCase("null") || bussinessfaxno.contains("meta")
						|| bussinessfaxno.contains("aspose")) {
					bussinessfaxno = "NA";
				}
			} catch (Exception e1) {
				bussinessfaxno = "NA";
			}
			try {

				birthday = String.valueOf(event.getBirthday());
			} catch (Exception ep) {
				birthday = "";
			}
			try {
				if (birthday.equalsIgnoreCase("null") || birthday.contains("meta") || birthday.contains("aspose")) {
					birthday = "NA";
				}
			} catch (Exception e1) {
				birthday = "NA";
			}
			try {

				WeddingAnniversary = String.valueOf(event.getWeddingAnniversary());
			} catch (Exception ep) {
				WeddingAnniversary = "";
			}
			try {
				if (WeddingAnniversary.equalsIgnoreCase("null") || WeddingAnniversary.contains("meta")
						|| WeddingAnniversary.contains("aspose")) {
					WeddingAnniversary = "NA";
				}
			} catch (Exception e1) {
				WeddingAnniversary = "NA";
			}
			try {

				Email1addresstype = email1.getAddressType();
			} catch (Exception ep) {
				Email1addresstype = "";
			}
			try {
				if (Email1addresstype.equalsIgnoreCase("null") || Email1addresstype.contains("meta")
						|| firstName.contains("aspose")) {
					Email1addresstype = "NA";
				}
			} catch (Exception e1) {
				Email1addresstype = "NA";
			}
			try {

				prefix = NamePropSet.getDisplayNamePrefix();
			} catch (Exception ep) {
				prefix = "";
			}
			try {
				if (prefix.equalsIgnoreCase("null") || prefix.contains("meta") || prefix.contains("aspose")) {
					prefix = "NA";
				}
			} catch (Exception e1) {
				prefix = "NA";
			}

			try {

				fileunder = NamePropSet.getFileUnder();
			} catch (Exception ep) {
				fileunder = "";
			}
			try {
				if (fileunder.equalsIgnoreCase("null") || fileunder.contains("meta") || fileunder.contains("aspose")) {
					fileunder = "NA";
				}
			} catch (Exception e1) {
				fileunder = "NA";
			}
			try {

				fileunderid = String.valueOf(NamePropSet.getFileUnderID());
			} catch (Exception ep) {
				fileunderid = "";
			}
			try {
				if (fileunderid.equalsIgnoreCase("null") || fileunderid.contains("meta")
						|| fileunderid.contains("aspose")) {
					fileunderid = "NA";
				}
			} catch (Exception e1) {
				fileunderid = "NA";
			}
			try {

				generation = String.valueOf(NamePropSet.getGeneration());
			} catch (Exception ep) {
				generation = "";
			}
			try {
				if (generation.equalsIgnoreCase("null") || generation.contains("meta")
						|| generation.contains("aspose")) {
					generation = "NA";
				}
			} catch (Exception e1) {
				generation = "NA";
			}
			try {

				account = String.valueOf(personfPropSet.getAccount());
			} catch (Exception ep) {
				account = "";
			}
			try {
				if (account.equalsIgnoreCase("null") || account.contains("meta") || account.contains("aspose")) {
					account = "NA";
				}
			} catch (Exception e1) {
				account = "NA";
			}
			try {

				BusinessHomePage = String.valueOf(personfPropSet.getBusinessHomePage());
			} catch (Exception ep) {
				BusinessHomePage = "";
			}
			try {
				if (BusinessHomePage.equalsIgnoreCase("null") || BusinessHomePage.contains("meta")
						|| BusinessHomePage.contains("aspose")) {
					BusinessHomePage = "NA";
				}
			} catch (Exception e1) {
				BusinessHomePage = "NA";
			}
			try {

				ComputerNetworkName = String.valueOf(personfPropSet.getComputerNetworkName());
			} catch (Exception ep) {
				ComputerNetworkName = "";
			}
			try {
				if (ComputerNetworkName.equalsIgnoreCase("null") || ComputerNetworkName.contains("meta")
						|| ComputerNetworkName.contains("aspose")) {
					ComputerNetworkName = "NA";
				}
			} catch (Exception e1) {
				ComputerNetworkName = "NA";
			}
			try {

				CustomerId = String.valueOf(personfPropSet.getCustomerId());
			} catch (Exception ep) {
				CustomerId = "";
			}
			try {
				if (CustomerId.equalsIgnoreCase("null") || CustomerId.contains("meta")
						|| CustomerId.contains("aspose")) {
					CustomerId = "NA";
				}
			} catch (Exception e1) {
				CustomerId = "NA";
			}
			try {

				FreeBusyLocation = String.valueOf(personfPropSet.getFreeBusyLocation());
			} catch (Exception ep) {
				FreeBusyLocation = "";
			}
			try {
				if (FreeBusyLocation.equalsIgnoreCase("null") || FreeBusyLocation.contains("meta")
						|| FreeBusyLocation.contains("aspose")) {
					FreeBusyLocation = "NA";
				}
			} catch (Exception e1) {
				FreeBusyLocation = "NA";
			}
			try {

				FtpSite = String.valueOf(personfPropSet.getFtpSite());
			} catch (Exception ep) {
				FtpSite = "";
			}
			try {
				if (FtpSite.equalsIgnoreCase("null") || FtpSite.contains("meta") || FtpSite.contains("aspose")) {
					FtpSite = "NA";
				}
			} catch (Exception e1) {
				FtpSite = "NA";
			}
			try {
				int i = personfPropSet.getGender();
				if (i == 0) {
					Gender = "Unspecified";
				} else if (i == 1) {
					Gender = "Female";
				} else if (i == 2) {
					Gender = "Male";
				}

			} catch (Exception ep) {
				Gender = "";
			}

			try {

				GovernmentIdNumber = String.valueOf(personfPropSet.getGovernmentIdNumber());
			} catch (Exception ep) {
				GovernmentIdNumber = "";
			}
			try {
				if (GovernmentIdNumber.equalsIgnoreCase("null") || GovernmentIdNumber.contains("meta")
						|| GovernmentIdNumber.contains("aspose")) {
					GovernmentIdNumber = "NA";
				}
			} catch (Exception e1) {
				GovernmentIdNumber = "NA";
			}
			try {

				Hobbies = String.valueOf(personfPropSet.getHobbies());
			} catch (Exception ep) {
				Hobbies = "";
			}
			try {
				if (Hobbies.equalsIgnoreCase("null") || Hobbies.contains("meta") || Hobbies.contains("aspose")) {
					Hobbies = "NA";
				}
			} catch (Exception e1) {
				Hobbies = "NA";
			}
			try {

				Html = String.valueOf(personfPropSet.getHtml());
			} catch (Exception ep) {
				Html = "";
			}
			try {
				if (Html.equalsIgnoreCase("null") || Html.contains("aspose")) {
					Html = "NA";
				}
			} catch (Exception e1) {
				Html = "NA";
			}
			try {

				InstantMessagingAddress = String.valueOf(personfPropSet.getInstantMessagingAddress());
			} catch (Exception ep) {
				InstantMessagingAddress = "";
			}
			try {
				if (InstantMessagingAddress.equalsIgnoreCase("null") || InstantMessagingAddress.contains("meta")
						|| InstantMessagingAddress.contains("aspose")) {
					InstantMessagingAddress = "NA";
				}
			} catch (Exception e1) {
				InstantMessagingAddress = "NA";
			}

			try {

				Language = String.valueOf(personfPropSet.getLanguage());
			} catch (Exception ep) {
				Language = "";
			}
			try {
				if (Language.equalsIgnoreCase("null") || Language.contains("meta") || Language.contains("aspose")) {
					Language = "NA";
				}
			} catch (Exception e1) {
				Language = "NA";
			}
			try {

				Location = String.valueOf(personfPropSet.getLocation());
			} catch (Exception ep) {
				Location = "";
			}
			try {
				if (Location.equalsIgnoreCase("null") || Location.contains("meta") || Location.contains("aspose")) {
					Location = "NA";
				}
			} catch (Exception e1) {
				Location = "NA";
			}
			try {

				Notes = String.valueOf(personfPropSet.getNotes());
			} catch (Exception ep) {
				Notes = "";
			}
			try {
				if (Notes.equalsIgnoreCase("null") || Notes.contains("meta") || Notes.contains("aspose")) {
					Notes = "NA";
				}
			} catch (Exception e1) {
				Notes = "NA";
			}
			try {

				OrganizationalIdNumber = String.valueOf(personfPropSet.getOrganizationalIdNumber());
			} catch (Exception ep) {
				OrganizationalIdNumber = "";
			}
			try {
				if (OrganizationalIdNumber.equalsIgnoreCase("null") || OrganizationalIdNumber.contains("meta")
						|| OrganizationalIdNumber.contains("aspose")) {
					OrganizationalIdNumber = "NA";
				}
			} catch (Exception e1) {
				OrganizationalIdNumber = "NA";
			}
			try {

				PersonalHomePage = String.valueOf(personfPropSet.getPersonalHomePage());
			} catch (Exception ep) {
				PersonalHomePage = "";
			}
			try {
				if (PersonalHomePage.equalsIgnoreCase("null") || PersonalHomePage.contains("meta")
						|| PersonalHomePage.contains("aspose")) {
					PersonalHomePage = "NA";
				}
			} catch (Exception e1) {
				PersonalHomePage = "NA";
			}
			try {

				ReferredByName = String.valueOf(personfPropSet.getReferredByName());
			} catch (Exception ep) {
				ReferredByName = "";
			}
			try {
				if (ReferredByName.equalsIgnoreCase("null") || ReferredByName.contains("meta")
						|| ReferredByName.contains("aspose")) {
					ReferredByName = "NA";
				}
			} catch (Exception e1) {
				ReferredByName = "NA";
			}
			try {

				SpouseName = String.valueOf(personfPropSet.getSpouseName());
			} catch (Exception ep) {
				SpouseName = "";
			}
			try {
				if (SpouseName.equalsIgnoreCase("null") || SpouseName.contains("meta")
						|| SpouseName.contains("aspose")) {
					SpouseName = "NA";
				}
			} catch (Exception e1) {
				SpouseName = "NA";
			}

			try {

				homegetStreet = String.valueOf(contacthomephys.getStreet());
			} catch (Exception ep) {
				homegetStreet = "";
			}
			try {
				if (homegetStreet.equalsIgnoreCase("null") || homegetStreet.contains("meta")
						|| homegetStreet.contains("aspose")) {
					homegetStreet = "NA";
				}
			} catch (Exception e1) {
				homegetStreet = "NA";
			}
			try {

				homeAddress = String.valueOf(contacthomephys.getAddress());
			} catch (Exception ep) {
				homeAddress = "";
			}
			try {
				if (homeAddress.equalsIgnoreCase("null") || homeAddress.contains("meta")
						|| homeAddress.contains("aspose")) {
					homeAddress = "NA";
				}
			} catch (Exception e1) {
				homeAddress = "NA";
			}
			try {

				homeCity = String.valueOf(contacthomephys.getCity());
			} catch (Exception ep) {
				homeCity = "";
			}
			try {
				if (homeCity.equalsIgnoreCase("null") || homeCity.contains("meta") || homeCity.contains("aspose")) {
					homeCity = "NA";
				}
			} catch (Exception e1) {
				homeCity = "NA";
			}
			try {

				homeCountry = String.valueOf(contacthomephys.getCountry());
			} catch (Exception ep) {
				homeCountry = "";
			}
			try {
				if (homeCountry.equalsIgnoreCase("null") || homeCountry.contains("meta")
						|| homeCountry.contains("aspose")) {
					homeCountry = "NA";
				}
			} catch (Exception e1) {
				homeCountry = "NA";
			}
			try {

				homeCountryCode = String.valueOf(contacthomephys.getCountryCode());
			} catch (Exception ep) {
				homeCountryCode = "";
			}
			try {
				if (homeCountryCode.equalsIgnoreCase("null") || homeCountryCode.contains("meta")
						|| homeCountryCode.contains("aspose")) {
					homeCountryCode = "NA";
				}
			} catch (Exception e1) {
				homeCountryCode = "NA";
			}
			try {

				homePostalCode = String.valueOf(contacthomephys.getPostalCode());
			} catch (Exception ep) {
				homePostalCode = "";
			}
			try {
				if (homePostalCode.equalsIgnoreCase("null") || homePostalCode.contains("meta")
						|| homePostalCode.contains("aspose")) {
					homePostalCode = "NA";
				}
			} catch (Exception e1) {
				homePostalCode = "NA";
			}
			try {

				homegetPostOfficeBox = String.valueOf(contacthomephys.getPostOfficeBox());
			} catch (Exception ep) {
				homegetPostOfficeBox = "";
			}
			try {
				if (homegetPostOfficeBox.equalsIgnoreCase("null") || homegetPostOfficeBox.contains("meta")
						|| homegetPostOfficeBox.contains("aspose")) {
					homegetPostOfficeBox = "NA";
				}
			} catch (Exception e1) {
				homegetPostOfficeBox = "NA";
			}
			try {

				homeStateOrProvince = String.valueOf(contacthomephys.getStateOrProvince());
			} catch (Exception ep) {
				homeStateOrProvince = "";
			}
			try {
				if (homeStateOrProvince.equalsIgnoreCase("null") || homeStateOrProvince.contains("meta")
						|| homeStateOrProvince.contains("aspose")) {
					homeStateOrProvince = "NA";
				}
			} catch (Exception e1) {
				homeStateOrProvince = "NA";
			}
			try {

				othergetStreet = String.valueOf(contactotherphys.getStreet());
			} catch (Exception ep) {
				othergetStreet = "";
			}
			try {
				if (othergetStreet.equalsIgnoreCase("null") || othergetStreet.contains("meta")
						|| othergetStreet.contains("aspose")) {
					othergetStreet = "NA";
				}
			} catch (Exception e1) {
				othergetStreet = "NA";
			}
			try {

				otherAddress = String.valueOf(contactotherphys.getAddress());
			} catch (Exception ep) {
				otherAddress = "";
			}
			try {
				if (otherAddress.equalsIgnoreCase("null") || otherAddress.contains("meta")
						|| otherAddress.contains("aspose")) {
					otherAddress = "NA";
				}
			} catch (Exception e1) {
				otherAddress = "NA";
			}
			try {

				otherCity = String.valueOf(contactotherphys.getCity());
			} catch (Exception ep) {
				otherCity = "";
			}
			try {
				if (otherCity.equalsIgnoreCase("null") || otherCity.contains("meta") || otherCity.contains("aspose")) {
					otherCity = "NA";
				}
			} catch (Exception e1) {
				otherCity = "NA";
			}
			try {

				otherCountry = String.valueOf(contactotherphys.getCountry());
			} catch (Exception ep) {
				otherCountry = "";
			}
			try {
				if (otherCountry.equalsIgnoreCase("null") || otherCountry.contains("meta")
						|| otherCountry.contains("aspose")) {
					otherCountry = "NA";
				}
			} catch (Exception e1) {
				otherCountry = "NA";
			}
			try {

				otherCountryCode = String.valueOf(contactotherphys.getCountryCode());
			} catch (Exception ep) {
				otherCountryCode = "";
			}
			try {
				if (otherCountryCode.equalsIgnoreCase("null") || otherCountryCode.contains("meta")
						|| otherCountryCode.contains("aspose")) {
					otherCountryCode = "NA";
				}
			} catch (Exception e1) {
				otherCountryCode = "NA";
			}
			try {

				otherPostalCode = String.valueOf(contactotherphys.getPostalCode());
			} catch (Exception ep) {
				otherPostalCode = "";
			}
			try {
				if (otherPostalCode.equalsIgnoreCase("null") || otherPostalCode.contains("meta")
						|| otherPostalCode.contains("aspose")) {
					otherPostalCode = "NA";
				}
			} catch (Exception e1) {
				otherPostalCode = "NA";
			}
			try {

				othergetPostOfficeBox = String.valueOf(contactotherphys.getPostOfficeBox());
			} catch (Exception ep) {
				othergetPostOfficeBox = "";
			}
			try {
				if (othergetPostOfficeBox.equalsIgnoreCase("null") || othergetPostOfficeBox.contains("meta")
						|| othergetPostOfficeBox.contains("aspose")) {
					othergetPostOfficeBox = "NA";
				}
			} catch (Exception e1) {
				othergetPostOfficeBox = "NA";
			}
			try {

				otherStateOrProvince = String.valueOf(contactotherphys.getStateOrProvince());
			} catch (Exception ep) {
				otherStateOrProvince = "";
			}
			try {
				if (otherStateOrProvince.equalsIgnoreCase("null") || otherStateOrProvince.contains("meta")
						|| otherStateOrProvince.contains("aspose")) {
					otherStateOrProvince = "NA";
				}
			} catch (Exception e1) {
				otherStateOrProvince = "NA";
			}
			try {

				workgetStreet = String.valueOf(contactworkphys.getStreet());
			} catch (Exception ep) {
				workgetStreet = "";
			}
			try {
				if (workgetStreet.equalsIgnoreCase("null") || workgetStreet.contains("meta")
						|| workgetStreet.contains("aspose")) {
					workgetStreet = "NA";
				}
			} catch (Exception e1) {
				workgetStreet = "NA";
			}
			try {

				workAddress = String.valueOf(contactworkphys.getAddress());
			} catch (Exception ep) {
				workAddress = "";
			}
			try {
				if (workAddress.equalsIgnoreCase("null") || workAddress.contains("meta")
						|| workAddress.contains("aspose")) {
					workAddress = "NA";
				}
			} catch (Exception e1) {
				workAddress = "NA";
			}
			try {

				workCity = String.valueOf(contactworkphys.getCity());
			} catch (Exception ep) {
				workCity = "";
			}
			try {
				if (workCity.equalsIgnoreCase("null") || workCity.contains("meta") || workCity.contains("aspose")) {
					workCity = "NA";
				}
			} catch (Exception e1) {
				workCity = "NA";
			}

			try {

				workCountry = String.valueOf(contactworkphys.getCountry());
			} catch (Exception ep) {
				workCountry = "";
			}
			try {
				if (workCountry.equalsIgnoreCase("null") || workCountry.contains("meta")
						|| workCountry.contains("aspose")) {
					workCountry = "NA";
				}
			} catch (Exception e1) {
				workCountry = "NA";
			}
			try {

				workCountryCode = String.valueOf(contactworkphys.getCountryCode());
			} catch (Exception ep) {
				workCountryCode = "";
			}
			try {
				if (workCountryCode.equalsIgnoreCase("null") || workCountryCode.contains("meta")
						|| workCountryCode.contains("aspose")) {
					workCountryCode = "NA";
				}
			} catch (Exception e1) {
				workCountryCode = "NA";
			}
			try {

				workPostalCode = String.valueOf(contactworkphys.getPostalCode());
			} catch (Exception ep) {
				workPostalCode = "";
			}
			try {
				if (workPostalCode.equalsIgnoreCase("null") || workPostalCode.contains("meta")
						|| workPostalCode.contains("aspose")) {
					workPostalCode = "NA";
				}
			} catch (Exception e1) {
				workPostalCode = "NA";
			}
			try {

				workgetPostOfficeBox = String.valueOf(contactworkphys.getPostOfficeBox());
			} catch (Exception ep) {
				workgetPostOfficeBox = "";
			}
			try {
				if (workgetPostOfficeBox.equalsIgnoreCase("null") || workgetPostOfficeBox.contains("meta")
						|| workgetPostOfficeBox.contains("aspose")) {
					workgetPostOfficeBox = "NA";
				}
			} catch (Exception e1) {
				workgetPostOfficeBox = "NA";
			}
			try {

				workStateOrProvince = String.valueOf(contactworkphys.getStateOrProvince());
			} catch (Exception ep) {
				workStateOrProvince = "";
			}
			try {
				if (workStateOrProvince.equalsIgnoreCase("null") || workStateOrProvince.contains("meta")
						|| workStateOrProvince.contains("aspose")) {
					workStateOrProvince = "NA";
				}
			} catch (Exception e1) {
				workStateOrProvince = "NA";
			}
			try {

				Assistant = String.valueOf(personPropSet.getAssistant());
			} catch (Exception ep) {
				Assistant = "";
			}
			try {
				if (Assistant.equalsIgnoreCase("null") || Assistant.contains("meta") || Assistant.contains("aspose")) {
					Assistant = "NA";
				}
			} catch (Exception e1) {
				Assistant = "NA";
			}
			try {

				CompanyName = String.valueOf(personPropSet.getCompanyName());
			} catch (Exception ep) {
				CompanyName = "";
			}
			try {
				if (CompanyName.equalsIgnoreCase("null") || CompanyName.contains("meta")
						|| CompanyName.contains("aspose")) {
					CompanyName = "NA";
				}
			} catch (Exception e1) {
				CompanyName = "NA";
			}
			try {

				DepartmentName = String.valueOf(personPropSet.getDepartmentName());
			} catch (Exception ep) {
				DepartmentName = "";
			}
			try {
				if (DepartmentName.equalsIgnoreCase("null") || DepartmentName.contains("meta")
						|| DepartmentName.contains("aspose")) {
					DepartmentName = "NA";
				}
			} catch (Exception e1) {
				DepartmentName = "NA";
			}

			try {

				ManagerName = String.valueOf(personPropSet.getManagerName());
			} catch (Exception ep) {
				ManagerName = "";
			}
			try {
				if (ManagerName.equalsIgnoreCase("null") || ManagerName.contains("meta")
						|| ManagerName.contains("aspose")) {
					ManagerName = "NA";
				}
			} catch (Exception e1) {
				ManagerName = "NA";
			}
			try {

				OfficeLocation = String.valueOf(personPropSet.getOfficeLocation());
			} catch (Exception ep) {
				OfficeLocation = "";
			}
			try {
				if (OfficeLocation.equalsIgnoreCase("null") || OfficeLocation.contains("meta")
						|| OfficeLocation.contains("aspose")) {
					OfficeLocation = "NA";
				}
			} catch (Exception e1) {
				OfficeLocation = "NA";
			}
			try {

				Profession = String.valueOf(personPropSet.getProfession());
			} catch (Exception ep) {
				Profession = "";
			}
			try {
				if (Profession.equalsIgnoreCase("null") || Profession.contains("meta")
						|| Profession.contains("aspose")) {
					Profession = "NA";
				}
			} catch (Exception e1) {
				Profession = "";
			}
			try {

				getTitle = String.valueOf(personPropSet.getTitle());
			} catch (Exception ep) {
				getTitle = "";
			}
			try {
				if (getTitle.equalsIgnoreCase("null") || getTitle.contains("meta") || getTitle.contains("aspose")) {
					getTitle = "NA";
				}
			} catch (Exception e1) {
				getTitle = "NA";
			}
			try {

				AssistantTelephoneNumber = String.valueOf(mapitelephone.getAssistantTelephoneNumber());
			} catch (Exception ep) {
				AssistantTelephoneNumber = "";
			}
			try {
				if (AssistantTelephoneNumber.equalsIgnoreCase("null") || homeCountry.contains("meta")
						|| AssistantTelephoneNumber.contains("meta") || AssistantTelephoneNumber.contains("aspose")) {
					AssistantTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				AssistantTelephoneNumber = "NA";
			}
			try {

				AssistantTelephoneNumber = String.valueOf(mapitelephone.getAssistantTelephoneNumber());
			} catch (Exception ep) {
				AssistantTelephoneNumber = "";
			}
			try {
				if (AssistantTelephoneNumber.equalsIgnoreCase("null") || AssistantTelephoneNumber.contains("meta")
						|| AssistantTelephoneNumber.contains("aspose")) {
					AssistantTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				AssistantTelephoneNumber = "NA";
			}
			try {

				Business2TelephoneNumber = String.valueOf(mapitelephone.getBusiness2TelephoneNumber());
			} catch (Exception ep) {
				Business2TelephoneNumber = "";
			}
			try {
				if (Business2TelephoneNumber.equalsIgnoreCase("null") || Business2TelephoneNumber.contains("meta")
						|| Business2TelephoneNumber.contains("aspose")) {
					Business2TelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				Business2TelephoneNumber = "NA";
			}
			try {

				BusinessTelephoneNumber = String.valueOf(mapitelephone.getBusinessTelephoneNumber());
			} catch (Exception ep) {
				BusinessTelephoneNumber = "";
			}
			try {
				if (BusinessTelephoneNumber.equalsIgnoreCase("null") || BusinessTelephoneNumber.contains("meta")
						|| BusinessTelephoneNumber.contains("aspose")) {
					BusinessTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				BusinessTelephoneNumber = "NA";
			}
			try {

				CallbackTelephoneNumber = String.valueOf(mapitelephone.getCallbackTelephoneNumber());
			} catch (Exception ep) {
				CallbackTelephoneNumber = "";
			}
			try {
				if (CallbackTelephoneNumber.equalsIgnoreCase("null") || CallbackTelephoneNumber.contains("meta")
						|| CallbackTelephoneNumber.contains("aspose")) {
					CallbackTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				CallbackTelephoneNumber = "NA";
			}
			try {

				CarTelephoneNumber = String.valueOf(mapitelephone.getCarTelephoneNumber());
			} catch (Exception ep) {
				CarTelephoneNumber = "";
			}
			try {
				if (CarTelephoneNumber.equalsIgnoreCase("null") || CarTelephoneNumber.contains("meta")
						|| CarTelephoneNumber.contains("aspose")) {
					CarTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				CarTelephoneNumber = "NA";
			}
			try {

				CompanyMainTelephoneNumber = String.valueOf(mapitelephone.getCompanyMainTelephoneNumber());
			} catch (Exception ep) {
				CompanyMainTelephoneNumber = "";
			}
			try {
				if (CompanyMainTelephoneNumber.equalsIgnoreCase("null") || CompanyMainTelephoneNumber.contains("meta")
						|| CompanyMainTelephoneNumber.contains("aspose")) {
					CompanyMainTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				CompanyMainTelephoneNumber = "NA";
			}
			try {

				Home2TelephoneNumber = String.valueOf(mapitelephone.getHome2TelephoneNumber());
			} catch (Exception ep) {
				Home2TelephoneNumber = "";
			}
			try {
				if (Home2TelephoneNumber.equalsIgnoreCase("null") || homeCountry.contains("meta")
						|| Home2TelephoneNumber.contains("meta") || Home2TelephoneNumber.contains("aspose")) {
					Home2TelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				Home2TelephoneNumber = "NA";
			}

			try {

				HomeTelephoneNumber = String.valueOf(mapitelephone.getHomeTelephoneNumber());
			} catch (Exception ep) {
				HomeTelephoneNumber = "";
			}
			try {
				if (HomeTelephoneNumber.equalsIgnoreCase("null") || HomeTelephoneNumber.contains("meta")
						|| HomeTelephoneNumber.contains("aspose")) {
					HomeTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				HomeTelephoneNumber = "NA";
			}
			try {

				IsdnNumber = String.valueOf(mapitelephone.getIsdnNumber());
			} catch (Exception ep) {
				IsdnNumber = "";
			}
			try {
				if (IsdnNumber.equalsIgnoreCase("null") || IsdnNumber.contains("meta")
						|| IsdnNumber.contains("aspose")) {
					IsdnNumber = "NA";
				}
			} catch (Exception e1) {
				IsdnNumber = "NA";
			}

			try {

				MobileTelephoneNumber = String.valueOf(mapitelephone.getMobileTelephoneNumber());
			} catch (Exception ep) {
				MobileTelephoneNumber = "";
			}
			try {
				if (MobileTelephoneNumber.equalsIgnoreCase("null") || MobileTelephoneNumber.contains("meta")
						|| MobileTelephoneNumber.contains("aspose")) {
					MobileTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				MobileTelephoneNumber = "NA";
			}
			try {

				OtherTelephoneNumber = String.valueOf(mapitelephone.getOtherTelephoneNumber());
			} catch (Exception ep) {
				OtherTelephoneNumber = "";
			}
			try {
				if (OtherTelephoneNumber.equalsIgnoreCase("null") || OtherTelephoneNumber.contains("meta")
						|| OtherTelephoneNumber.contains("aspose")) {
					OtherTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				OtherTelephoneNumber = "NA";
			}
			try {

				PagerTelephoneNumber = String.valueOf(mapitelephone.getPagerTelephoneNumber());
			} catch (Exception ep) {
				PagerTelephoneNumber = "";
			}
			try {
				if (PagerTelephoneNumber.equalsIgnoreCase("null") || PagerTelephoneNumber.contains("meta")
						|| PagerTelephoneNumber.contains("aspose")) {
					PagerTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				PagerTelephoneNumber = "NA";
			}
			try {

				PrimaryTelephoneNumber = String.valueOf(mapitelephone.getPrimaryTelephoneNumber());
			} catch (Exception ep) {
				PrimaryTelephoneNumber = "";
			}
			try {
				if (PrimaryTelephoneNumber.equalsIgnoreCase("null") || PrimaryTelephoneNumber.contains("meta")
						|| PrimaryTelephoneNumber.contains("aspose")) {
					PrimaryTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				PrimaryTelephoneNumber = "NA";
			}
			try {

				RadioTelephoneNumber = String.valueOf(mapitelephone.getRadioTelephoneNumber());
			} catch (Exception ep) {
				RadioTelephoneNumber = "";
			}
			try {
				if (RadioTelephoneNumber.equalsIgnoreCase("null") || homeCountry.contains("meta")
						|| RadioTelephoneNumber.contains("meta") || RadioTelephoneNumber.contains("aspose")) {
					RadioTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				RadioTelephoneNumber = "NA";
			}
			try {

				TelexNumber = String.valueOf(mapitelephone.getTelexNumber());
			} catch (Exception ep) {
				TelexNumber = "";
			}
			try {
				if (TelexNumber.equalsIgnoreCase("null") || TelexNumber.contains("meta")
						|| TelexNumber.contains("aspose")) {
					TelexNumber = "NA";
				}
			} catch (Exception e1) {
				TelexNumber = "NA";
			}
			try {

				TtyTddPhoneNumber = String.valueOf(mapitelephone.getTtyTddPhoneNumber());
			} catch (Exception ep) {
				TtyTddPhoneNumber = "";
			}
			try {
				if (TtyTddPhoneNumber.equalsIgnoreCase("null") || TtyTddPhoneNumber.contains("meta")
						|| TtyTddPhoneNumber.contains("aspose")) {
					TtyTddPhoneNumber = "NA";
				}
			} catch (Exception e1) {
				TtyTddPhoneNumber = "NA";
			}

			try {
				String[] data1 = { firstName, middleName, lastname, prefix, Email1addresstype, Email1displayname,
						Email1address, Email1fax, Email2addresstype, Email2displayname, Email2fax, Email2address,
						Email3address, Email3addresstype, Email3displayname, Email3fax, homefaxaddress,
						homefaxaddresstype, homefaxdisplayname, homefaxno, primaryfaxaddress, primaryfaxaddresstype,
						primaryfaxdisplayname, primaryfaxno, bussinessfaxaddress, bussinessfaxaddresstype,
						bussinessfaxdisplayname, bussinessfaxno, WeddingAnniversary, birthday, fileunder, fileunderid,
						generation, title, account, BusinessHomePage, ComputerNetworkName, CustomerId, FreeBusyLocation,
						FtpSite, Gender, GovernmentIdNumber, Hobbies, Html, InstantMessagingAddress, Language, Location,
						Notes, OrganizationalIdNumber, PersonalHomePage, ReferredByName, SpouseName, homeAddress,
						homegetStreet, homeCity, homeCountry, homeCountryCode, homePostalCode, homegetPostOfficeBox,
						homeStateOrProvince, otherAddress, othergetStreet, otherCity, otherCountry, otherCountryCode,
						otherPostalCode, othergetPostOfficeBox, otherStateOrProvince, workAddress, workgetStreet,
						workCity, workCountry, workCountryCode, workPostalCode, workgetPostOfficeBox,
						workStateOrProvince, Assistant, CompanyName, DepartmentName, ManagerName, OfficeLocation,
						Profession, getTitle, AssistantTelephoneNumber, Business2TelephoneNumber,
						BusinessTelephoneNumber, CallbackTelephoneNumber, CarTelephoneNumber,
						CompanyMainTelephoneNumber, Home2TelephoneNumber, HomeTelephoneNumber, IsdnNumber,
						MobileTelephoneNumber, OtherTelephoneNumber, PagerTelephoneNumber, PrimaryTelephoneNumber,
						RadioTelephoneNumber, TelexNumber, TtyTddPhoneNumber };

				writer.writeNext(data1);

			} catch (Error e) {
				mf.logger.warning("ERROR : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

			catch (Exception e) {
				mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

		} else if (message.getMessageClass().equals("IPM.Appointment")
				|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

			MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

			String subject = null;
			String startdate = null;
			String enddate = null;
			String alldayevent = null;
			String reminder = null;
			String remindertime = null;
			String requiredattend = null;
			String categories = null;
			String location = "";
			String mileage = "";
			try {
				subject = cal.getSubject();
			} catch (Exception e) {
				subject = "";
			}
			try {
				if (subject.equalsIgnoreCase("null") || subject.contains("meta") || subject.contains("aspose")) {
					subject = "NA";
				}
			} catch (Exception e1) {
				subject = "NA";
			}
			try {
				startdate = cal.getStartDate().toString();
			} catch (Exception e) {
				startdate = "";
			}
			try {
				if (startdate.equalsIgnoreCase("null") || startdate.contains("meta") || startdate.contains("aspose")) {
					startdate = "NA";
				}
			} catch (Exception e1) {
				startdate = "NA";
			}
			try {
				enddate = cal.getEndDate().toString();
			} catch (Exception e) {
				enddate = "";
			}
			try {
				if (enddate.equalsIgnoreCase("null") || enddate.contains("meta") || enddate.contains("aspose")) {
					enddate = "NA";
				}
			} catch (Exception e1) {
				enddate = "NA";
			}
			try {
				alldayevent = String.valueOf(cal.isAllDay());
			} catch (Exception e) {
				alldayevent = "";
			}
			try {
				if (alldayevent.equalsIgnoreCase("null") || alldayevent.contains("meta")
						|| alldayevent.contains("aspose")) {
					alldayevent = "NA";
				}
			} catch (Exception e1) {
				alldayevent = "NA";
			}
			try {
				reminder = String.valueOf(cal.getReminderSet());
			} catch (Exception e) {
				reminder = "";
			}
			try {
				if (reminder.equalsIgnoreCase("null") || reminder.contains("meta") || reminder.contains("aspose")) {
					reminder = "NA";
				}
			} catch (Exception e1) {
				reminder = "NA";
			}

			try {
				remindertime = String.valueOf(cal.getReminderDelta());
			} catch (Exception e) {
				remindertime = "";
			}
			try {
				if (remindertime.equalsIgnoreCase("null") || remindertime.contains("meta")
						|| remindertime.contains("aspose")) {
					remindertime = "NA";
				}
			} catch (Exception e1) {
				remindertime = "NA";
			}
			try {
				requiredattend = String.valueOf(cal.getAttendees());
			} catch (Exception e) {
				requiredattend = "";
			}
			try {
				if (requiredattend.equalsIgnoreCase("null") || requiredattend.contains("meta")
						|| requiredattend.contains("aspose")) {
					requiredattend = "NA";
				}
			} catch (Exception e1) {
				requiredattend = "NA";
			}
			try {
				categories = String.valueOf(cal.getCategories());
			} catch (Exception e) {
				categories = "";
			}
			try {
				if (categories.equalsIgnoreCase("null") || categories.contains("meta")
						|| categories.contains("aspose")) {
					categories = "NA";
				}
			} catch (Exception e1) {
				categories = "NA";
			}
			try {
				location = String.valueOf(cal.getLocation());
			} catch (Exception e) {
				location = "";
			}
			try {
				if (location.equalsIgnoreCase("null") || location.contains("meta") || location.contains("aspose")) {
					location = "NA";
				}
			} catch (Exception e1) {
				location = "NA";
			}
			try {
				mileage = String.valueOf(cal.getMileage());
			} catch (Exception e) {
				mileage = "";
			}
			try {
				if (mileage.equalsIgnoreCase("null") || mileage.contains("meta") || mileage.contains("aspose")) {
					mileage = "NA";
				}
			} catch (Exception e1) {
				mileage = "NA";
			}
			try {
				String[] data1 = { subject, startdate, enddate, alldayevent, reminder, requiredattend, remindertime,
						requiredattend, categories, location, mileage };

				writer.writeNext(data1);

			} catch (Error e) {
				mf.logger.warning("ERROR : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

			catch (Exception e) {
				mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

		} else if (message.getMessageClass().equals("IPM.StickyNote") || message.getMessageClass().equals("IPM.Task")) {

			MapiTask task = (MapiTask) message.toMapiMessageItem();

			String subject = "";
			String startdate = "";
			String getDueDate = "";
			String getPercentComplete = "";
			String getEstimatedEffort = "";
			String getActualEffort = "";
			String getOwner = "";
			String getLastUser = "";
			String getLastDelegate = "";
			String getAttendeessizesize = "";
			String getOriginalDisplayName = "";
			String getDisplayName = "";
			String getEmailAddress = "";
			String getFaxNumber = "";
			String getAddressType = "";
			String comapanies = "";
			String Categories = "";
			String getMileage = "";
			String getBilling = "";
			String getSensitivity = "";
			String getStatus = "";
			String getHistory = "";

			MapiTaskUsers user = null;
			MapiElectronicAddress address = null;
			String[] company = null;
			String[] getCategories = null;
			try {
				user = task.getUsers();
			} catch (Exception e1) {

			}
			try {
				address = user.getAssigner();
			} catch (Exception e1) {

			}

			try {
				company = task.getCompanies();
			} catch (Exception e1) {

			}
			try {
				getCategories = task.getCategories();
			} catch (Exception e1) {

			}

			try {
				subject = task.getSubject();
			} catch (Exception e) {
				subject = "";
			}
			try {
				if (subject.equalsIgnoreCase("null") || subject.contains("meta") || subject.contains("aspose")) {
					subject = "NA";
				}
			} catch (Exception e1) {
				subject = "NA";
			}
			try {
				startdate = task.getStartDate().toString();
			} catch (Exception e) {
				startdate = "";
			}
			try {
				if (startdate.equalsIgnoreCase("null") || startdate.contains("meta") || startdate.contains("aspose")) {
					startdate = "NA";
				}
			} catch (Exception e1) {
				startdate = "NA";
			}
			try {
				getDueDate = task.getDueDate().toString();
			} catch (Exception e) {
				getDueDate = "";
			}
			try {
				if (getDueDate.equalsIgnoreCase("null") || getDueDate.contains("meta")
						|| getDueDate.contains("aspose")) {
					getDueDate = "NA";
				}
			} catch (Exception e1) {
				getDueDate = "NA";
			}
			try {
				getPercentComplete = String.valueOf(task.getPercentComplete());
			} catch (Exception e) {
				getPercentComplete = "";
			}
			try {
				if (getPercentComplete.equalsIgnoreCase("null") || getPercentComplete.contains("meta")
						|| getPercentComplete.contains("aspose")) {
					getPercentComplete = "NA";
				}
			} catch (Exception e1) {
				getPercentComplete = "NA";
			}

			try {
				getEstimatedEffort = String.valueOf(task.getEstimatedEffort());
			} catch (Exception e) {
				getEstimatedEffort = "";
			}
			try {
				if (getEstimatedEffort.equalsIgnoreCase("null") || getEstimatedEffort.contains("meta")
						|| getEstimatedEffort.contains("aspose")) {
					getEstimatedEffort = "NA";
				}
			} catch (Exception e1) {
				getEstimatedEffort = "NA";
			}
			try {
				getActualEffort = String.valueOf(task.getActualEffort());
			} catch (Exception e) {
				getActualEffort = "";
			}
			try {
				if (getActualEffort.equalsIgnoreCase("null") || getActualEffort.contains("meta")
						|| getActualEffort.contains("aspose")) {
					getActualEffort = "NA";
				}
			} catch (Exception e1) {
				getActualEffort = "";
			}
			try {
				getOwner = String.valueOf(user.getOwner());
			} catch (Exception e) {
				getOwner = "";
			}
			try {
				if (getOwner.equalsIgnoreCase("null") || getOwner.contains("meta") || getOwner.contains("aspose")) {
					getOwner = "NA";
				}
			} catch (Exception e1) {
				getOwner = "NA";
			}
			try {
				getLastUser = String.valueOf(user.getLastUser());
			} catch (Exception e) {
				getLastUser = "";
			}
			try {
				if (getLastUser.equalsIgnoreCase("null") || getLastUser.contains("meta")
						|| getLastUser.contains("aspose")) {
					getLastUser = "NA";
				}
			} catch (Exception e1) {
				getLastUser = "NA";
			}
			try {
				getLastDelegate = String.valueOf(user.getLastDelegate());
			} catch (Exception e) {
				getLastDelegate = "";
			}
			try {
				if (getLastDelegate.equalsIgnoreCase("null") || getLastDelegate.contains("meta")
						|| getLastDelegate.contains("aspose")) {
					getLastDelegate = "NA";
				}
			} catch (Exception e1) {
				getLastDelegate = "NA";
			}
			try {
				getAttendeessizesize = String.valueOf(user.getAttendees().size());
			} catch (Exception e) {
				getAttendeessizesize = "";
			}
			try {
				if (getAttendeessizesize.equalsIgnoreCase("null") || getAttendeessizesize.contains("meta")
						|| getAttendeessizesize.contains("aspose")) {
					getAttendeessizesize = "NA";
				}
			} catch (Exception e2) {
				getAttendeessizesize = "NA";
			}
			try {
				getOriginalDisplayName = String.valueOf(address.getOriginalDisplayName());

			} catch (Exception e) {
				getOriginalDisplayName = "";
			}
			try {
				if (getOriginalDisplayName.equalsIgnoreCase("null") || getOriginalDisplayName.contains("meta")
						|| getOriginalDisplayName.contains("aspose")) {
					getOriginalDisplayName = "NA";
				}
			} catch (Exception e1) {
				getOriginalDisplayName = "NA";
			}
			try {
				getDisplayName = String.valueOf(address.getDisplayName());

			} catch (Exception e) {
				getDisplayName = "";
			}
			if (getDisplayName.equalsIgnoreCase("null") || getDisplayName.contains("meta")
					|| getDisplayName.contains("aspose")) {
				getDisplayName = "NA";
			}
			try {
				getEmailAddress = String.valueOf(address.getEmailAddress());

			} catch (Exception e) {
				getEmailAddress = "";
			}
			try {
				if (getEmailAddress.equalsIgnoreCase("null") || getEmailAddress.contains("meta")
						|| getEmailAddress.contains("aspose")) {
					getEmailAddress = "NA";
				}
			} catch (Exception e1) {
				getEmailAddress = "NA";
			}
			try {
				getFaxNumber = String.valueOf(address.getFaxNumber());

			} catch (Exception e) {
				getFaxNumber = "";
			}
			try {
				if (getFaxNumber.equalsIgnoreCase("null") || getFaxNumber.contains("meta")
						|| getFaxNumber.contains("aspose")) {
					getFaxNumber = "NA";
				}
			} catch (Exception e1) {
				getFaxNumber = "NA";
			}
			try {
				getAddressType = String.valueOf(address.getAddressType());

			} catch (Exception e) {
				getAddressType = "";
			}
			try {
				if (getAddressType.equalsIgnoreCase("null") || getAddressType.contains("meta")
						|| getAddressType.contains("aspose")) {
					getAddressType = "NA";
				}
			} catch (Exception e1) {
				getAddressType = "NA";
			}
			try {
				for (int i = 0; i < company.length; i++) {
					if (i == 0) {
						comapanies = company[i];
					} else {
						comapanies = comapanies + "," + company[i];
					}
				}

			} catch (Exception e) {
				comapanies = "";
			}
			if (comapanies.equalsIgnoreCase("null") || comapanies.contains("meta") || comapanies.contains("aspose")) {
				comapanies = "NA";
			}

			try {
				for (int i = 0; i < getCategories.length; i++) {
					if (i == 0) {
						Categories = getCategories[i];
					} else {
						Categories = Categories + "," + getCategories[i];
					}
				}

			} catch (Exception e) {
				Categories = "";
			}
			try {
				if (Categories.equalsIgnoreCase("null") || Categories.contains("meta")
						|| Categories.contains("aspose")) {
					Categories = "NA";
				}
			} catch (Exception e1) {
				Categories = "NA";
			}
			try {
				getMileage = String.valueOf(task.getMileage());
			} catch (Exception e) {
				getMileage = "";
			}
			try {
				if (getMileage.equalsIgnoreCase("null") || getMileage.contains("meta")
						|| getMileage.contains("aspose")) {
					getMileage = "NA";
				}
			} catch (Exception e1) {
				getMileage = "NA";
			}
			try {
				getBilling = String.valueOf(task.getBilling());
			} catch (Exception e) {
				getBilling = "";
			}
			try {
				if (getBilling.equalsIgnoreCase("null") || getBilling.contains("meta")
						|| getBilling.contains("aspose")) {
					getBilling = "NA";
				}
			} catch (Exception e1) {
				getBilling = "NA";
			}
			try {
				int i = task.getSensitivity();
				if (i == 0) {
					getSensitivity = "None";
				} else if (i == 1) {
					getSensitivity = "Personal";

				} else if (i == 2) {
					getSensitivity = "Private";
				} else if (i == 3) {
					getSensitivity = "Company Confidential";
				}

			} catch (Exception e) {
				getSensitivity = "";
			}
			try {
				if (getSensitivity.equalsIgnoreCase("null") || getSensitivity.contains("meta")
						|| getSensitivity.contains("aspose")) {
					getSensitivity = "NA";
				}
			} catch (Exception e2) {
				getSensitivity = "NA";
			}
			try {
				int i = task.getStatus();
				if (i == 0) {
					getStatus = "Not Started";
				} else if (i == 1) {
					getStatus = "In Progress";

				} else if (i == 2) {
					getStatus = "Complete";
				} else if (i == 3) {
					getStatus = "Waiting";
				} else if (i == 4) {
					getStatus = "Deferred";
				}

			} catch (Exception e) {
				getStatus = "";
			}
			try {
				if (getStatus.equalsIgnoreCase("null") || getStatus.contains("meta") || getStatus.contains("aspose")) {
					getStatus = "NA";
				}
			} catch (Exception e1) {
				getStatus = "NA";
			}
			try {
				int i = task.getHistory();
				if (i == 0) {
					getHistory = "No Changes";
				} else if (i == 1) {
					getHistory = "Accepted";

				} else if (i == 2) {
					getHistory = "Rejected";
				} else if (i == 3) {
					getHistory = "Another Property Changed";
				} else if (i == 4) {
					getHistory = "Due Date Changed";
				} else if (i == 5) {
					getHistory = "Assigned";
				}

			} catch (Exception e) {
				getStatus = "";
			}
			try {
				if (getHistory.equalsIgnoreCase("null") || getHistory.contains("meta")
						|| getHistory.contains("aspose")) {
					getHistory = "NA";
				}
			} catch (Exception e) {
				getHistory = "NA";
			}

			String[] data1 = { subject, startdate, getDueDate, getPercentComplete, getEstimatedEffort, getActualEffort,
					getOwner, getLastUser, getLastDelegate, getAttendeessizesize, getOriginalDisplayName,
					getDisplayName, getEmailAddress, getFaxNumber, getAddressType, comapanies, Categories, getMileage,
					getBilling, getSensitivity, getStatus, getHistory };

			writer.writeNext(data1);
		} else {

			try {
				String date = null;
				try {
					date = message.getDeliveryTime().toString();
				} catch (Exception e) {
					date = "NA";
				}

				try {
					if (date.equalsIgnoreCase("null") || date.contains("meta") || date.contains("aspose")) {
						date = "NA";
					}
				} catch (Exception e1) {
					date = "NA";
				}
				String subject = null;
				try {
					subject = message.getSubject();
				} catch (Exception e) {

					subject = message.getBodyHtml();
				}

				try {
					if (subject.equalsIgnoreCase("null") || subject.contains("meta") || subject.contains("aspose")) {
						subject = "NA";
					}
				} catch (Exception e1) {
					subject = "NA";
				}
				String getBody = null;
				try {
					getBody = message.getBody();
				} catch (Exception e) {
					getBody = message.getBodyHtml();
//					getBody = "NA";
				}

				if (getBody.equalsIgnoreCase("")) {
					getBody = message.getBodyHtml();

				}

				try {
					if (getBody.equalsIgnoreCase("null") || getBody.contains("meta") || getBody.contains("aspose")) {
						getBody = "NA";
					}

				} catch (Exception e1) {
					getBody = "NA";
				}

				String getSenderEmailAddress = null;
				try {
					getSenderEmailAddress = message.getSenderEmailAddress();
				} catch (Exception e) {
					getSenderEmailAddress = "NA";
				}

				try {
					if (getSenderEmailAddress.equalsIgnoreCase("null") || getSenderEmailAddress.contains("meta")
							|| getSenderEmailAddress.contains("aspose")) {
						getSenderEmailAddress = "NA";
					}
				} catch (Exception e1) {
					getSenderEmailAddress = "NA";
				}
				String getReplyTo = null;
				try {

					for (int i = 0; i < message.getRecipients().size(); i++) {
						String toid = null;
						try {
							toid = message.getRecipients().get_Item(i).getEmailAddress();
						} catch (Exception e) {

						}
						if (i == 0) {
							getReplyTo = toid;
						} else {
							getReplyTo = getReplyTo + "," + toid;

						}

					}

				} catch (Exception e) {
					getReplyTo = "NA";
				}

				try {
					if (getReplyTo.equalsIgnoreCase("null") || getReplyTo.contains("meta")
							|| getReplyTo.contains("aspose")) {
						getReplyTo = "NA";
					}
				} catch (Exception e1) {
					getReplyTo = "NA";
				}

				String getDisplayCc = null;
				try {
					getDisplayCc = message.getDisplayCc();
				} catch (Exception e) {
					getDisplayCc = "NA";
				}

				try {
					if (getDisplayCc.equalsIgnoreCase("null") || getDisplayCc.contains("meta")
							|| getDisplayCc.contains("aspose")) {
						getDisplayCc = "NA";
					}
				} catch (Exception e1) {
					getDisplayCc = "NA";
				}

				String getDisplayBcc = null;
				try {
					getDisplayBcc = message.getDisplayBcc();
				} catch (Exception e) {
					getDisplayBcc = "NA";
				}

				try {
					if (getDisplayBcc.equalsIgnoreCase("null") || getDisplayBcc.contains("meta")
							|| getDisplayBcc.contains("aspose")) {
						getDisplayBcc = "NA";
					}
				} catch (Exception e1) {
					getDisplayBcc = "NA";
				}

				if (message.getAttachments().size() > 0) {
					File fd = new File(destination_path + File.separator + path + File.separator + "Attachment"
							+ File.separator + subname);

					fd.mkdirs();

					String[] data1 = { date, subject, getBody, getSenderEmailAddress, getReplyTo, getDisplayCc,
							getDisplayBcc, fd.getAbsolutePath() };

					writer.writeNext(data1);

				} else {
					String[] data1 = { date, subject, getBody, getSenderEmailAddress, getReplyTo, getDisplayCc,
							getDisplayBcc };

					writer.writeNext(data1);
				}

			} catch (Error e) {
				mf.logger.warning("ERROR : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

			catch (Exception e) {
				mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

		}

		try {
			count_destination++;

			for (int j = 0; j < message.getAttachments().size(); j++) {

				MapiAttachment att = message.getAttachments().get_Item(j);

				String s = getFileExtension(att.getLongFileName());
				String attFileName = getRidOfIllegalFileNameCharacters(att.getLongFileName().replace("." + s, ""));

				att.save(destination_path + File.separator + path + File.separator + "Attachment" + File.separator
						+ subname + File.separator + attFileName + "." + s);

			}
		} catch (Exception e) {

		}

	}

	public void Mailmess_CSV(MailMessage message, CSVWriter writer) {

		String subname = getRidOfIllegalFileNameCharacters(message.getSubject() + message.getDate());

		MapiMessage mp = MapiMessage.fromMailMessage(message);

		try {
			String date = null;
			try {
				date = mp.getDeliveryTime().toString();
			} catch (Exception e) {
				date = "NA";
			}

			try {
				if (date.equalsIgnoreCase("null") || date.contains("meta") || date.contains("aspose")) {
					date = "NA";
				}
			} catch (Exception e1) {
				date = "NA";
			}
			String subject = null;
			try {
				subject = message.getSubject();
			} catch (Exception e) {

				subject = mp.getBodyHtml();
			}

			try {
				if (subject.equalsIgnoreCase("null") || subject.contains("meta") || subject.contains("aspose")) {
					subject = "NA";
				}
			} catch (Exception e1) {
				subject = "NA";
			}
			String getBody = null;
			try {
				getBody = message.getBody();
			} catch (Exception e) {
				getBody = mp.getBodyHtml();
//				getBody = "NA";
			}

			if (getBody.equalsIgnoreCase("")) {
				getBody = mp.getBodyHtml();

			}

			try {
				if (getBody.equalsIgnoreCase("null") || getBody.contains("meta") || getBody.contains("aspose")) {
					getBody = "NA";
				}

			} catch (Exception e1) {
				getBody = "NA";
			}

			String getSenderEmailAddress = null;
			try {
				getSenderEmailAddress = mp.getSenderEmailAddress();
			} catch (Exception e) {
				getSenderEmailAddress = "NA";
			}

			try {
				if (getSenderEmailAddress.equalsIgnoreCase("null") || getSenderEmailAddress.contains("meta")
						|| getSenderEmailAddress.contains("aspose")) {
					getSenderEmailAddress = "NA";
				}
			} catch (Exception e1) {
				getSenderEmailAddress = "NA";
			}
			String getReplyTo = null;
			try {

				for (int i = 0; i < mp.getRecipients().size(); i++) {
					String toid = null;
					try {
						toid = mp.getRecipients().get_Item(i).getEmailAddress();
					} catch (Exception e) {

					}
					if (i == 0) {
						getReplyTo = toid;
					} else {
						getReplyTo = getReplyTo + "," + toid;

					}

				}

			} catch (Exception e) {
				getReplyTo = "NA";
			}

			try {
				if (getReplyTo.equalsIgnoreCase("null") || getReplyTo.contains("meta")
						|| getReplyTo.contains("aspose")) {
					getReplyTo = "NA";
				}
			} catch (Exception e1) {
				getReplyTo = "NA";
			}

			String getDisplayCc = null;
			try {
				getDisplayCc = mp.getDisplayCc();
			} catch (Exception e) {
				getDisplayCc = "NA";
			}

			try {
				if (getDisplayCc.equalsIgnoreCase("null") || getDisplayCc.contains("meta")
						|| getDisplayCc.contains("aspose")) {
					getDisplayCc = "NA";
				}
			} catch (Exception e1) {
				getDisplayCc = "NA";
			}

			String getDisplayBcc = null;
			try {
				getDisplayBcc = mp.getDisplayBcc();
			} catch (Exception e) {
				getDisplayBcc = "NA";
			}

			try {
				if (getDisplayBcc.equalsIgnoreCase("null") || getDisplayBcc.contains("meta")
						|| getDisplayBcc.contains("aspose")) {
					getDisplayBcc = "NA";
				}
			} catch (Exception e1) {
				getDisplayBcc = "NA";
			}

			if (message.getAttachments().size() > 0) {
				File fd = new File(destination_path + File.separator + path + File.separator + "Attachment"
						+ File.separator + subname);

				fd.mkdirs();

				String[] data1 = { date, subject, getBody, getSenderEmailAddress, getReplyTo, getDisplayCc,
						getDisplayBcc, fd.getAbsolutePath() };

				writer.writeNext(data1);

			} else {
				String[] data1 = { date, subject, getBody, getSenderEmailAddress, getReplyTo, getDisplayCc,
						getDisplayBcc };

				writer.writeNext(data1);
			}
			if (message.getAttachments().size() > 0) {
				new File(destination_path + File.separator + path + File.separator + "Attachment" + File.separator
						+ subname).mkdirs();

			}

			for (int j = 0; j < message.getAttachments().size(); j++) {
				Attachment att = (Attachment) message.getAttachments().get_Item(j);

				String s = getFileExtension(att.getName());
				String attFileName = getRidOfIllegalFileNameCharacters(att.getName().replace("." + s, ""));

				att.save(destination_path + File.separator + path + File.separator + "Attachment" + File.separator
						+ subname + attFileName + "." + s);
			}

			count_destination++;

		} catch (Error e) {
			mf.logger.warning(
					"ERROR : " + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
		}

		catch (Exception e) {
			mf.logger.warning(
					"Exception : " + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
			count_destination++;
			return;
		}

	}

	public void connectionHandle(String gotMessage) {
		lbl_progressreport.setText("Internet Connection  LOST ");
		System.out.println("internet not connected ");
		label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));

		while (true) {
			try {
				lbl_progressreport.setText("Connecting to Server Please Wait");
				System.out.println("please check connection");
				if (filetype.equalsIgnoreCase("OFFICE 365")) {
					conntiontooffice365_output();

				} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {

					connectiontoinaws_output();

				} else if (filetype.equalsIgnoreCase("Yandex Mail")) {
					connectiontoYandex_output();

				} else if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")) {
					if (modern_Authentication.isSelected()) {
						String token = GetToken.refreshToken_Gmail_Output();
						if (token != null) {
							System.out.println("this is modern auth connection handle calling ");
							clientforimap_output.dispose();
							clientforimap_output = GetToken.loginGmail_output(token);

						}
					} else {
						connectiontogmail_output();
					}
					System.out.println("Connection Done !!");
				} else if (filetype.equalsIgnoreCase("Hostgator email")) {
					connectiontoHostgator_output();

				} else if (filetype.equalsIgnoreCase("Icloud")) {
					connectiontoicloud_output();

				} else if (filetype.equalsIgnoreCase("GoDaddy email")) {
					connectiontoGoDaddy_output();

				} else if (filetype.equalsIgnoreCase("Live Exchange")) {
					connectionwithexchangeserver_output();

				} else if (filetype.equalsIgnoreCase("IMAP")) {

					connectiontoimap_output();

				} else if (filetype.equalsIgnoreCase("Hotmail")) {
					conntiontohotmail_output();

				} else if (filetype.equalsIgnoreCase("Zoho MAIL")) {

					connectiontozoho_output();

				} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

					connectiontoyahoo_output();

				} else if (filetype.equalsIgnoreCase("AOL")) {

					connectiontoaol_output();

				}
				System.out.println("this is after creating connection   10487");

				label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));
				lbl_progressreport.setText("Connection established Retrieving Message");
				System.out.println("this is after creating connection  10491 ");

				break;
			} catch (Exception e) {
				lbl_progressreport.setText("INTERNET Connection  LOST ");

			}

		}

		Progressbar.setVisible(true);

	}

	public void connectionHandle() {
		lbl_progressreport.setText("Internet Connection  LOST ");
		System.out.println("internet not conn");
		label_12.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));

		while (true) {
			System.out.println("Connection not established ");
			try {
				lbl_progressreport.setText("Connecting to Server Please Wait...");
				System.out.println("please check conncetion");
				if (filetype.equalsIgnoreCase("OFFICE 365")) {
					conntiontooffice365_output();
				}

				else if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")) {

					connectiontogmail_output();

				} else if (filetype.equalsIgnoreCase("Yandex Mail")) {
					connectiontoYandex_output();

				} else if (filetype.equalsIgnoreCase("Hostgator email")) {
					connectiontoHostgator_output();

				} else if (filetype.equalsIgnoreCase("Icloud")) {
					connectiontoicloud_output();

				} else if (filetype.equalsIgnoreCase("GoDaddy email")) {
					connectiontoGoDaddy_output();

				} else if (filetype.equalsIgnoreCase("Live Exchange")) {
					connectionwithexchangeserver_output();

				} else if (filetype.equalsIgnoreCase("IMAP")) {

					connectiontoimap_output();

				} else if (filetype.equalsIgnoreCase("Hotmail")) {
					conntiontohotmail_output();

				} else if (filetype.equalsIgnoreCase("Zoho MAIL")) {

					connectiontozoho_output();

				} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

					connectiontoyahoo_output();

				} else if (filetype.equalsIgnoreCase("AOL")) {

					connectiontoaol_output();

				} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {

					connectiontoinaws_output();

				}

				label_12.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));
				lbl_progressreport.setText("Connection established Retrieving Message");
				break;
			} catch (Exception e) {
				e.printStackTrace();
				lbl_progressreport.setText("Internet Connection  LOST ");

			}

		}

		Progressbar.setVisible(true);

	}

	public boolean isValid(String email) {
		String emailRegex = "^[a-zA-Z0-9_+&*-]+(?:\\." + "[a-zA-Z0-9_+&*-]+)*@" + "(?:[a-zA-Z0-9-]+\\.)+[a-z"
				+ "A-Z]{2,7}$";

		Pattern pat = Pattern.compile(emailRegex);
		if (email == null)
			return false;
		return pat.matcher(email).matches();
	}

	public void connectionHandle1() {
		label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));
		while (true) {
			try {
				if (filetype.equalsIgnoreCase("OFFICE 365")) {
					conntiontooffice365_output();
				} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {

					connectiontoinaws_output();

				} else if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")) {

					connectiontogmail_output();

				} else if (filetype.equalsIgnoreCase("Yandex Mail")) {
					connectiontoYandex_output();

				} else if (filetype.equalsIgnoreCase("Hostgator email")) {
					connectiontoHostgator_output();

				} else if (filetype.equalsIgnoreCase("Icloud")) {
					connectiontoicloud_output();

				} else if (filetype.equalsIgnoreCase("GoDaddy email")) {
					connectiontoGoDaddy_output();

				} else if (filetype.equalsIgnoreCase("Live Exchange")) {
					connectionwithexchangeserver_output();

				} else if (filetype.equalsIgnoreCase("IMAP")) {

					connectiontoimap_output();

				} else if (filetype.equalsIgnoreCase("Hotmail")) {
					conntiontohotmail_output();

				} else if (filetype.equalsIgnoreCase("Zoho MAIL")) {

					connectiontozoho_output();

				} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

					connectiontoyahoo_output();

				} else if (filetype.equalsIgnoreCase("AOL")) {

					connectiontoaol_output();

				}
				label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));
				break;

			} catch (Exception e) {
				e.printStackTrace();
			}
		}

	}

	public static void visitAllNodes(DefaultMutableTreeNode roe) {

		Enumeration<TreeNode> e = roe.depthFirstEnumeration();
		while (e.hasMoreElements()) {
			DefaultMutableTreeNode node = (DefaultMutableTreeNode) e.nextElement();

			lists.add(node);
			listst.add(node.toString().replace("<html><b>", ""));

		}

	}

	public IEWSClient connectionwithexchangeserver_output() throws Exception {

		clientforexchange_output = EWSClient.getEWSClient("https://" + domain_p3 + "/ews/Exchange.asmx", username_p3,
				password_p3);

		clientforexchange_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_output;

	}

	public IEWSClient conntiontohotmail_output() throws Exception {
		clientforexchange_output = EWSClient.getEWSClient("https://outlook.live.com/EWS/Exchange.asmx", username_p3,
				password_p3);

		clientforexchange_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_output;
	}

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

	public ImapClient connectiontoimap_output() throws Exception {
		clientforimap_output = new ImapClient(domain_p3, portnofiletype, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		clientforimap_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

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

	public ImapClient connectiontoGoDaddy_output() throws Exception {
		clientforimap_output = new ImapClient("imap.secureserver.net", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		// clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoHostgator_output() throws Exception {
		clientforimap_output = new ImapClient(domain_p3, portnofiletype, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		// clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

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

	public ImapClient connectiontogmail_output() throws Exception {
		clientforimap_output = new ImapClient("imap.gmail.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoinaws_output() throws Exception {
		clientforimap_output = new ImapClient(domain_p3, portnofiletype, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		clientforimap_output.setTimeout(5 * 60 * 1000);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoaol_output() throws Exception {
		clientforimap_output = new ImapClient("imap.aol.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoyahoo_output() throws Exception {
		clientforimap_output = new ImapClient("imap.mail.yahoo.com", 993, username_p3, password_p3);
		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);
		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;

	}

	public static IEWSClient conntiontooffice365_output() throws Exception {
//		if (modern_Authentication.isSelected()) {
//		String token = One.Chilkat_Connection();
//		NetworkCredential credentials = new OAuthNetworkCredential(token);
//		EWSClient.useSAAJAPI(true);
//		clientforexchange_output = EWSClient.getEWSClient("https://outlook.office365.com/ews/exchange.asmx",
//				credentials);
//		clientforexchange_output.setTimeout(5 * 60 * 1000);
//		EmailClient.setSocketsLayerVersion2(true);
//		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
//		System.out.println("Connection Done : ");
//	} else {
//		EWSClient.useSAAJAPI(true);
//		clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3, password_p3);
//		EmailClient.setSocketsLayerVersion2(true);
//		clientforexchange_output.setTimeout(5 * 60 * 1000);
//		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
//	}
//	return clientforexchange_output;

//	try {
		EWSClient.useSAAJAPI(true);

		clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3, password_p3);
		EmailClient.setSocketsLayerVersion2(true);

		clientforexchange_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_output;

	}

	public ImapClient conntiontooffice365_outputImap() throws Exception {
		if (modern_Authentication.isSelected()) {
			String token = One.Chilkat_Connection();
			clientforimap_output = new ImapClient("outlook.office365.com", 993, username_p3, token, true);
			clientforimap_output.setSecurityOptions(SecurityOptions.SSLAuto);
			clientforimap_output.setSecurityOptions(SecurityOptions.SSLImplicit);
			clientforimap_output.setTimeout(60 * 1000);
			iconnforimap_output = clientforimap_output.createConnection();
			System.out.println("Connection Done : Office");
		}
		return clientforimap_output;
	}

	public static String getRidOfIllegalFileNameCharacters(String strName) {
		String strLegalName = strName;
		if (strLegalName.length() >= 80) {
			strLegalName = strLegalName.substring(0, 80);
		}
		strLegalName = strName.replace(":", " ").replace("\\", "").replace("?", "").replace("/", "").replace("|", "")
				.replace("*", "").replace("<", "").replace(">", "").replace("\t", "").replace("\"", "").replace(",", "")
				.replace("", "");
		if (strLegalName.length() >= 100) {
			strLegalName = strLegalName.substring(0, 100);
		}
		return strLegalName;
	}

	static String namegetRidOfIllegalFileNameCharacters(String strName) {
		String strLegalName = strName;
		if (strLegalName.length() >= 80) {
			strLegalName = strLegalName.substring(0, 80);
		}
		strLegalName = strName.replace(":", " ").replace("\\", "").replace("?", "").replace("/", "").replace("|", "")
				.replace("*", "").replace("<", "").replace(">", "").replace("\t", "").replace("\"", "").replace(",", "")
				.replace(" ", "");

		if (strLegalName.length() >= 100) {
			strLegalName = strLegalName.substring(0, 100);
		}
		return strLegalName;
	}

	public static void search(TreeNode rootNode, DefaultMutableTreeNode searchNode) {

		for (int i = 0; i < rootNode.getChildCount(); i++) {

			if (rootNode.getChildAt(i).toString().equals(searchNode.toString())) {

				lastNode = (DefaultMutableTreeNode) rootNode.getChildAt(i);
			}

			else {

				search((DefaultMutableTreeNode) rootNode.getChildAt(i), searchNode);
			}

		}

	}

	void openBrowser(String url) {
		if (Desktop.isDesktopSupported()) {
			Desktop desktop = Desktop.getDesktop();
			try {
				desktop.browse(new URI(url));
			} catch (IOException | URISyntaxException e) {
				// mf.logger.warning("Warning : " + e.getMessage());
			}
		} else {
			Runtime runtime = Runtime.getRuntime();
			try {
				runtime.exec("xdg-open " + url);
			} catch (IOException e) {
				// .warning("Warning : " + e.getMessage());
			}
		}
	}

	private static String getFileExtension(File file) {
		String fileName = file.getName();
		if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
			return fileName.substring(fileName.lastIndexOf(".") + 1);
		else
			return "";
	}

	static String getFileExtension(String fileName) {

		if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
			return fileName.substring(fileName.lastIndexOf(".") + 1);
		else
			return "";
	}

	private StringBuilder Duration(long startTime) {
		long elapsedTime = System.currentTimeMillis() - startTime;
		long elapsedSeconds = elapsedTime / 1000;
		long secondsDisplay = elapsedSeconds % 60;
		long elapsedMinutes = elapsedSeconds / 60;
		StringBuilder br = new StringBuilder();
		br.append(elapsedMinutes);
		br.append(':');
		br.append(secondsDisplay);
		return br;
	}

	private String filepath(File file) {
		String fileName = file.getAbsolutePath();
		String filepath = fileName.replace(file.getName(), "");
		return filepath;
	}

	String duplicacymail(MailMessage msg) {
		String frm;
		try {
			frm = getRidOfIllegalFileNameCharacters(msg.getFrom().toString());
		} catch (Exception ep) {
			frm = "";
		}
		if (frm != null) {

		} else {
			frm = "";
		}
		String to;
		try {
			to = msg.getTo().get_Item(0).getAddress();
		} catch (Exception e) {
			to = "";
		}
		if (to != null) {

		} else {
			to = "";
		}
		String body;
		try {
			body = msg.getBody();
		} catch (Exception ep) {
			body = "";
		}

		if (body != null) {
			try {
				body = body.substring(0, 1000);
			} catch (Exception e) {

			}

		} else {
			body = "";
		}

		String bcc;
		try {
			bcc = msg.getBcc().get_Item(0).getAddress();
		} catch (Exception ep) {
			bcc = "";
		}

		if (bcc != null) {

		} else {
			bcc = "";
		}

		String sub;
		try {
			sub = getRidOfIllegalFileNameCharacters(msg.getSubject());
		} catch (Exception ep) {
			sub = "";
		}

		if (sub != null) {

		} else {
			sub = "";
		}

		SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-YYYY");
		String dstr;
		Date d;
		try {
			d = msg.getDate();
			dstr = formatter.format(d);
		} catch (Exception ep) {
			dstr = "";
		}

		String value = frm + dstr + to + body;
		String input = value.replace(" ", "");
//        System.out.println("Input :" + input);
		MessageDigest md = null;
		try {
			md = MessageDigest.getInstance("MD5");
		} catch (NoSuchAlgorithmException e) {

			e.printStackTrace();
		}
		byte[] messageDigest = md.digest(input.trim().getBytes());

		// Convert byte array into signum representation
		BigInteger no = new BigInteger(1, messageDigest);
		String hashtext = no.toString(16);
		while (hashtext.length() < 32) {
			hashtext = "0" + hashtext;
		}
		return hashtext;

	}

	String duplicacymapi(MapiMessage msg) {
		MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
		MailConversionOptions de = new MailConversionOptions();
		MailMessage mess = msg.toMailMessage(de);
		MapiMessage message = MapiMessage.fromMailMessage(mess, d);
		String frm;
		try {
			frm = msg.getSenderEmailAddress();
		} catch (Exception ep) {
			frm = "";
		}
		if (frm != null) {

		} else {
			frm = "";
		}

		String to;
		try {
			to = msg.getDisplayTo();
		} catch (Exception e) {
			to = "";
		}
		if (to != null) {

		} else {
			to = "";
		}
		String sub;
		try {
			sub = msg.getSubject();
		} catch (Exception ep) {
			sub = "";
		}

		if (sub != null) {

		} else {
			sub = "";
		}
		String body;
		try {
			body = message.getBody();
		} catch (Exception ep) {
			body = "";
		}

		if (body != null) {
			try {
				body = body.substring(0, 1000);
			} catch (Exception e) {

			}

		} else {
			body = "";
		}

		String bcc;
		try {
			bcc = msg.getDisplayBcc();
		} catch (Exception ep) {
			bcc = "";
		}

		if (bcc != null) {

		} else {
			bcc = "";
		}

		String value = sub + frm + body + to + bcc;
		String input = value.replace(" ", "");
//		System.out.println("Input : " + input);
		MessageDigest md = null;
		try {
			md = MessageDigest.getInstance("MD5");
		} catch (NoSuchAlgorithmException e) {

			e.printStackTrace();
		}
		byte[] messageDigest = md.digest(input.getBytes());

		BigInteger no = new BigInteger(1, messageDigest);
		String hashtext = no.toString(16);
		while (hashtext.length() < 32) {
			hashtext = "0" + hashtext;
		}
		return hashtext;

	}

	String duplicacymapiCal(MapiCalendar calender) {

		String checkBoxCalRecipient = null;
		String checkBoxCalSub = null;
		String chckbxStartdate = null;
		String chckbxEnddate = null;
		MessageDigest md = null;

		try {
			MapiElectronicAddress mrc = calender.getOrganizer();
			checkBoxCalRecipient = mrc.getEmailAddress();
		} catch (Exception ep) {
			checkBoxCalRecipient = "";
		}

		if (checkBoxCalRecipient != null) {

		} else {
			checkBoxCalRecipient = "";
		}

		try {
			checkBoxCalSub = calender.getSubject();
		} catch (Exception ep) {
			checkBoxCalSub = "";
		}

		if (checkBoxCalSub != null) {

		} else {
			checkBoxCalSub = "";
		}

		try {

			chckbxStartdate = calender.getStartDate().toString();
		} catch (Exception ep) {
			chckbxStartdate = "";
		}

		if (chckbxStartdate != null) {

		} else {
			chckbxStartdate = "";
		}

		try {
			chckbxEnddate = calender.getEndDate().toString();
		} catch (Exception ep) {
			chckbxStartdate = "";
		}
		if (chckbxEnddate != null) {

		} else {
			chckbxEnddate = "";
		}
		String value = checkBoxCalSub + chckbxStartdate + chckbxEnddate + checkBoxCalRecipient;
		String input = value.replace(" ", "");
//		System.out.println("Input : " + input);
		try {
			md = MessageDigest.getInstance("MD5");
		} catch (NoSuchAlgorithmException e) {

			e.printStackTrace();
		}
		byte[] messageDigest = md.digest(input.getBytes());

		BigInteger no = new BigInteger(1, messageDigest);
		String hashtext = no.toString(16);
		while (hashtext.length() < 32) {
			hashtext = "0" + hashtext;
		}
		return hashtext;
	}

	String duplicacymapiTask(MapiTask task) {

		String chckbxtaskenddate = null;
		String checkBoxtaskibilling = null;
		String chckbxtaskCategories_1 = null;
		String chckbxtaskCompanies = null;
		String chckbxtaskStatus = null;
		String checkBoxtasksub = null;
		String checkBox_taskbody = null;
		String checkBox_taskstartdate = null;

		MessageDigest md = null;

		try {

			checkBoxtaskibilling = task.getBilling();
		} catch (Exception ep) {
			checkBoxtaskibilling = "";
		}
		if (checkBoxtaskibilling != null) {
//			System.out.println(checkBoxtaskibilling);
		} else {
			checkBoxtaskibilling = "";
		}

		try {

			int status = task.getStatus();

			chckbxtaskStatus = String.valueOf(status);

		} catch (Exception ep) {
			chckbxtaskStatus = "";
		}

		if (chckbxtaskStatus != null) {

		} else {
			chckbxtaskStatus = "";
		}

		try {
			checkBoxtasksub = task.getSubject();
		} catch (Exception ep) {
			checkBoxtasksub = "";
		}

		if (checkBoxtasksub != null) {

		} else {
			checkBoxtasksub = "";
		}

		try {

			checkBox_taskstartdate = task.getStartDate().toString();
		} catch (Exception ep) {
			checkBox_taskstartdate = "";
		}

		if (checkBox_taskstartdate != null) {

		} else {
			checkBox_taskstartdate = "";
		}

		try {
			chckbxtaskenddate = task.getDueDate().toString();
		} catch (Exception ep) {
			chckbxtaskenddate = "";
		}
		if (chckbxtaskenddate != null) {

		} else {
			chckbxtaskenddate = "";
		}

		try {

			String[] st = task.getCategories();
			chckbxtaskCategories_1 = st[0].toString();
		} catch (Exception ep) {
			chckbxtaskCategories_1 = "";
		}
		if (chckbxtaskCategories_1 != null) {

		} else {
			chckbxtaskCategories_1 = "";
		}

		try {

			checkBox_taskbody = task.getBody();
		} catch (Exception ep) {
			checkBox_taskbody = "";
		}
		if (checkBox_taskbody != null) {

		} else {
			checkBox_taskbody = "";
		}

		try {

			String[] st = task.getCategories();
			chckbxtaskCompanies = st[0].toString();
		} catch (Exception ep) {
			chckbxtaskCompanies = "";
		}
		if (chckbxtaskCompanies != null) {

		} else {
			chckbxtaskCompanies = "";
		}

		String input = chckbxtaskCompanies + checkBox_taskbody + chckbxtaskCategories_1 + checkBox_taskstartdate
				+ chckbxtaskenddate + checkBoxtasksub + chckbxtaskStatus + checkBoxtaskibilling;
		System.out.println("input : " + input);
		try {
			md = MessageDigest.getInstance("MD5");
		} catch (NoSuchAlgorithmException e) {

			e.printStackTrace();
		}
		byte[] messageDigest = md.digest(input.getBytes());
		BigInteger no = new BigInteger(1, messageDigest);
		String hashtext = no.toString(16);
		while (hashtext.length() < 32) {
			hashtext = "0" + hashtext;
		}

		return hashtext;
	}

	public String removefolder(String path) {

		String[] str1;
		if (OS.contains("windows")) {
			str1 = path.split("\\\\");
		} else {
			str1 = path.split(File.separator);
		}

		for (int j = 0; j < str1.length; j++) {

			if (j == 0) {
				path = str1[j];
			} else if (!(j == str1.length - 1)) {

				path = path + File.separator + str1[j];
			}
		}
		return path;
	}

	public String removefoldergmail(String path) {
		String[] str1 = path.split("/");

		for (int j = 0; j < str1.length; j++) {

			if (j == 0) {
				path = str1[j];
			} else if (j == str1.length - 1) {
			} else {
				path = path + "/" + str1[j];
			}
		}
		return path;
	}

	void testemldd(File fol, CustomTreeNode mainnode) {

		File[] files = fol.listFiles();

		CustomTreeNode child = new CustomTreeNode("<html><b>" + fol.getName().replace("]", "").trim());
		mainnode.add(child);
		child.filepath = fol.getAbsolutePath();
		int messagesize = files.length;
		foldercountcheck++;
		for (int j = 0; j < messagesize; j++) {

			if (files[j].isDirectory()) {

				testemldd(files[j], child);

			} else {

			}

		}
	}

	void addemldd(File fol) {

		File[] files = fol.listFiles();
		try {

			int messagesize = files.length;

			for (int j = 0; j < messagesize; j++) {

				if (files[j].isDirectory()) {

					addemldd(files[j]);

				} else {

					if (fileoptionm.equalsIgnoreCase("EML File (.eml)")) {
						if (getFileExtension(files[j]).equalsIgnoreCase("eml")) {
							hm.put(files[j].getAbsolutePath(), null);
						}
					} else if (fileoptionm.equalsIgnoreCase("EMLX File (.emlx)")) {
						if (getFileExtension(files[j]).equalsIgnoreCase("emlx")) {
							hm.put(files[j].getAbsolutePath(), null);
						}
					} else if (fileoptionm.equalsIgnoreCase("OFT File (.oft)")) {
						if (getFileExtension(files[j]).equalsIgnoreCase("oft")) {
							hm.put(files[j].getAbsolutePath(), null);
						}
					} else if (fileoptionm.equalsIgnoreCase("Message File (.msg)")) {
						if (getFileExtension(files[j]).equalsIgnoreCase("msg")) {
							hm.put(files[j].getAbsolutePath(), null);
						}
					} else if (fileoptionm.equalsIgnoreCase("Maildir")) {

						hm.put(files[j].getAbsolutePath(), null);

					}

				}

			}
		} catch (Exception e) {
			if (fol.isFile()) {
				hm.put(fol.getAbsolutePath(), null);
			}
		}
	}

	int convertMillisintohour(int Millis) {

		int i = Millis / (1000 * 60 * 60);
		return i;
	}

	int convertMillisintomin(int Millis) {
		int i = (Millis % (1000 * 60 * 60)) / (1000 * 60);
		return i;
	}

	String contactname(MapiContact ap) {
		String s = "";
		try {
			s = getRidOfIllegalFileNameCharacters(ap.getSubject());
		} catch (Exception e) {
			s = "Contact";
		}
		s = getRidOfIllegalFileNameCharacters(s);
		return s;

	}

	String duplicacymapiContact(MapiContact Contact) {
		String checkBoxsub = null;
		String chckbxFullName = null;
		String chckbxEmailAddress = null;
		String chckbxMobilenumber = null;
		String chckbxJobtitle = null;
		String chckbxLocation = null;
		String chckbxCompany = null;
		String chckbxBirthday = null;
		MessageDigest md = null;

		try {
			new MapiContactNamePropertySet();
			chckbxFullName = Contact.getNameInfo().getDisplayName().replaceAll("[\\[\\]]", "").toString();
		} catch (Exception ep) {
			chckbxFullName = "";
		}
		if (chckbxFullName != null) {
		} else {
			chckbxFullName = "";
		}

		try {

			Contact.getElectronicAddresses().getDefaultEmailAddress();
			chckbxEmailAddress = MapiContactElectronicAddress.to_MapiContactElectronicAddress(
					Contact.getElectronicAddresses().getDefaultEmailAddress().toString()).toString();

		} catch (Exception ep) {
			chckbxEmailAddress = "";
		}

		if (chckbxEmailAddress != null) {
		} else {
			chckbxEmailAddress = "";
		}

		try {
			checkBoxsub = Contact.getSubject().toString();
		} catch (Exception ep) {
			checkBoxsub = "";
		}

		if (checkBoxsub != null) {
		} else {
			checkBoxsub = "";
		}

		try {
			chckbxMobilenumber = Contact.getTelephones().getMobileTelephoneNumber().toString();
		} catch (Exception e) {
			chckbxMobilenumber = "";
		}
		if (chckbxMobilenumber != null) {
		} else {
			chckbxMobilenumber = "";
		}

		try {
			chckbxJobtitle = Contact.getProfessionalInfo().getTitle().toString();
		} catch (Exception ep) {
			chckbxJobtitle = "";
		}

		if (chckbxJobtitle != null) {
		} else {
			chckbxJobtitle = "";
		}

		try {
			chckbxLocation = Contact.getPersonalInfo().getLocation().toString();
		} catch (Exception ep) {
			chckbxLocation = "";
		}
		if (chckbxLocation != null) {
		} else {
			chckbxLocation = "";
		}

		try {

			chckbxJobtitle = Contact.getProfessionalInfo().getCompanyName().toString();
		} catch (Exception ep) {
			chckbxCompany = "";
		}
		if (chckbxCompany != null) {
		} else {
			chckbxCompany = "";
		}

		try {
			chckbxBirthday = Contact.getEvents().getBirthday().toString();
		} catch (Exception ep) {
			chckbxBirthday = "";
		}
		if (chckbxBirthday != null) {
		} else {
			chckbxBirthday = "";
		}
		String value = chckbxBirthday + chckbxCompany + chckbxLocation + chckbxJobtitle + chckbxJobtitle
				+ chckbxEmailAddress + chckbxFullName + checkBoxsub;
		String input = value.replace(" ", "");
		try {
			md = MessageDigest.getInstance("MD5");
		} catch (NoSuchAlgorithmException e) {

			e.printStackTrace();
		}
		byte[] messageDigest = md.digest(input.getBytes());

		BigInteger no = new BigInteger(1, messageDigest);
		String hashtext = no.toString(16);
		while (hashtext.length() < 32) {
			hashtext = "0" + hashtext;
		}
		return hashtext;
	}

	boolean checkdate(MapiMessage message1, MailMessage mess) {
		for (int k = 0; k < fromList.size(); k++) {
			fromdate = fromList.get(k);
			todate = toList.get(k);
			Date d = message1.getDeliveryTime();
			String s = d.toString();

			if (s.length() == 28) {
				checkDate = true;
			} else {
				checkDate = false;
			}
			if (checkDate) {
				if (message1.getDeliveryTime() != null) {
					System.out.println(" DeliveryTime : " + message1.getDeliveryTime());
					if (message1.getDeliveryTime().after(fromdate) && message1.getDeliveryTime().before(todate)) {
						datevalidflag = true;
						break;
					} else {
						datevalidflag = false;
					}
				}
			} else {
				if (mess.getDate() != null) {
					System.out.println(" Mess : " + mess.getDate());
					if (mess.getDate().after(fromdate) && mess.getDate().before(todate)) {
						datevalidflag = true;
						break;
					} else {
						datevalidflag = false;
					}
				}
			}
		}

		return datevalidflag;
	}

	private String bytes2String(long sizeInBytes) {
		NumberFormat nf = new DecimalFormat();
		nf.setMaximumFractionDigits(2);

		try {
			if (sizeInBytes < SPACE_KB) {
				return nf.format(sizeInBytes) + " Byte(s)";
			} else if (sizeInBytes < SPACE_MB) {
				return nf.format(sizeInBytes / SPACE_KB) + " KB";
			} else if (sizeInBytes < SPACE_GB) {
				return nf.format(sizeInBytes / SPACE_MB) + " MB";
			} else if (sizeInBytes < SPACE_TB) {
				return nf.format(sizeInBytes / SPACE_GB) + " GB";
			} else {
				return nf.format(sizeInBytes / SPACE_TB) + " TB";
			}
		} catch (Exception e) {
			return sizeInBytes + " Byte(s)";
		}
	}

	//
	String namingconventionmapi(MapiMessage msg) {
		String filename = null;
		String frm;
		try {
			frm = msg.getSenderEmailAddress();
		} catch (Exception ep) {
			frm = "Na";
		}
		if (frm != null) {

		} else {
			frm = "Na";
		}

		if (frm.length() > 20) {
			frm = frm.substring(0, 20);
		}

		String sub;
		try {
			sub = msg.getSubject();
		} catch (Exception ep) {
			sub = "Na";
		}

		if (sub != null) {

		} else {
			sub = "Na";
		}

		if (sub.length() > 40) {
			sub = sub.substring(0, 40);
		}
		String dstr = "";
		Date d;
		String combox_selected = comboBox.getSelectedItem().toString();
		try {
			d = msg.getDeliveryTime();
			Calendar cal = Calendar.getInstance();
			cal.setTime(d);

			DecimalFormat formatter = new DecimalFormat("00");

			int date = cal.get(Calendar.DATE);
			String dateformate = formatter.format(date);

			int month = cal.get(Calendar.MONTH);
			month++;
			String monthformate = formatter.format(month);

			int year = cal.get(Calendar.YEAR);
			if (combox_selected.contains("DD-MM-YYYY")) {

				dstr = dateformate + "-" + monthformate + "-" + year;
			} else if (combox_selected.contains("MM-DD-YYYY")) {

				dstr = monthformate + "-" + dateformate + "-" + year;
			} else if (combox_selected.contains("YYYY-MM-DD")) {

				dstr = year + "-" + monthformate + "-" + dateformate;
			} else if (combox_selected.contains("YYYY-DD-MM")) {

				dstr = year + "-" + dateformate + "-" + monthformate;
			}

		} catch (Exception ep) {
			dstr = "Na";
		}

		if (combox_selected.equalsIgnoreCase("Subject")) {
			filename = sub;
		} else if (combox_selected.contains("From_Subject_Date")) {
			filename = frm + "_" + sub + "_" + dstr;
		} else if (combox_selected.contains("Subject_Date")) {
			filename = sub + "_" + dstr;
		} else if (combox_selected.contains("Date_Subject")) {
			filename = dstr + "_" + sub;
		} else if (combox_selected.contains("Date_From_Subject")) {
			filename = dstr + "_" + frm + "_" + sub;
		}
		filename = getRidOfIllegalFileNameCharacters(filename);
		return filename;
	}

	String namingconventionmapi(MapiMessage msg, File file) {
		String filename = "";
		String frm;
		try {
			frm = msg.getSenderEmailAddress();
		} catch (Exception ep) {
			frm = "";
		}
		if (frm != null) {

		} else {
			frm = "";
		}

		if (frm.length() > 20) {
			frm = frm.substring(0, 20);
		}

		String sub;
		try {
			sub = msg.getSubject();
		} catch (Exception ep) {
			sub = "";
		}

		if (sub != null) {

		} else {
			sub = "";
		}

		if (sub.length() > 80) {
			sub = sub.substring(0, 80);
		}

		String dstr = "";
		Date d;
		String combox_selected = comboBox.getSelectedItem().toString();
		try {
			d = msg.getDeliveryTime();
			Calendar cal = Calendar.getInstance();
			cal.setTime(d);
			DecimalFormat formatter = new DecimalFormat("00");

			int date = cal.get(Calendar.DATE);
			String dateformate = formatter.format(date);

			int month = cal.get(Calendar.MONTH);
			month++;
			String monthformate = formatter.format(month);

			int year = cal.get(Calendar.YEAR);
			if (combox_selected.contains("DD-MM-YYYY")) {

				dstr = dateformate + "-" + monthformate + "-" + year;
			} else if (combox_selected.contains("MM-DD-YYYY")) {

				dstr = monthformate + "-" + dateformate + "-" + year;
			} else if (combox_selected.contains("YYYY-MM-DD")) {

				dstr = year + "-" + monthformate + "-" + dateformate;
			} else if (combox_selected.contains("YYYY-DD-MM")) {

				dstr = year + "-" + dateformate + "-" + monthformate;
			}

		} catch (Exception ep) {
			dstr = "";
		}

		if (combox_selected.contains("From_Subject_Date")) {
			filename = frm + "_" + sub + "_" + dstr;
		} else if (combox_selected.equalsIgnoreCase("Subject")) {
			filename = sub;
		} else if (combox_selected.contains("Subject_Date")) {
			filename = sub + "_" + dstr;
		} else if (combox_selected.contains("Date_Subject")) {
			filename = dstr + "_" + sub;
		} else if (combox_selected.contains("Date_From_Subject")) {
			filename = dstr + "_" + frm + "_" + sub;
		} else if (combox_selected.equalsIgnoreCase("Original File Name")) {
			filename = file.getName().replace(".msg", "").replace(".eml", "").replace(".emlx", "");
		}

		filename = getRidOfIllegalFileNameCharacters(filename);
		return filename;
	}

	// Gmail
	public void ConvertPSTOST_gmail() {

		pst = PersonalStorage.fromFile(filepath);
		MailConversionOptions options = new MailConversionOptions();

		FolderInfo folderInfo2 = pst.getRootFolder();
		String Folder = folderInfo2.getDisplayName();
		Folder = Folder.replace(",", "").replace(".", "");
		Folder = getRidOfIllegalFileNameCharacters(Folder);
		Folder = Folder.replaceAll("[\\[\\]]", "");
		Folder = Folder.trim();
		if (Folder.equalsIgnoreCase("")) {
			Folder = "Root Folder";
		}
		if (filetype.equalsIgnoreCase("GoDaddy email")) {
			Folder = Folder.replaceAll("[^a-zA-Z0-9]", "");

		}

		String path1 = Folder;
//		path = path + "/" + Folder;
		parent1 = path + "/" + Folder;

		if (clientforimap_output.existFolder(parent1)) {
			clientforimap_output.selectFolder(iconnforimap_output, parent1);
		} else {
			clientforimap_output.createFolder(iconnforimap_output, parent1);
			clientforimap_output.selectFolder(iconnforimap_output, parent1);
		}

		listdupliccal.clear();
		listduplicacy.clear();
		listdupliccontact.clear();
		listduplictask.clear();

		MessageInfoCollection messageInfoCollection1 = folderInfo2.getContents();
		int countr = 0;
		int messagesize1;
		boolean s2 = false;
		if (demo) {
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

				if (stop) {
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
				MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
				MailConversionOptions de = new MailConversionOptions();
				MailMessage mess1 = message1.toMailMessage(de);

				MapiMessage message = MapiMessage.fromMailMessage(mess1, d);

				MailMessage mess = message.toMailMessage(options);
				if (chckbxMigrateOrBackup.isSelected()) {
					mess1.getAttachments().clear();
					mess.getAttachments().clear();
					message1.getAttachments().clear();
				}
				if (main_multiplefile.datefilter.isSelected()) {
					datevalidflag = checkdate(message, mess);
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
							mapi.setBody(con.getBody());
						} catch (Exception e) {

							mapi.setBody("");
						}

						try {
							message1.setSenderEmailAddress(mess.getFrom().toString());
							mapi.setFrom(mess.getFrom());
						} catch (Exception e) {
							e.printStackTrace();
						}

						con.save(temppathm + File.separator + namingconventionmapi(message1) + "_" + i + ".vcf",
								ContactSaveFormat.VCard);
						File file = new File(
								temppathm + File.separator + namingconventionmapi(message1) + "_" + i + ".vcf");
						mapi.addAttachment(new Attachment(
								temppathm + File.separator + namingconventionmapi(message1) + "_" + i + ".vcf"));
						file.delete();

						if (chckbxRemoveDuplicacy.isSelected()) {

							String input = duplicacymapiContact(con);
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listdupliccontact.contains(input)) {
								System.out.println("Not a duplicate message");
								listdupliccontact.add(input);

								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										clientforimap_output.appendMessage(iconnforimap_output, parent1, mapi);
										count_destination++;
									}
								} else {
									clientforimap_output.appendMessage(iconnforimap_output, parent1, mapi);
									count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforimap_output.appendMessage(iconnforimap_output, parent1, mapi);
									count_destination++;
								}
							} else {
								clientforimap_output.appendMessage(iconnforimap_output, parent1, mapi);
								count_destination++;
							}
						}
						countr++;
					} catch (OutOfMemoryError ep) {
						ep.printStackTrace();
						mf.logger.info("Out of memory error:" + ep.getMessage() + "  " + namingconventionmapi(message));
					} catch (Exception e) {
						e.printStackTrace();
						System.out.println(e.getMessage() + "  ===> 270 Contact " + Folder + count_destination);
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
							Progressbar.setVisible(false);
							i--;
						}

						while (!checkInternet()) {
							Progressbar.setText("connecting to server...");
						}

						connectionHandle(e.getMessage());
						mf.logger.warning(e.getMessage() + "Contact" + " " + countr + System.lineSeparator());
						continue;
					}

				} else if (message1.getMessageClass().equals("IPM.Appointment")
						|| message1.getMessageClass().contains("IPM.Schedule.Meeting.Request")
								&& !message1.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
					MailMessage mapi = new MailMessage();
					try {

						MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();

						try {
							mapi.setSubject(cal.getSubject() + "_" + i);
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
							e.printStackTrace();
						}

						cal.save(temppathm + File.separator + namingconventionmapi(message1)

								+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
						File file = new File(temppathm + File.separator + namingconventionmapi(message1)

								+ "_" + i + ".ics");

						mapi.addAttachment(new Attachment(temppathm + File.separator + namingconventionmapi(message1)

								+ "_" + i + ".ics"));
						file.delete();

						if (chckbxRemoveDuplicacy.isSelected()) {
							String input = duplicacymapiCal(cal);
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listdupliccal.contains(input)) {
								listdupliccal.add(input);

								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										clientforimap_output.appendMessage(iconnforimap_output, parent1, mapi);
										count_destination++;
									}
								} else {
									clientforimap_output.appendMessage(iconnforimap_output, parent1, mapi);
									count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforimap_output.appendMessage(iconnforimap_output, parent1, mapi);
									count_destination++;
								}
							} else {
								clientforimap_output.appendMessage(iconnforimap_output, parent1, mapi);
								count_destination++;
							}
						}
						countr++;
					} catch (OutOfMemoryError ep) {
						ep.printStackTrace();
						mf.logger.info("Out of memory error:" + ep.getMessage() + "  " + namingconventionmapi(message));
					} catch (Exception e) {
						e.printStackTrace();
						System.out.println(e.getMessage() + "  ===> 361 Calendar " + Folder + count_destination);
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
							Progressbar.setVisible(false);

							i--;
						}

						while (!checkInternet()) {
							Progressbar.setText("connecting to server...");
						}
						connectionHandle(e.getMessage());

						mf.logger.warning(e.getMessage() + "Calendar" + " " + countr + System.lineSeparator());
						continue;
					}

				} else if (message1.getMessageClass().equals("IPM.Task")) {

					String s = filetype;
					try {

						MapiTask task = (MapiTask) message1.toMapiMessageItem();
						MailMessage mapi = new MailMessage();
						File file = new File(temppathm + namingconventionmapi(message1) + i + ".msg");
						file.createNewFile();

						message1.save(file.getAbsolutePath(), SaveOptions.getDefaultMsg());
						mess.addAttachment(new Attachment(temppathm + namingconventionmapi(message1) + i + ".msg"));
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

						if (chckbxRemoveDuplicacy.isSelected()) {
							String input = "";
							if (message1.getMessageClass().equals("IPM.Task")) {
								input = duplicacymapiTask(task);
							}
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listduplictask.contains(input)) {
								listduplictask.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag)
										clientforimap_output.appendMessage(iconnforimap_output, parent1, mess);
									count_destination++;
								} else {
									clientforimap_output.appendMessage(iconnforimap_output, parent1, mess);
									count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag)
									clientforimap_output.appendMessage(iconnforimap_output, parent1, mess);
								count_destination++;
							} else {
								clientforimap_output.appendMessage(iconnforimap_output, parent1, mess);
								count_destination++;
							}
						}
					} catch (OutOfMemoryError ep) {
						ep.printStackTrace();
						mf.logger.info("Out of memory error:" + ep.getMessage() + "  " + namingconventionmapi(message));
					} catch (Exception e) {
						e.printStackTrace();
						System.out.println(e.getMessage() + "  ===> 430 Task " + Folder + count_destination);
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
							Progressbar.setVisible(false);

							i--;
						}
						while (!checkInternet()) {
							Progressbar.setText("connecting to server...");
						}
						connectionHandle(e.getMessage());
						mf.logger.warning(e.getMessage() + "Task" + " " + countr + System.lineSeparator());
						continue;
					} finally {
						filetype = s;
					}

				} else {
					try {
						if (message1.getMessageClass().equals("IPM.StickyNote")) {
							mess.setSubject(mess.getSubject() + "_" + i);
						}
						String messageid = mailimap(mess, parent1);
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
						ep.printStackTrace();
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						e.printStackTrace();
						StringWriter sw = new StringWriter();
						e.printStackTrace(new PrintWriter(sw));
						String exceptionAsString = sw.toString();

						System.out.println(e.getMessage() + Folder + count_destination);
						if (exceptionAsString.contains("Message too large")) {
							File f = new File((System.getProperty("user.home") + File.separator + "Desktop")
									+ File.separator + calendertime + File.separator + "Attachment" + File.separator
									+ mf.namingconventionmapi(message));
							f.mkdirs();
							mf.logger.info(
									"Message size was greater than allowed size so attachment has been deleted and saved in "
											+ f.getAbsolutePath());
							for (MapiAttachment attachment : message.getAttachments()) {

								attachment.save(f.getAbsolutePath() + File.separator
										+ getRidOfIllegalFileNameCharacters(attachment.getLongFileName()));

							}
							try {
								mess.getAttachments().clear();

								System.out.println(mess.getAttachments().size());

								String messageid = mailimap(mess, parent1);
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

								e1.printStackTrace();
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
							Progressbar.setVisible(false);

							i--;
						}
						while (!checkInternet()) {
							Progressbar.setText("connecting to server...");
						}
						connectionHandle(e.getMessage());

						mf.logger.warning(
								e.getMessage() + "Message" + " " + message.getDeliveryTime() + System.lineSeparator());
						continue;
					}

				}
				lbl_progressreport.setText("Total message Saved Count " + count_destination + "  " + Folder
						+ " Extarcting messsage " + message.getSubject());

			} catch (Exception e) {
				e.printStackTrace();
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
					Progressbar.setVisible(false);
//					i--;
				}
				while (!checkInternet()) {
					Progressbar.setText("connecting to server...");
				}
				connectionHandle(e.getMessage());

				mf.logger.warning(e.getMessage() + System.lineSeparator());

				continue;

			}

		}

		FolderInfoCollection folderInf = pst.getRootFolder().getSubFolders();

		for (int j = 0; j < folderInf.size(); j++) {
			try {

				if (stop) {
					break;
				}
				boolean s22 = false;
				FolderInfo folderInfo = folderInf.get_Item(j);
				String Folder1 = folderInfo.getDisplayName();
				Folder1 = Folder1.replace(",", "").replace(".", "");
				Folder1 = getRidOfIllegalFileNameCharacters(Folder1);
				Folder1 = Folder1.replaceAll("[\\[\\]]", "");
				Folder1 = Folder1.trim();

				Folder = path1 + File.separator + Folder1;

				String sfolder = Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (stop) {
						break;
					}
					if (pstfolderlist.get(l).equalsIgnoreCase(Folder)) {

						lbl_progressreport.setText("Getting Folder " + Folder);

						String fol = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");

						if (filetype.equalsIgnoreCase("GoDaddy email")) {
							fol = fol.replaceAll("[^a-zA-Z0-9]", "");

						}

//						path = path + "/" + fol;

						// Zoho
						if (filetype.equalsIgnoreCase("Zoho Mail")) {
							if (fol.equalsIgnoreCase("Inbox") || fol.equalsIgnoreCase("Drafts")
									|| fol.equalsIgnoreCase("Outbox")) {
//								path = path + "/" + fol + "_" + "zoho";
								path = parent1 + "/" + fol + "_" + "zoho";
							} else {
//								path = path + "/" + fol;
								path = parent1 + "/" + fol;
							}
						}
						// Yandex
						else if (filetype.equalsIgnoreCase("Yandex Mail")) {
//							path = path + "|" + fol.substring(0, 4);
							path = parent1 + "|" + fol.substring(0, 4);
						} else {
//							path = path + "/" + fol;
							path = parent1 + "/" + fol;
						}

						System.out.println(path + "  179 ");

						try {
							if (clientforimap_output.existFolder(path)) {
								clientforimap_output.selectFolder(path);
							} else {
								clientforimap_output.createFolder(iconnforimap_output, path);
								clientforimap_output.selectFolder(iconnforimap_output, path);
							}
							folder_check = true;
						} catch (Exception e3) {
							folder_check = false;
							System.out.println(e3.getMessage());
							e3.printStackTrace();
							if (e3.getMessage().contains("Operation failed")
									|| e3.getMessage().contains("Connection reset by peer: socket write error")) {
								l--;
								connectionHandle(e3.getMessage());
							}

						}

						listdupliccal.clear();
						listduplicacy.clear();
						listdupliccontact.clear();
						listduplictask.clear();
						if (folder_check) {

							MessageInfoCollection messageInfoCollection = folderInfo.getContents();
							int messagesize;
							if (demo) {
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
									if (stop) {
										break;
									}
									if ((i % 100) == 0) {
										System.gc();
									}
//									if ((count_destination % 1000) == 0) {
//										if (s22) {
//											connectionHandle1();

//									connectionHandle  replace with above**********
//										}
//										s22 = true;
//									}

									MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

									MapiMessage message1 = pst.extractMessage(messageInfo);
									MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
									MailConversionOptions de = new MailConversionOptions();
									MailConversionOptions options1 = new MailConversionOptions();
									MailMessage mess = message1.toMailMessage(options1);
									if (chckbxMigrateOrBackup.isSelected()) {
										mess.getAttachments().clear();
										message1.getAttachments().clear();
									}
									if (main_multiplefile.datefilter.isSelected()) {
										datevalidflag = checkdate(message1, mess);
									}
									MapiMessage message = MapiMessage.fromMailMessage(mess, d);

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
												mapi.setBody(con.getBody());
											} catch (Exception e) {

												mapi.setBody("");
											}
											try {
												message1.setSenderEmailAddress(mess.getFrom().toString());
												mapi.setFrom(mess.getFrom());
											} catch (Exception e) {
											}
											con.save(temppathm + File.separator + namingconventionmapi(message) + "_"
													+ i + ".vcf", ContactSaveFormat.VCard);
											File file = new File(temppathm + File.separator
													+ namingconventionmapi(message) + "_" + i + ".vcf");
											mapi.addAttachment(new Attachment(temppathm + File.separator
													+ namingconventionmapi(message) + "_" + i + ".vcf"));
											file.delete();

											if (chckbxRemoveDuplicacy.isSelected()) {

												String input = duplicacymapiContact(con);
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

										} catch (OutOfMemoryError ep) {
											ep.printStackTrace();
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ namingconventionmapi(message));
										} catch (Exception e) {
											e.printStackTrace();
											System.out.println(
													e.getMessage() + "  ===> 714" + Folder1 + count_destination);
											if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
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
												Progressbar.setVisible(false);
												i--;
											}
											while (!checkInternet()) {
												Progressbar.setText("connecting to server...");
											}
											connectionHandle(e.getMessage());
											mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " " + i
													+ namingconventionmapi(message) + System.lineSeparator());
											continue;
										}
									} else if (message1.getMessageClass().equals("IPM.Appointment")
											|| message1.getMessageClass().contains("IPM.Schedule.Meeting.Request")
													&& !message1.getMessageClass()
															.contains("IPM.Schedule.Meeting.Request.NDR")) {
										MailMessage mapi = new MailMessage();
										try {

											MapiCalendar cal = null;
											try {
												cal = (MapiCalendar) message1.toMapiMessageItem();
											} catch (Exception e2) {

											}

											try {
												mapi.setSubject(cal.getSubject() + "_" + i);
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
//											try {
//												message1.setSenderEmailAddress(message.getSenderEmailAddress());
//												mapi.setFrom(new MailAddress(message.getSenderEmailAddress()));
//											} catch (Exception e1) {
//												message1.setSenderEmailAddress(mess.getFrom().toString());
//												mapi.setFrom(new MailAddress(mess.getFrom().toString()));
//											}
												e.printStackTrace();
											}
											cal.save(temppathm + File.separator + namingconventionmapi(message) + "_"
													+ i + ".ics", AppointmentSaveFormat.Ics);
											File file = new File(temppathm + File.separator
													+ namingconventionmapi(message) + "_" + i + ".ics");

											mapi.addAttachment(new Attachment(temppathm + File.separator
													+ namingconventionmapi(message) + "_" + i + ".ics"));
											file.delete();
											if (chckbxRemoveDuplicacy.isSelected()) {
												String input = duplicacymapiCal(cal);
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

										} catch (OutOfMemoryError ep) {
											ep.printStackTrace();
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ namingconventionmapi(message));
										} catch (Exception e) {
											e.printStackTrace();
											System.out.println(
													e.getMessage() + "  ===> 1309" + Folder + count_destination);
											StringWriter sw = new StringWriter();
											e.printStackTrace(new PrintWriter(sw));
											String exceptionAsString = sw.toString();

											System.out.println(e.getMessage() + "-=-====" + Folder + count_destination);
											if (exceptionAsString.contains("Message too large")
													|| exceptionAsString.contains("TOOBIG ") || exceptionAsString
															.contains("The operation 'AppendMessage' terminated")) {
												File f1 = new File(
														(System.getProperty("user.home") + File.separator + "Desktop")
																+ File.separator + calendertime + File.separator
																+ "Attachment" + File.separator
																+ namingconventionmapi(message));
												f1.mkdirs();
												mf.logger.info(
														"Message size was greater than allowed size so attachment has been deleted and saved in "
																+ f1.getAbsolutePath());
												for (MapiAttachment attachment : message1.getAttachments()) {

													try {
														attachment.getLongFileName();
														System.out.println(attachment.getLongFileName() + " ==="
																+ attachment.getDisplayName());
														attachment.save(f1.getAbsolutePath().trim() + File.separator
																+ getRidOfIllegalFileNameCharacters(
																		attachment.getLongFileName()));
													} catch (Exception e1) {
														e1.printStackTrace();
														attachment.save(f1.getAbsolutePath().trim() + File.separator
																+ getRidOfIllegalFileNameCharacters(
																		attachment.getDisplayName()));
													}

												}
												try {
													mess.getAttachments().clear();
													System.out.println(mess.getAttachments().size());

													MailMessage mapi1 = new MailMessage();
													MapiCalendar cal = null;
													try {
														cal = (MapiCalendar) message1.toMapiMessageItem();
													} catch (Exception e1) {

													}

													try {
														mapi1.setSubject(cal.getSubject() + "_" + i);
													} catch (Exception e1) {
														mapi1.setSubject("");
													}
													try {
														mapi1.setBody(cal.getBody());
													} catch (Exception e1) {
														mapi1.setBody("");
													}

													try {
														message1.setSenderEmailAddress(
																cal.getOrganizer().getDisplayName());
														mapi1.setFrom(
																new MailAddress(cal.getOrganizer().getDisplayName()));
													} catch (Exception e1) {
													}

													cal.save(temppathm + File.separator + namingconventionmapi(message)
															+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
													File file = new File(temppathm + File.separator
															+ namingconventionmapi(message) + "_" + i + ".ics");

													mapi1.addAttachment(new Attachment(temppathm + File.separator
															+ namingconventionmapi(message) + "_" + i + ".ics"));
													if (chckbxRemoveDuplicacy.isSelected()) {
														String input = duplicacymapiCal(cal);
														input = input.replaceAll("\\s", "");
														input = input.trim();
														if (!listdupliccal.contains(input)) {
															listdupliccal.add(input);
															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	clientforimap_output.appendMessage(
																			iconnforimap_output, path, mapi1);
																	count_destination++;
																}
															} else {
																clientforimap_output.appendMessage(iconnforimap_output,
																		path, mapi1);
																count_destination++;
															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																clientforimap_output.appendMessage(iconnforimap_output,
																		path, mapi1);
																count_destination++;
															}
														} else {
															clientforimap_output.appendMessage(iconnforimap_output,
																	path, mapi1);
															count_destination++;
														}
													}
													file.delete();

												} catch (Exception e1) {

													e1.printStackTrace();
												}

											} else if (e.getMessage()
													.contains("The operation 'FetchMessage' terminated.")
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
												Progressbar.setVisible(false);
												i--;
											}
											while (!checkInternet()) {
												Progressbar.setText("connecting to server...");
											}
											connectionHandle(e.getMessage());
											mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + i
													+ namingconventionmapi(message) + System.lineSeparator());
											continue;
										}
									} else if (message1.getMessageClass().equals("IPM.Task")) {
										String s = filetype;
										try {

											MapiTask task = (MapiTask) message1.toMapiMessageItem();
											MailMessage mapi = new MailMessage();
											File file = new File(
													temppathm + namingconventionmapi(message1) + "_" + i + ".msg");
											file.createNewFile();

											message1.save(file.getAbsolutePath(), SaveOptions.getDefaultMsg());
											mess.addAttachment(new Attachment(
													temppathm + namingconventionmapi(message1) + "_" + i + ".msg"));
											file.delete();

											if (mess.getBody() != null) {
												mapi.setBody(mess.getBody());
											} else {
												mapi.setBody(message1.getBody());
											}
											try {
												mess.setSubject(mess.getSubject() + "_" + i);
											} catch (Exception e1) {
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

											if (chckbxRemoveDuplicacy.isSelected()) {
												String input = "";
												if (message1.getMessageClass().equals("IPM.Task")) {
													input = duplicacymapiTask(task);
												}
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listduplictask.contains(input)) {
													System.out.println("Not a duplicate message");
													listduplictask.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforimap_output.appendMessage(iconnforimap_output,
																	path, mess);
															count_destination++;
														}
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
													clientforimap_output.appendMessage(iconnforimap_output, path, mess);
													count_destination++;
												}
											}

										} catch (OutOfMemoryError ep) {
											ep.printStackTrace();
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ namingconventionmapi(message));
										} catch (Exception e) {
											e.printStackTrace();
											System.out.println(
													e.getMessage() + "  ===> 892" + Folder1 + count_destination);
											if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
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
												Progressbar.setVisible(false);

												i--;

											}
											while (!checkInternet()) {
												Progressbar.setText("connecting to server...");
											}
											connectionHandle(e.getMessage());
											mf.logger.warning("Exception : " + e.getMessage() + "Task" + " " + i
													+ namingconventionmapi(message) + System.lineSeparator());
											continue;
										} finally {
											filetype = s;
										}

									} else {
										try {
											if (message1.getMessageClass().equals("IPM.StickyNote")) {
												mess.setSubject(mess.getSubject() + "_" + i);
											}

//											

											String messageid = mailimap(mess, path);
											System.out.println(count_destination);
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
											ep.printStackTrace();
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ mf.namingconventionmapi(message));
										} catch (Exception e) {
											e.printStackTrace();
											System.out.println(Folder1 + count_destination + "      @@@@@@@@@@@@@@@  "
													+ e.getMessage());

											StringWriter sw = new StringWriter();
											e.printStackTrace(new PrintWriter(sw));
											String exceptionAsString = sw.toString();

											System.out.println(e.getMessage());
											if (exceptionAsString.contains("Message too large")) {
												File f = new File(
														(System.getProperty("user.home") + File.separator + "Desktop")
																+ File.separator + calendertime + File.separator
																+ "Attachment" + File.separator
																+ mf.namingconventionmapi(message));
												f.mkdirs();
												mf.logger.info(
														"Message size was greater than allowed size so attachment has been deleted and saved in "
																+ f.getAbsolutePath());
												for (MapiAttachment attachment : message.getAttachments()) {

													attachment.save(f.getAbsolutePath() + File.separator
															+ getRidOfIllegalFileNameCharacters(
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

													e1.printStackTrace();
												}

											} else if (e.getMessage()
													.contains("The operation 'FetchMessage' terminated.")
													|| e.getMessage().contains("ConnectFailure")
													|| e.getMessage().contains("Operation has been canceled")
													// || e.getMessage().contains("Operation failed.")
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

												Progressbar.setVisible(false);

												i--;

											}
//											else if(e.getMessage()
//													.contains( "Literal too large")){
//												i++;
//												
//											}

											connectionHandle(e.getMessage());
											mf.logger.warning("Exception : " + e.getMessage() + "Message " + " " + i
													+ mf.namingconventionmapi(message) + System.lineSeparator());
											e.printStackTrace();
											continue;
										}

									}
									lbl_progressreport.setText("  Total Message Saved Count  " + count_destination
											+ "  " + Folder + "   Extracting messsage " + message.getSubject());

								} catch (Exception e) {
									e.printStackTrace();
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
											|| e.getMessage()
													.contains("Connection reset by peer: socket write error")) {
										Progressbar.setVisible(false);

//									i--;

									}
									while (!checkInternet()) {
										Progressbar.setText("connecting to server...");
									}
									connectionHandle(e.getMessage());
									continue;
								}

							}
						}

					}
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderforpstost_gmail(folderInfo, sfolder, path);
				}
				path = removefoldergmail(path);
			} catch (Exception e) {
				continue;
			}

		}

	}

	public void getsubfolderforpstost_gmail(FolderInfo f, String sfolder, String path) {
		FolderInfoCollection subfolder = f.getSubFolders();
		String path11 = "";
		for (int k = 0; k < subfolder.size(); k++) {
			try {
				if (stop) {
					break;
				}
				FolderInfo folderf = subfolder.get_Item(k);

				String Folder = folderf.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();
				if (filetype.equalsIgnoreCase("GoDaddy email")) {
					Folder = Folder.replaceAll("[^a-zA-Z0-9]", "");

				}

				//
				if (filetype.equalsIgnoreCase("Zoho Mail")) {
					if (Folder.equalsIgnoreCase("Inbox") || Folder.equalsIgnoreCase("Drafts")
							|| Folder.equalsIgnoreCase("Outbox")) {
						Folder = Folder + "_" + "zoho";
					} else {
						Folder = Folder;
					}
				} else if (filetype.equalsIgnoreCase("Yandex Mail")) {
					Folder = Folder.substring(0, 4);
				}

				//

				sfolder = sfolder + File.separator + Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (stop) {
						break;
					}

					if (pstfolderlist.get(l).equalsIgnoreCase(sfolder)) {

//						path11 = path + "/" + Folder;
//						if (filetype.equalsIgnoreCase("Zoho Mail")) {
//							if (Folder.equalsIgnoreCase("Inbox") || Folder.equalsIgnoreCase("Drafts")
//									|| Folder.equalsIgnoreCase("Outbox")) {
//								path11 = path + "/" + Folder + "_" + "zoho";
//							} else {
//								path11 = path + "/" + Folder;
//							}
//						}
//						// Yandex
//						else if (filetype.equalsIgnoreCase("Yandex Mail")) {
//							path11 = path + "|" + Folder.substring(0, 4);
//						} else {
////							path11 = path + "/" + Folder;
//
//						}
						String new_path = path4 + "\\" + main_multiplefile.fname + "\\" + sfolder;
						String[] p1 = new_path.split("\\\\");
						x1 = path4;
						lbl_progressreport.setText("Getting : " + Folder);
//						clientforimap_output.createFolder(iconnforimap_output, path11);
//						clientforimap_output.selectFolder(iconnforimap_output, path11);
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
								folder_check = true;
							} catch (Exception e) {
								folder_check = false;
								System.out.println(e.getMessage());
								e.printStackTrace();
								if (e.getMessage().contains("Operation failed")
										|| e.getMessage().contains("Connection reset by peer: socket write error")) {
									l--;
									while (!checkInternet()) {
										Progressbar.setText("connecting to server...");
									}
									connectionHandle(e.getMessage());
								}

							}
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
							if (demo) {
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
									if (stop) {
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
									MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
									MailConversionOptions de = new MailConversionOptions();
									MailConversionOptions options = new MailConversionOptions();
									MailMessage mess = message1.toMailMessage(options);
									if (chckbxMigrateOrBackup.isSelected()) {
										mess.getAttachments().clear();
										message1.getAttachments().clear();
									}

									if (main_multiplefile.datefilter.isSelected()) {
										datevalidflag = checkdate(message1, mess);
									}
									MapiMessage message = MapiMessage.fromMailMessage(mess, d);
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
												mapi.setBody(con.getBody());
											} catch (Exception e) {

												mapi.setBody("");
											}
											try {
												message1.setSenderEmailAddress(mess.getFrom().toString());
												mapi.setFrom(mess.getFrom());
											} catch (Exception e) {
											}
											con.save(temppathm + File.separator + namingconventionmapi(message) + "_"
													+ i + ".vcf", ContactSaveFormat.VCard);
											File file = new File(temppathm + File.separator
													+ namingconventionmapi(message) + "_" + i + ".vcf");
											mapi.addAttachment(new Attachment(temppathm + File.separator
													+ namingconventionmapi(message) + "_" + i + ".vcf"));
											file.delete();

											if (chckbxRemoveDuplicacy.isSelected()) {

												String input = duplicacymapiContact(con);
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listdupliccontact.contains(input)) {
													System.out.println("Not a duplicate message");
													listdupliccontact.add(input);

													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforimap_output.appendMessage(iconnforimap_output, x1,
																	mapi);
															count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, x1,
																mapi);
														count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, x1,
																mapi);
														count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, x1, mapi);
													count_destination++;
												}
											}

										} catch (Exception e) {
											e.printStackTrace();
											System.out.println(
													e.getMessage() + "  ===> 1200" + Folder + count_destination);
											if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
													|| e.getMessage().equalsIgnoreCase("ConnectFailure")
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
												Progressbar.setVisible(false);
												i--;

											}
											while (!checkInternet()) {
												Progressbar.setText("connecting to server...");
											}
											connectionHandle(e.getMessage());
											mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
													+ System.lineSeparator());
											continue;
										}

									} else if (message1.getMessageClass().equals("IPM.Appointment")
											|| message1.getMessageClass().contains("IPM.Schedule.Meeting.Request")
													&& !message1.getMessageClass()
															.contains("IPM.Schedule.Meeting.Request.NDR")) {

										try {
											MailMessage mapi = new MailMessage();
											MapiCalendar cal = null;
											try {
												cal = (MapiCalendar) message1.toMapiMessageItem();
											} catch (Exception e1) {

											}

											try {
												mapi.setSubject(cal.getSubject() + "_" + i);
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
											}

											cal.save(temppathm + File.separator + namingconventionmapi(message) + "_"
													+ i + ".ics", AppointmentSaveFormat.Ics);
											File file = new File(temppathm + File.separator
													+ namingconventionmapi(message) + "_" + i + ".ics");

											mapi.addAttachment(new Attachment(temppathm + File.separator
													+ namingconventionmapi(message) + "_" + i + ".ics"));
											if (chckbxRemoveDuplicacy.isSelected()) {
												String input = duplicacymapiCal(cal);
												input = input.replaceAll("\\s", "");
												input = input.trim();
												if (!listdupliccal.contains(input)) {
													listdupliccal.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforimap_output.appendMessage(iconnforimap_output, x1,
																	mapi);
															count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, x1,
																mapi);
														count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, x1,
																mapi);
														count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, x1, mapi);
													count_destination++;
												}
											}
											file.delete();
										} catch (Exception e) {
											e.printStackTrace();
											System.out.println(
													e.getMessage() + "  ===> 1309" + Folder + count_destination);
											StringWriter sw = new StringWriter();
											e.printStackTrace(new PrintWriter(sw));
											String exceptionAsString = sw.toString();

											System.out.println(e.getMessage() + "-=-====" + Folder + count_destination);
											if (exceptionAsString.contains("Message too large")
													|| exceptionAsString.contains("TOOBIG ")) {
												File f1 = new File(
														(System.getProperty("user.home") + File.separator + "Desktop")
																+ File.separator + calendertime + File.separator
																+ "Attachment" + File.separator
																+ mf.namingconventionmapi(message));
												f1.mkdirs();
												mf.logger.info(
														"Message size was greater than allowed size so attachment has been deleted and saved in "
																+ f1.getAbsolutePath());
												for (MapiAttachment attachment : message.getAttachments()) {

													attachment.save(f1.getAbsolutePath() + File.separator
															+ getRidOfIllegalFileNameCharacters(
																	attachment.getLongFileName()));

												}
												try {
													mess.getAttachments().clear();
													System.out.println(mess.getAttachments().size());

													MailMessage mapi = new MailMessage();
													MapiCalendar cal = null;
													try {
														cal = (MapiCalendar) message1.toMapiMessageItem();
													} catch (Exception e1) {

													}

													try {
														mapi.setSubject(cal.getSubject() + "_" + i);
													} catch (Exception e1) {
														mapi.setSubject("");
													}
													try {
														mapi.setBody(cal.getBody());
													} catch (Exception e1) {
														mapi.setBody("");
													}

													try {
														message1.setSenderEmailAddress(
																cal.getOrganizer().getDisplayName());
														mapi.setFrom(
																new MailAddress(cal.getOrganizer().getDisplayName()));
													} catch (Exception e1) {
													}

													cal.save(temppathm + File.separator + namingconventionmapi(message)
															+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
													File file = new File(temppathm + File.separator
															+ namingconventionmapi(message) + "_" + i + ".ics");

													mapi.addAttachment(new Attachment(temppathm + File.separator
															+ namingconventionmapi(message) + "_" + i + ".ics"));
													if (chckbxRemoveDuplicacy.isSelected()) {
														String input = duplicacymapiCal(cal);
														input = input.replaceAll("\\s", "");
														input = input.trim();
														if (!listdupliccal.contains(input)) {
															listdupliccal.add(input);
															if (main_multiplefile.datefilter.isSelected()) {
																if (datevalidflag) {
																	clientforimap_output.appendMessage(
																			iconnforimap_output, x1, mapi);
																	count_destination++;
																}
															} else {
																clientforimap_output.appendMessage(iconnforimap_output,
																		x1, mapi);
																count_destination++;
															}
														}
													} else {
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																clientforimap_output.appendMessage(iconnforimap_output,
																		x1, mapi);
																count_destination++;
															}
														} else {
															clientforimap_output.appendMessage(iconnforimap_output, x1,
																	mapi);
															count_destination++;
														}
													}
													file.delete();

												} catch (Exception e1) {

													e1.printStackTrace();
												}

											} else if (e.getMessage()
													.contains("The operation 'FetchMessage' terminated.")
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
												Progressbar.setVisible(false);
												i--;
											}
											while (!checkInternet()) {
												Progressbar.setText("connecting to server...");
											}
											connectionHandle(e.getMessage());
											mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " "
													+ System.lineSeparator());
											continue;
										}

									} else if (message1.getMessageClass().equals("IPM.Task")) {
										try {
											MapiTask task = (MapiTask) message1.toMapiMessageItem();
											MailMessage mapi = new MailMessage();
											File file = new File(
													temppathm + namingconventionmapi(message1) + "_" + i + ".msg");
											file.createNewFile();

											message1.save(file.getAbsolutePath(), SaveOptions.getDefaultMsg());
											mess.addAttachment(new Attachment(
													temppathm + namingconventionmapi(message1) + "_" + i + ".msg"));
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
												mess.setSubject(mess.getSubject() + "_" + i);
											} catch (Exception e) {
											}
											try {
												mapi.setDate(mess.getDate());
											} catch (Exception e) {
												mapi.setDate(null);
											}

											if (chckbxRemoveDuplicacy.isSelected()) {
												String input = "";
												if (message1.getMessageClass().equals("IPM.Task")) {
													input = duplicacymapiTask(task);
												}
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listduplictask.contains(input)) {
													listduplictask.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag)
															clientforimap_output.appendMessage(iconnforimap_output, x1,
																	mess);
														count_destination++;
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, x1,
																mess);
														count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag)
														clientforimap_output.appendMessage(iconnforimap_output, x1,
																mess);
													count_destination++;
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, x1, mess);
													count_destination++;
												}
											}
										} catch (Exception e) {
											e.printStackTrace();
											System.out.println(
													e.getMessage() + "  ===> 1370" + Folder + count_destination);
											if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
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
												Progressbar.setVisible(false);
												i--;
											}
											while (!checkInternet()) {
												Progressbar.setText("connecting to server...");
											}
											connectionHandle(e.getMessage());
											mf.logger.warning("Exception : " + e.getMessage() + "Task" + " "
													+ System.lineSeparator());
											continue;
										}

									} else {
										try {
											if (message1.getMessageClass().equals("IPM.StickyNote")) {
												mess.setSubject(mess.getSubject() + "_" + i);
											}
											String messageid = mailimap(mess, x1);

											System.out.println(count_destination);
//											if (!messageid.equalsIgnoreCase("")) {
//												if (((message.getFlags()
//														& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
//													clientforimap_output.changeMessageFlags(iconnforimap_output,
//															messageid, ImapMessageFlags.isRead());
//												} else {
//													clientforimap_output.removeMessageFlags(iconnforimap_output,
//															messageid, ImapMessageFlags.isRead());
//												}
//											}
										} catch (Exception e) {
											e.printStackTrace();
											StringWriter sw = new StringWriter();
											e.printStackTrace(new PrintWriter(sw));
											String exceptionAsString = sw.toString();

											System.out.println(e.getMessage() + "-=-====" + Folder + count_destination);
											if (exceptionAsString.contains("Message too large")) {
												File f1 = new File(
														(System.getProperty("user.home") + File.separator + "Desktop")
																+ File.separator + calendertime + File.separator
																+ "Attachment" + File.separator
																+ mf.namingconventionmapi(message));
												f1.mkdirs();
												mf.logger.info(
														"Message size was greater than allowed size so attachment has been deleted and saved in "
																+ f1.getAbsolutePath());
												for (MapiAttachment attachment : message.getAttachments()) {

													attachment.save(f1.getAbsolutePath() + File.separator
															+ getRidOfIllegalFileNameCharacters(
																	attachment.getLongFileName()));

												}
												try {
													mess.getAttachments().clear();

													System.out.println(mess.getAttachments().size());

													String messageid = mailimap(mess, x1);
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

											} else if (e.getMessage()
													.contains("The operation 'FetchMessage' terminated.")
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
												Progressbar.setVisible(false);

												i--;

											}
											while (!checkInternet()) {
												Progressbar.setText("connecting to server...");
											}
											connectionHandle(e.getMessage());
											mf.logger.warning("Exception : " + e.getMessage() + "Message" + " "
													+ System.lineSeparator());
											continue;
										}

									}
									lbl_progressreport.setText("Total Message Saved Count  " + count_destination + "  "
											+ Folder + "   Extarcting messsage " + message.getSubject());

								} catch (Exception e) {
									e.printStackTrace();
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
											|| e.getMessage()
													.contains("Connection reset by peer: socket write error")) {
										Progressbar.setVisible(false);
//										i--;
									}
									while (!checkInternet()) {
										Progressbar.setText("connecting to server...");
									}
									connectionHandle(e.getMessage());
									continue;
								}

							}
						}
					}
				}
				if (folderf.hasSubFolders()) {
					getsubfolderforpstost_gmail(folderf, sfolder, path11);
				}

				path = removefoldergmail(path);
				sfolder = removefolder(sfolder);
			} catch (Exception e) {
				continue;
			}
		}

	}

	String mailimap(MailMessage message, String path) throws Exception {
		String Messageid = "";

		try {
			if (chckbxRemoveDuplicacy.isSelected()) {

				String input = duplicacymail(message);

				if (!listduplicacy.contains(input)) {
					System.out.println("Not a duplicate message");
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
		} catch (Exception e) {
			while (!checkInternet()) {
				Progressbar.setText("connecting to server...");
				lbl_progressreport.setText("connecting to server...");
			}
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

	// Imap
	public void ConvertPSTOST_imap() {

		System.out.println("Starting......");

		match = path;
		parent_path = path;
		count_destination = 0;
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

//		path = path + sepreter + Folder;
//		clientforimap_output.createFolder(iconnforimap_output, parent1);
//		clientforimap_output.selectFolder(iconnforimap_output, parent1);

		parent1 = path + sepreter + Folder;
		if (clientforimap_output.existFolder(parent1)) {
			clientforimap_output.selectFolder(iconnforimap_output, parent1);
		} else {
			clientforimap_output.createFolder(iconnforimap_output, parent1);
			clientforimap_output.selectFolder(iconnforimap_output, parent1);
		}

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

				if (stop) {
					break;
				}

				if ((i % 100) == 0) {
					System.gc();
				}
				if ((count_destination % 1000) == 0 && (count_destination != 0)) {
					if (s2) {
						connectionHandle1();
					}
					s2 = true;
				}

				MessageInfo messageInfo = (MessageInfo) messageInfoCollection1.get_Item(i);

				MapiMessage message1 = pst.extractMessage(messageInfo);
				MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
				MailConversionOptions de = new MailConversionOptions();
				MailMessage mess1 = message1.toMailMessage(de);
				if (chckbxMigrateOrBackup.isSelected()) {
					mess1.getAttachments().clear();
					message1.getAttachments().clear();
				}
				MapiMessage message = MapiMessage.fromMailMessage(mess1, d);
				if (main_multiplefile.datefilter.isSelected()) {
					datevalidflag = checkdate(message, mess1);
				}
				MailMessage mess = message.toMailMessage(options);
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
							mapi.setBody(con.getBody());
						} catch (Exception e) {

							mapi.setBody("");
						}
						try {
							message.setSenderEmailAddress(mess.getFrom().toString());
							mapi.setFrom(mess.getFrom());
						} catch (Exception e) {
						}

						con.save(temppathm + File.separator + namingconventionmapi(message) + "_" + count_destination
								+ ".vcf", ContactSaveFormat.VCard);
						File file = new File(temppathm + File.separator + namingconventionmapi(message) + "_"
								+ count_destination + ".vcf");
						mapi.addAttachment(new Attachment(temppathm + File.separator + namingconventionmapi(message)
								+ "_" + count_destination + ".vcf"));
						file.delete();

						if (chckbxRemoveDuplicacy.isSelected()) {

							String input = duplicacymapiContact(con);
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listdupliccontact.contains(input)) {
								listdupliccontact.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
										count_destination++;
									}
								} else {
									clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
									count_destination++;
								}

							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
									count_destination++;
								}
							} else {
								clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
								count_destination++;
							}
						}

						countr++;
					} catch (OutOfMemoryError ep) {
						mf.logger.info("Out of memory error:" + ep.getMessage() + "  " + namingconventionmapi(message));
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
							Progressbar.setVisible(false);
							i--;
						}
						connectionHandle(e.getMessage());
						mf.logger.warning(e.getMessage() + "Contact" + " " + countr + System.lineSeparator());
						continue;
					}

				} else if (message1.getMessageClass().equals("IPM.Appointment")
						|| message1.getMessageClass().contains("IPM.Schedule.Meeting.Request")
								&& !message1.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
					MailMessage mapi = new MailMessage();
					try {

						MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();

						try {
							mapi.setSubject(cal.getSubject() + "_" + i);
						} catch (Exception e) {

							mapi.setSubject("");
						}
						try {
							message.setSenderEmailAddress(mess.getFrom().toString());
							mapi.setFrom(mess.getFrom());
						} catch (Exception e) {
						}
						try {
							mapi.setBody(cal.getBody());
						} catch (Exception e) {

							mapi.setBody("");
						}
						cal.save(temppathm + File.separator + namingconventionmapi(message) + "_" + count_destination
								+ ".ics", AppointmentSaveFormat.Ics);
						File file = new File(temppathm + File.separator + namingconventionmapi(message) + "_"
								+ count_destination + ".ics");

						mapi.addAttachment(new Attachment(temppathm + File.separator + namingconventionmapi(message)
								+ "_" + count_destination + ".ics"));
						file.delete();

						if (chckbxRemoveDuplicacy.isSelected()) {
							String input = duplicacymapiCal(cal);
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listdupliccal.contains(input)) {
								listdupliccal.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
										count_destination++;
									}
								} else {
									clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
									count_destination++;
								}

							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
									count_destination++;
								}
							} else {
								clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
								count_destination++;
							}
						}
						countr++;
					} catch (OutOfMemoryError ep) {
						mf.logger.info("Out of memory error:" + ep.getMessage() + "  " + namingconventionmapi(message));
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
							Progressbar.setVisible(false);
							i--;
						}
						connectionHandle(e.getMessage());
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
						if (chckbxRemoveDuplicacy.isSelected()) {
							String input = "";
							if (message1.getMessageClass().equals("IPM.Task")) {
								input = duplicacymapiTask(task);
							}
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listduplictask.contains(input)) {
								System.out.println("Not a duplicate message");
								listduplictask.add(input);

								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										clientforimap_output.appendMessage(iconnforimap_output, path, mess);
										count_destination++;
									}
								} else {
									clientforimap_output.appendMessage(iconnforimap_output, path, mess);
									count_destination++;
								}

							}
						} else {

							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									clientforimap_output.appendMessage(iconnforimap_output, path, mess);
									count_destination++;
								}
							} else {
								clientforimap_output.appendMessage(iconnforimap_output, path, mess);
								count_destination++;
							}
						}
					} catch (OutOfMemoryError ep) {
						mf.logger.info("Out of memory error:" + ep.getMessage() + "  " + namingconventionmapi(message));
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
							Progressbar.setVisible(false);
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
						if (message1.getMessageClass().equals("IPM.StickyNote")) {
							mess.setFrom(new MailAddress(message1.getSenderEmailAddress()));
						}
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
						mf.logger.info("Out of memory error:" + ep.getMessage() + "  " + namingconventionmapi(message));
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
							Progressbar.setVisible(false);
							i--;
						}
						connectionHandle(e.getMessage());
						mf.logger.warning(
								e.getMessage() + "Message" + " " + message.getDeliveryTime() + System.lineSeparator());
						continue;
					}
				}
				lbl_progressreport.setText("Total message Saved Count " + count_destination + "  " + Folder
						+ " Extracting messsage " + message.getSubject());
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
					Progressbar.setVisible(false);
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

				if (stop) {
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

					if (stop) {
						break;
					}
					if (pstfolderlist.get(l).equalsIgnoreCase(Folder)) {
						sepreter = clientforimap_output.getDelimiter();

//						path = path + sepreter + folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");

						path = parent1 + sepreter + folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");

						lbl_progressreport.setText(" Getting Folder " + Folder);
						try {
							if (clientforimap_output.existFolder(path)) {
								clientforimap_output.selectFolder(path);
							} else {
								clientforimap_output.createFolder(iconnforimap_output, path);
								clientforimap_output.selectFolder(iconnforimap_output, path);
							}
							folder_check = true;
						} catch (Exception e1) {
							folder_check = false;
							System.out.println(e1.getMessage());
							e1.printStackTrace();
							if (e1.getMessage().contains("Operation failed")
									|| e1.getMessage().contains("Connection reset by peer: socket write error")) {
								l--;
								connectionHandle(e1.getMessage());
							}

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

								if (stop) {
									break;
								}
								if ((i % 100) == 0) {
									System.gc();
								}
								if ((count_destination % 1000) == 0 && (count_destination != 0)) {
									if (s22) {
										connectionHandle1();
									}
									s22 = true;
								}
								MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);
								MapiMessage message1 = pst.extractMessage(messageInfo);
								MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
								MailConversionOptions de = new MailConversionOptions();
								MailMessage mess = message1.toMailMessage(de);
								if (chckbxMigrateOrBackup.isSelected()) {
									mess.getAttachments().clear();
									message1.getAttachments().clear();
								}
								MapiMessage message = MapiMessage.fromMailMessage(mess, d);

								Date Receiveddate = message.getDeliveryTime();
								if (main_multiplefile.datefilter.isSelected()) {
									datevalidflag = checkdate(message, mess);
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
											message.setSenderEmailAddress(mess.getFrom().toString());
											mapi.setFrom(mess.getFrom());
										} catch (Exception e) {
										}
										try {
											mapi.setBody(con.getBody());
										} catch (Exception e) {
											mapi.setBody("");
										}
										con.save(temppathm + File.separator + namingconventionmapi(message) + "_"
												+ count_destination + ".vcf", ContactSaveFormat.VCard);
										File file = new File(temppathm + File.separator + namingconventionmapi(message)
												+ "_" + count_destination + ".vcf");
										mapi.addAttachment(new Attachment(temppathm + File.separator
												+ namingconventionmapi(message) + "_" + count_destination + ".vcf"));
										file.delete();

										if (chckbxRemoveDuplicacy.isSelected()) {

											String input = duplicacymapiContact(con);
											input = input.replaceAll("\\s", "");
											input = input.trim();

											if (!listdupliccontact.contains(input)) {
												System.out.println("Not a duplicate message");
												listdupliccontact.add(input);
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
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													count_destination++;
												}
											} else {
												clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
												count_destination++;
											}
										}
										countr++;
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ namingconventionmapi(message));
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
											Progressbar.setVisible(false);
											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning(
												e.getMessage() + "Contact" + " " + countr + System.lineSeparator());
										continue;
									}
								} else if (message1.getMessageClass().equals("IPM.Appointment") || message1
										.getMessageClass().contains("IPM.Schedule.Meeting.Request")
										&& !message1.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
									MailMessage mapi = new MailMessage();
									try {

										MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();

										try {
											mapi.setSubject(cal.getSubject());
										} catch (Exception e) {
											mapi.setSubject("");
										}
										try {
											message.setSenderEmailAddress(mess.getFrom().toString());
											mapi.setFrom(mess.getFrom());
										} catch (Exception e) {
											MapiElectronicAddress mrc = cal.getOrganizer();
											message1.setSenderEmailAddress(mrc.getEmailAddress());
											mapi.setFrom(MailAddress.to_MailAddress(mrc.getEmailAddress()));
										}
										try {
											mapi.setBody(cal.getBody());
										} catch (Exception e) {
											mapi.setBody("");
										}
										cal.save(temppathm + File.separator + namingconventionmapi(message) + "_"
												+ count_destination + ".ics", AppointmentSaveFormat.Ics);
										File file = new File(temppathm + File.separator + namingconventionmapi(message)
												+ "_" + count_destination + ".ics");

										mapi.addAttachment(new Attachment(temppathm + File.separator
												+ namingconventionmapi(message) + "_" + count_destination + ".ics"));
										file.delete();
										Receiveddate = cal.getStartDate();
										if (chckbxRemoveDuplicacy.isSelected()) {

											String input = duplicacymapiCal(cal);
											input = input.replaceAll("\\s", "");
											input = input.trim();

											if (!listdupliccal.contains(input)) {
												listdupliccal.add(input);
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
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
													count_destination++;
												}
											} else {
												clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
												count_destination++;
											}
										}
										countr++;
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ namingconventionmapi(message));
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
											Progressbar.setVisible(false);
											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning(
												e.getMessage() + "Calendar" + " " + countr + System.lineSeparator());
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

										if (chckbxRemoveDuplicacy.isSelected()) {
											String input = "";
											if (message1.getMessageClass().equals("IPM.Task")) {
												input = duplicacymapiTask(task);
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
														count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, path, mess);
													count_destination++;
												}
											}
										} else {

											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													clientforimap_output.appendMessage(iconnforimap_output, path, mess);
													count_destination++;
												}
											} else {
												clientforimap_output.appendMessage(iconnforimap_output, path, mess);
												count_destination++;
											}
										}
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ namingconventionmapi(message));
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
											Progressbar.setVisible(false);

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
												+ namingconventionmapi(message));
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
											Progressbar.setVisible(false);
											i--;
										}
										connectionHandle(e.getMessage());
										mf.logger.warning(e.getMessage() + "Message" + " " + message.getDeliveryTime()
												+ System.lineSeparator());
										continue;
									}
								}
								lbl_progressreport.setText("Total message Saved Count " + count_destination + "  "
										+ Folder + " Extracting messsage " + message.getSubject());
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
									Progressbar.setVisible(false);
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
					Progressbar.setVisible(false);
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

				if (stop) {
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
					if (stop) {
						break;
					}
					if (pstfolderlist.get(l).equalsIgnoreCase(sfolder)) {

						//
						sepreter = clientforimap_output.getDelimiter();
						String new_path = path4 + "\\" + main_multiplefile.fname + "\\"
								+ sfolder.replace(File.separator, "\\");
						System.out.println(new_path);
						String[] p1 = new_path.split("\\\\");
						x1 = path4;
						//
						lbl_progressreport.setText("Getting Folder " + Folder);

						for (int i1 = 1; i1 < p1.length; i1++) {

							x1 = x1 + sepreter + p1[i1];

							path = x1;

							try {
								if (clientforimap_output.existFolder(x1)) {
									clientforimap_output.selectFolder(x1);
								} else {
									clientforimap_output.createFolder(x1);
									clientforimap_output.selectFolder(x1);
								}
								folder_check = true;
							} catch (Exception e) {
								folder_check = false;
								System.out.println(e.getMessage());
								e.printStackTrace();
								if (e.getMessage().contains("Operation failed")
										|| e.getMessage().contains("Connection reset by peer: socket write error")) {
									l--;
									connectionHandle(e.getMessage());
								}

							}
						}
						System.out.println(sepreter + "  sepreter >>");
//						path = parent_path + sepreter + sfolder.replace(File.separator, sepreter);

//						if (clientforimap_output.existFolder(path)) {
//							clientforimap_output.selectFolder(path);
//						} else {
//							clientforimap_output.createFolder(iconnforimap_output, path);
//							clientforimap_output.selectFolder(iconnforimap_output, path);
//						}

						listduplicacy.clear();
						listdupliccal.clear();
						listdupliccontact.clear();
						listduplictask.clear();
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
									if (stop) {
										break;
									}
									if ((i % 100) == 0) {
										System.gc();
									}
									if ((count_destination % 1000) == 0 && (count_destination != 0)) {
										if (s22) {
											connectionHandle1();
										}
										s22 = true;
									}

									MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

									MapiMessage message1 = pst.extractMessage(messageInfo);
									MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
									MailConversionOptions de = new MailConversionOptions();
									MailMessage mess = message1.toMailMessage(de);
									if (chckbxMigrateOrBackup.isSelected()) {
										mess.getAttachments().clear();
										message1.getAttachments().clear();
									}
									MapiMessage message = MapiMessage.fromMailMessage(mess, d);
									if (main_multiplefile.datefilter.isSelected()) {
										datevalidflag = checkdate(message, mess);
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
												message.setSenderEmailAddress(mess.getFrom().toString());
												mapi.setFrom(mess.getFrom());
											} catch (Exception e) {
											}
											con.save(temppathm + File.separator + namingconventionmapi(message) + "_"
													+ count_destination + ".vcf", ContactSaveFormat.VCard);
											File file = new File(temppathm + File.separator
													+ namingconventionmapi(message) + "_" + count_destination + ".vcf");
											mapi.addAttachment(new Attachment(
													temppathm + File.separator + namingconventionmapi(message) + "_"
															+ count_destination + ".vcf"));
											file.delete();

											if (chckbxRemoveDuplicacy.isSelected()) {

												String input = duplicacymapiContact(con);
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listdupliccontact.contains(input)) {
													System.out.println("Not a duplicate message");
													listdupliccontact.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforimap_output.appendMessage(iconnforimap_output, x1,
																	mapi);
															count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, x1,
																mapi);
														count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, x1,
																mapi);
														count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, x1, mapi);
													count_destination++;
												}
											}
											countr++;
										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ namingconventionmapi(message));
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
												Progressbar.setVisible(false);

												i--;
											}
											connectionHandle(e.getMessage());
											mf.logger.warning(
													e.getMessage() + "Contact" + " " + countr + System.lineSeparator());
											continue;
										}

									} else if (message1.getMessageClass().equals("IPM.Appointment")
											|| message1.getMessageClass().contains("IPM.Schedule.Meeting.Request")
													&& !message1.getMessageClass()
															.contains("IPM.Schedule.Meeting.Request.NDR")) {
										MailMessage mapi = new MailMessage();
										try {

											MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();

											try {
												mapi.setSubject(cal.getSubject());
											} catch (Exception e) {

												mapi.setSubject("");
											}
											try {
												message.setSenderEmailAddress(mess.getFrom().toString());
												mapi.setFrom(mess.getFrom());
											} catch (Exception e) {
											}
											try {
												mapi.setBody(cal.getBody());
											} catch (Exception e) {

												mapi.setBody("");
											}
											cal.save(temppathm + File.separator + namingconventionmapi(message) + "_"
													+ count_destination + ".ics", AppointmentSaveFormat.Ics);
											File file = new File(temppathm + File.separator
													+ namingconventionmapi(message) + "_" + count_destination + ".ics");

											mapi.addAttachment(new Attachment(
													temppathm + File.separator + namingconventionmapi(message) + "_"
															+ count_destination + ".ics"));
											file.delete();
											if (chckbxRemoveDuplicacy.isSelected()) {

												String input = duplicacymapiCal(cal);
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listdupliccal.contains(input)) {
													listdupliccal.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforimap_output.appendMessage(iconnforimap_output, x1,
																	mapi);
															count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, x1,
																mapi);
														count_destination++;
													}

												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, x1,
																mapi);
														count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, x1, mapi);
													count_destination++;
												}
											}
											countr++;
										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ namingconventionmapi(message));
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
												Progressbar.setVisible(false);

												i--;
											}
											connectionHandle(e.getMessage());
											mf.logger.warning(e.getMessage() + "Calendar" + " " + countr
													+ System.lineSeparator());
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
											if (chckbxRemoveDuplicacy.isSelected()) {
												String input = "";
												if (messageInfo.getMessageClass().equals("IPM.Task")) {
													input = duplicacymapiTask(task);
												}
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listduplictask.contains(input)) {
													System.out.println("Not a duplicate message");
													listduplictask.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															clientforimap_output.appendMessage(iconnforimap_output, x1,
																	mess);
															count_destination++;
														}
													} else {
														clientforimap_output.appendMessage(iconnforimap_output, x1,
																mess);
														count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														clientforimap_output.appendMessage(iconnforimap_output, x1,
																mess);
														count_destination++;
													}
												} else {
													clientforimap_output.appendMessage(iconnforimap_output, x1, mess);
													count_destination++;
												}
											}

										} catch (OutOfMemoryError ep) {
											mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
													+ namingconventionmapi(message));
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
												Progressbar.setVisible(false);
												i--;
											}
											connectionHandle(e.getMessage());
											mf.logger.warning(
													e.getMessage() + "Task" + " " + countr + System.lineSeparator());
											continue;
										} finally {
											filetype = s;
										}

									}

									else {
										try {
											if (message1.getMessageClass().equals("IPM.StickyNote")) {
												mess.setFrom(new MailAddress(message1.getSenderEmailAddress()));
											}
											String messageid = mailimap(mess, x1);
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
													+ namingconventionmapi(message));
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
												Progressbar.setVisible(false);

												i--;
											}
											connectionHandle(e.getMessage());
											mf.logger.warning(e.getMessage() + "Message" + " "
													+ message.getDeliveryTime() + System.lineSeparator());
											continue;
										}

									}
									lbl_progressreport.setText("Total message Saved Count " + count_destination + "  "
											+ Folder + " Extracting messsage " + message.getSubject());

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
										Progressbar.setVisible(false);

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
				sfolder = removefolder(sfolder);
			} catch (Exception e) {

				if (e.getMessage().contains("The operation 'FetchMessage' terminated.")
						|| e.getMessage().contains("Operation failed") || e.getMessage().contains("Rate limit hit")
						|| e.getMessage().contains("Operation has been canceled")
						|| e.getMessage().contains("The operation 'AppendMessage' terminated.")
						|| e.getMessage().contains("Software caused connection abort: recv failed")
						|| e.getMessage().contains("Timeout")
						|| e.getMessage().contains("Network is unreachable: connect")) {
					Progressbar.setVisible(false);

				}
				connectionHandle(e.getMessage());
			}

		}

	}

	@SuppressWarnings("deprecation")
	public IEWSClient conntiontooffice365_output(IEWSClient clientforexchange_output) throws Exception {
		while (true) {
			try {
				if (main_multiplefile.modern_Authentication.isSelected()) {
					String token = Refresh_Token.refreshinput();
					NetworkCredential credentials = new OAuthNetworkCredential(token);
					EWSClient.useSAAJAPI(true);
					clientforexchange_output = EWSClient.getEWSClient("https://outlook.office365.com/ews/exchange.asmx",
							credentials);
					clientforexchange_output.setTimeout(5 * 60 * 1000);
					EmailClient.setSocketsLayerVersion2(true);
					EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
				} else {
					clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3, password_p3);
					clientforexchange_output.setTimeout(5 * 60 * 1000);
				}
				System.out.println("Connection Done : ");
				break;
			} catch (Exception e) {
			}
		}
		return clientforexchange_output;
	}

	public void m1() {

		pst = PersonalStorage.fromFile(filepath);

		FolderInfo folderInfo2 = pst.getRootFolder();
		String Folder = folderInfo2.getDisplayName();
		Folder = Folder.replace(",", "").replace(".", "");
		Folder = getRidOfIllegalFileNameCharacters(Folder);
		Folder = Folder.replaceAll("[\\[\\]]", "");
		Folder = Folder.trim();
		if (Folder.equalsIgnoreCase("")) {
			Folder = "Root Folder";
		}
		if (filetype.equalsIgnoreCase("GoDaddy email")) {
			Folder = Folder.replaceAll("[^a-zA-Z0-9]", "");

		}
		String path1 = Folder;
//		String path_new = path + "/" + Folder;
//		parent_path = path;
		path = path + "/" + Folder;

		clientforimap_output.createFolder(iconnforimap_output, path);
		clientforimap_output.selectFolder(iconnforimap_output, path);

		listdupliccal.clear();
		listduplicacy.clear();
		listdupliccontact.clear();
		listduplictask.clear();

		FolderInfoCollection folderInf = pst.getRootFolder().getSubFolders();

		for (int j = 0; j < folderInf.size(); j++) {
			try {

				if (stop) {
					break;
				}
				boolean s22 = false;
				FolderInfo folderInfo = folderInf.get_Item(j);
				String Folder1 = folderInfo.getDisplayName();
				Folder1 = Folder1.replace(",", "").replace(".", "");
				Folder1 = getRidOfIllegalFileNameCharacters(Folder1);
				Folder1 = Folder1.replaceAll("[\\[\\]]", "");
				Folder1 = Folder1.trim();

				Folder = path1 + File.separator + Folder1;

				String sfolder = Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (stop) {
						break;
					}
					if (pstfolderlist.get(l).equalsIgnoreCase(Folder)) {

						lbl_progressreport.setText("Getting Folder " + Folder);

						String fol = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");

						if (filetype.equalsIgnoreCase("GoDaddy email")) {
							fol = fol.replaceAll("[^a-zA-Z0-9]", "");
						}
						path = path + "/" + fol;
						clientforimap_output.createFolder(iconnforimap_output, path);
						clientforimap_output.selectFolder(iconnforimap_output, path);

						listdupliccal.clear();
						listduplicacy.clear();
						listdupliccontact.clear();
						listduplictask.clear();

					}
				}
				if (folderInfo.hasSubFolders()) {
					m2(folderInfo, sfolder, path);
				}
				path = removefoldergmail(path);
			} catch (Exception e) {
				continue;
			}

		}

	}

	public void m2(FolderInfo f, String sfolder, String path) {

		FolderInfoCollection subfolder = f.getSubFolders();
		String path11 = "";
		for (int k = 0; k < subfolder.size(); k++) {
			try {
				if (stop) {
					break;
				}
				FolderInfo folderf = subfolder.get_Item(k);

				String Folder = folderf.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();
				if (filetype.equalsIgnoreCase("GoDaddy email")) {
					Folder = Folder.replaceAll("[^a-zA-Z0-9]", "");

				}

				sfolder = sfolder + File.separator + Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (stop) {
						break;
					}

					if (pstfolderlist.get(l).equalsIgnoreCase(sfolder)) {

						path11 = path + "/" + Folder;
						lbl_progressreport.setText("Getting : " + Folder);
						clientforimap_output.createFolder(iconnforimap_output, path11);
						clientforimap_output.selectFolder(iconnforimap_output, path11);

						listdupliccal.clear();
						listduplicacy.clear();
						listdupliccontact.clear();
						listduplictask.clear();

					}
				}
				if (folderf.hasSubFolders()) {
					m2(folderf, sfolder, path11);
				}

				path = removefoldergmail(path);
				sfolder = removefolder(sfolder);
			} catch (Exception e) {
				continue;
			}
		}

	}

	public static boolean checkInternet() {
		label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));
		System.out.println("Please connect internet connection !");
		try {
			URL url = new URL("http://www.google.com");
			URLConnection connection = url.openConnection();
			connection.connect();
			Progressbar.setText("");
			System.out.println("Internet is connected");
			return true;
		} catch (MalformedURLException e) {
		} catch (IOException e) {
		}
		return false;
	}

	public void officeEws365(int fileCounter) {

		pst = PersonalStorage.fromFile(filepath);
		FolderInfo folderInfo2 = pst.getRootFolder();
//		count_destination_total = 0;

		locdat = getRidOfIllegalFileNameCharacters(Calendar.getInstance().getTime().toString());
		locdat = locdat.replaceAll(":", " ");
		if (chckbxCustomFolderName.isSelected()) {
			String customerfolder = textField_customfolder.getText().replace("//s", "");

			customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
			locdat = customerfolder;

		}
		except = false;
		// int fileno has to be brought;

		try {

			if (fileCounter == 1) {
				firstFolderMaker(locdat.toString());
			}

			if (except == true) {
				stop = true;
				return;
			}
			// base folder being created here

//			try {
//				  Folder folderroot = new Folder(service);
//					folderroot.setDisplayName(locdat.toString());
//					
//				if (filterselected.equals("mailboxsel")) {
//					folderroot.save(WellKnownFolderName.MsgFolderRoot);
//					rootfolderid = folderroot.getId();
//				} else if (filterselected.equals("archivesel")) {
//					folderroot.save(WellKnownFolderName.ArchiveMsgFolderRoot);
//					rootfolderid = folderroot.getId();
//				} else if (filterselected.equals("publicfoldersel")) {
//					folderroot.save(WellKnownFolderName.PublicFoldersRoot);
//					rootfolderid = folderroot.getId();
//				}
//				
//			   
//			} catch (Exception e1) {
//				if (e1.getMessage().contains("Access is denied")) {
//					String warn = "You Have not access to Public folder ";
//					JOptionPane.showMessageDialog(main_multiplefile.this, warn, messageboxtitle,
//							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
//					except = true;
//					e1.printStackTrace();
//					return;
//				} else if (e1.getMessage().contains("The request failed. outlook.office365.com")
//						|| e1.getMessage().contains("The request failed. The request failed. Connection reset")
//						|| e1.getMessage().contains("Connection reset")
//						|| e1.getMessage().contains("outlook.office365.com")
//						|| e1.getMessage().contains("The request failed. java.net.SocketException: Connection reset")
//						|| e1.getMessage().contains(
//								"The request failed. The request failed. No such host is known (outlook.office365.com)")
//						|| e1.getMessage().contains("The request failed. The request failed. outlook.office365.com")
//						|| e1.getMessage()
//								.contains("The request failed. Network is unreachable: no further information")) {
//					e1.printStackTrace();
//					StringWriter sw = new StringWriter();
//					PrintWriter pw = new PrintWriter(sw);
//					e1.printStackTrace(pw);
//					mf.logger.warning(sw.toString());
//
//					connectionHandleoffice_output("");
//				} else if (e1.getMessage().contains("A folder with the specified name already exists.")) {
//					JOptionPane.showMessageDialog(main_multiplefile.this,
//							"This Folder Name is already exist please change custom Folder name and try again ",
//							messageboxtitle, JOptionPane.ERROR_MESSAGE,
//							new ImageIcon(Main_Frame.class.getResource("/information.png")));
//					textField_customfolder.setEditable(true);
//					except = true;
//					e1.printStackTrace();
//					return;
//				} else {
//					e1.printStackTrace();
//				}
//			}

			Folder secondFolder = secondFolderMaker(filepath);
			if (stop) {

				return;
			}
			if (secondFolder == null) {
				secondFolder = secondFolderMaker(filepath);
			}

			thirdFolderMaker(map, folderInfo2, secondFolder.getId());
			// make below code a function in order to handle connection
//			  try {
//				    Folder nameOfFile = new Folder(service);
//				    nameOfFile.setDisplayName(new File(filepath).getName().replaceAll("[\\[\\]]", "").trim());
//				    nameOfFile.save(rootfolderid);
//				   FolderId pstNameId=nameOfFile.getId();
//					Folder firstTopOfFolder=new Folder(service);
//					firstTopOfFolder.setDisplayName(folderInfo2.getDisplayName().replaceAll("[\\[\\]]", "").trim());
//					firstTopOfFolder.save(pstNameId);
//					FolderId base=firstTopOfFolder.getId();
//					System.out.println(folderInfo2.getDisplayName()+"  folderInfo2.getDisplayName()");
//				    map.put(folderInfo2.getDisplayName(), base);  
//					
//				   
//				} catch (Exception e) {
//
//					mf.logger.info(
//							"The request failed. The server cannot service this request right now. Try again later"
//									+ e.getMessage());
//
//					if (e.getMessage().contains(
//							"The request failed. The server cannot service this request right now. Try again later")) {
//						
//						if (stop) {
//							break;
//						}
//						connectionHandleoffice_output("");
//						e.printStackTrace();
//					} else if (e.getMessage().contains("The request failed. outlook.office365.com")
//							|| e.getMessage().contains(
//									"The request failed. The request failed. Connection reset")
//							|| e.getMessage().contains("Connection reset")
//							|| e.getMessage().contains("outlook.office365.com")
//							|| e.getMessage().contains(
//									"The request failed. java.net.SocketException: Connection reset")
//							|| e.getMessage().contains(
//									"The request failed. The request failed. No such host is known (outlook.office365.com)")
//							|| e.getMessage().contains(
//									"The request failed. The request failed. outlook.office365.com")
//							|| e.getMessage().contains(
//									"The request failed. Network is unreachable: no further information")) {
//						if (stop) {
//							break;
//						}
//						connectionHandleoffice_output("");
//						e.printStackTrace();
//					}
//					e.printStackTrace();
//				}

			// enter the root folder into the base folder
			// top of the personal folder to be created and data to be entered in it.
			// create all the folders that are part of heirarchy and enter the data only in
			// the last one.
			// if A folder is found in pstfolderlist try getting back its id and then enter
			// the data into it.
			foldermessagecount = 0;
			count_destination = 0;
			for (FolderInfo folder : folderInfo2.getSubFolders()) {

				path = folderInfo2.getDisplayName() + File.separator;

				String foldername = folder.getDisplayName();
				foldername = foldername.replace(",", "").replace(".", "");
				foldername = getRidOfIllegalFileNameCharacters(foldername);
				foldername = foldername.replaceAll("[\\[\\]]", "");
				foldername = foldername.trim();

				path += foldername;
//				String subfolder = Folderuri;
				System.out.println(path + "  path of main folder 16580");
				// below loop not needed
//				for (int l = 0; l < pstfolderlist.size(); l++) {
				if (stop) {
					break;
				}

				if (pstfolderlist.contains(path)) {

//						System.out.println("this is follist " + folderc + " path "+ path);
//						
//						String[] aa;
//						if (System.getProperty("os.name").toLowerCase().contains("windows")) {
//							aa = folderc.trim().split("\\\\");
//						} else {
//							aa = folderc.trim().split("/");
//						}

					// here reconstruct the whole path by creating all the folders while also
					// checking if it already exists or not
					// after that put data in last folder
					for (int i = 0; i < 1; i++) {// if (i == 0) {

						if (!map.containsKey(path)) {
							// if (!map.containsValue(aa[i])) {
							try {

								Folder folder1 = new Folder(service);
								folder1.setDisplayName(folder.getDisplayName().replaceAll("[\\[\\]]", "").trim());
								folder1.setFolderClass(folder.getContainerClass());
								System.out.println(folder.getContainerClass());
								System.out.println(folder1.getClass().toString());
								System.out.println(folder1.getClass().getName());
								FolderId parentFolderId = map.get(folderInfo2.getDisplayName());

//										String foldername = folder1.getDisplayName();
//										System.out.println("this is folder displsy name " + foldername);
//										if (foldername.contains("Tasks") || foldername.contains("Task")
//												|| foldername.contains("task")) {
//											folder1.setFolderClass("IPF.Task");
//											// folder1.save(folderid);
//										} else if (foldername.contains("Notes") || foldername.contains("Note")
//												|| foldername.contains("note") || foldername.contains("Note")) {
//											folder1.setFolderClass("IPF.Note");
//											// folder1.save(folderid);
//										} else if (path.contains("Contacts") || foldername.contains("contacts")
//												|| foldername.contains("Contacts") || foldername.contains("Contact")
//												|| foldername.contains("contact")
//												|| foldername.contains("Address Book")) {
//											folder1.setFolderClass("IPF.Contact");
//											// folder1.save(folderid);
//										} else if (path.contains("Calendar") || foldername.contains("Calendars")
//												|| foldername.contains("calendars") || foldername.contains("Birthday")
//												|| foldername.contains("Birthdays")
//												|| foldername.contains("United States holidays")
//												|| foldername.contains("holidays") || foldername.contains("Holidays")) {
//											folder1.setFolderClass("IPF.Appointment");
//											// folder1.save(folderid);
//
//										} else if (foldername.equals("Journal")) {
//											folder1.setFolderClass("IPF.Journal");
//											// folder1.save(folderid);
//										} else {
//											folder1.setFolderClass("IPF.Note");
//											// folder1.save(folderid);
//										}
//										folderid = folder1.getId();

								folder1.save(parentFolderId);
								System.out.println("folder created at path " + path);

								folderid = folder1.getId();
								map.put(path, folderid);
								MessageInfoCollection messageInfoCollection2 = folder.getContents();
								int messageco = messageInfoCollection2.size();

								if (messageco > 0) {

									addmessage(folder, folderid);

								}

							} catch (Exception e) {
								while (!checkInternet()) {
									Progressbar.setText("connecting to server...");
									lbl_progressreport.setText("Please chech your internet");
								}
								mf.logger.info(
										"The request failed. The server cannot service this request right now. Try again later"
												+ e.getMessage());

								if (e.getMessage().contains(
										"The request failed. The server cannot service this request right now. Try again later")) {
									i--;
									if (stop) {
										break;
									}
									connectionHandleoffice_output("");
									e.printStackTrace();
								} else if (e.getMessage().contains("The request failed. outlook.office365.com")
										|| e.getMessage()
												.contains("The request failed. The request failed. Connection reset")
										|| e.getMessage().contains("Connection reset")
										|| e.getMessage().contains("outlook.office365.com")
										|| e.getMessage().contains(
												"The request failed. java.net.SocketException: Connection reset")
										|| e.getMessage().contains(
												"The request failed. The request failed. No such host is known (outlook.office365.com)")
										|| e.getMessage().contains(
												"The request failed. The request failed. outlook.office365.com")
										|| e.getMessage().contains(
												"The request failed. Network is unreachable: no further information")) {
									if (stop) {
										break;
									}
									connectionHandleoffice_output("");
									e.printStackTrace();
								}
								e.printStackTrace();
							}
						}
					}
//							} else {
//								try {
//									if (aa[i].equalsIgnoreCase("INBOX")) {
//										aa[i] = aa[i].toLowerCase();
//									}
////									if (!map.containsKey(aa[i]) && map.containsKey(aa[i - 1])) {
//									if (!map.containsValue(aa[i]) && map.containsValue(aa[i - 1])) {
//										folderid = map.get(aa[i - 1]);
//										System.out.println("exist");
//										folder1 = new Folder(service);
//										folder1.setDisplayName(aa[i]);
//										//added extra by wasi
//										folder1.setFolderClass(folder.getContainerClass());
//										folder1.save(folderid);
//										folderid = folder1.getId();
//										if (pstfolderlist.get(l).equalsIgnoreCase(path)) {
//											addmessage(folder, folderid);
//										}
//										System.out.println("folder name : " + aa[i]);
//										map.put(aa[i], folderid);
//									} else {
//										folderid = map.get(aa[i]);
//									}
//								} catch (Exception e) {
//
//									mf.logger.info(
//											"The request failed. The server cannot service this request right now. Try again later"
//													+ e.getMessage());
//
//									if (e.getMessage().contains(
//											"The request failed. The server cannot service this request right now. Try again later")) {
//										i--;
//										connectionHandleoffice_output("");
//										e.printStackTrace();
//									} else if (e.getMessage().contains("The request failed. outlook.office365.com")
//											|| e.getMessage().contains(
//													"The request failed. The request failed. Connection reset")
//											|| e.getMessage().contains("Connection reset")
//											|| e.getMessage().contains("outlook.office365.com")
//											|| e.getMessage().contains(
//													"The request failed. java.net.SocketException: Connection reset")
//											|| e.getMessage().contains(
//													"The request failed. The request failed. No such host is known (outlook.office365.com)")
//											|| e.getMessage().contains(
//													"The request failed. The request failed. outlook.office365.com")
//											|| e.getMessage().contains(
//													"The request failed. Network is unreachable: no further information")) {
//
//										e.printStackTrace();
//										StringWriter sw = new StringWriter();
//										PrintWriter pw = new PrintWriter(sw);
//										e.printStackTrace(pw);
//										mf.logger.warning(sw.toString());
//										if (stop) {
//											break;
//										}
//										connectionHandleoffice_output("");
//									}
//									System.out.println("20573........");
//									e.printStackTrace();
//									StringWriter sw = new StringWriter();
//									PrintWriter pw = new PrintWriter(sw);
//									e.printStackTrace(pw);
//									mf.logger.warning(sw.toString());
//								}
//							}
//						}
					lbl_progressreport.setText("Getting Folder " + folder.getDisplayName().replaceAll("[\\[\\]]", ""));
					Folder = path;

				}
//				}

				if (folder.getSubFolders().size() > 0) {

					System.out.println("this is child folder calling ");

					getFolderforexchange(folder, path, subfolder);
				}
			}
		} catch (Exception e1) {
			while (!checkInternet()) {
				Progressbar.setText("connecting to server...");
				lbl_progressreport.setText("Please chech your internet");
			}
			if (e1.getMessage() != null) {
				if (e1.getMessage().contains("Access is denied")) {

					JOptionPane.showMessageDialog(main_multiplefile.this, e1.getMessage(), messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					return;
				} else if (e1.getMessage().contains("The request failed. outlook.office365.com")
						|| e1.getMessage().contains("The request failed. The request failed. Connection reset")
						|| e1.getMessage().contains("Connection reset")
						|| e1.getMessage().contains("outlook.office365.com")
						|| e1.getMessage().contains("The request failed. java.net.SocketException: Connection reset")
						|| e1.getMessage().contains(
								"The request failed. The request failed. No such host is known (outlook.office365.com)")
						|| e1.getMessage().contains("The request failed. The request failed. outlook.office365.com")
						|| e1.getMessage()
								.contains("The request failed. Network is unreachable: no further information")) {
					try {

						connectionHandleoffice_output("");
					} catch (Exception e) {
						e.printStackTrace();
						StringWriter sw = new StringWriter();
						PrintWriter pw = new PrintWriter(sw);
						e.printStackTrace(pw);
						mf.logger.warning(sw.toString());
					}
				}
				e1.printStackTrace();
			}

		} finally {
			pst.dispose();
			pst.close();
		}

	}

	@SuppressWarnings("resource")
	private void getFolderforexchange(FolderInfo folder, String pathx1, String subfol) {
		try {
			for (FolderInfo subFolderx1 : folder.getSubFolders()) {
				System.out.println("this is sub folder name " + subFolderx1.getDisplayName());

				String foldername = subFolderx1.getDisplayName();
				foldername = foldername.replace(",", "").replace(".", "");
				foldername = getRidOfIllegalFileNameCharacters(foldername);
				foldername = foldername.replaceAll("[\\[\\]]", "");
				foldername = foldername.trim();
				String pathx2 = pathx1 + File.separator + foldername;
				String subFold = foldername;
				Folder folder1 = null;
//				for (int l = 0; l < pstfolderlist.size(); l++) {
				if (stop) {
					break;
				}
				foldermessagecount = 0;
				if (pstfolderlist.contains(pathx2)) {
					String folderc = pathx2;
					System.out.println("this is follist in sub folder  " + folderc);
//						try {
					String[] aa;
					if (System.getProperty("os.name").toLowerCase().contains("windows")) {
						aa = folderc.split("\\\\");
					} else {
						aa = folderc.split("/");
					}
//							List<String> Foldername = new ArrayList<String>();
					String childPath = aa[0].substring(0);
					System.out.println(childPath + "  childpath");
					String parentPath = aa[0].substring(0);
					for (int i = 1; i < aa.length; i++) {

//								if (Foldername.contains(aa[i])) {
//									aa[i] = aa[i] + i;
//								} else {
//
//									Foldername.add(aa[i]);
//								}
						if (i > 1) {
							parentPath += File.separator + aa[i - 1];
						}
						childPath += File.separator + aa[i];
						// CHECK PATHN from begining to end
						System.out.println(parentPath + " parentPath");
//								if (i == 0) {
						if (!map.containsKey(childPath)) {
//									if (!map.containsValue(aa[i])) {
							try {
								folder1 = new Folder(service);
								folder1.setDisplayName(aa[i].replaceAll("[\\[\\]]", "").trim());
								// use key and get Value
								FolderId parent = map.get(parentPath);
								// even though subFolder class has to be reconstructed by making a map and
								// traversing only the folders.
								folder1.setFolderClass(subFolderx1.getContainerClass());
								folder1.save(parent);
								folderid = folder1.getId();
								map.put(childPath, folderid);
								if (pstfolderlist.contains(childPath) && i == aa.length - 1) {
									// only the selected folder data to be migrated
									MessageInfoCollection messageInfoCollection2 = subFolderx1.getContents();
									int messageco = messageInfoCollection2.size();

									if (messageco > 0) {
										addmessage(subFolderx1, folderid);
									}
								}

							} catch (Exception ex) {
								while (!checkInternet()) {
									Progressbar.setText("connecting to server...");
									lbl_progressreport.setText("Please chech your internet");
								}
								mf.logger.info(
										"The request failed. The server cannot service this request right now. Try again later"
												+ ex.getMessage());
								if (ex.getMessage().contains(
										"The request failed. The server cannot service this request right now. Try again later")) {
									System.out.println("line no 20856");
									i--;
									connectionHandleoffice_output("");
								} else if (ex.getMessage().contains("The request failed. outlook.office365.com")
										|| ex.getMessage()
												.contains("The request failed. The request failed. Connection reset")
										|| ex.getMessage().contains("Connection reset")
										|| ex.getMessage().contains("outlook.office365.com")
										|| ex.getMessage().contains(
												"The request failed. java.net.SocketException: Connection reset")
										|| ex.getMessage().contains(
												"The request failed. The request failed. No such host is known (outlook.office365.com)")
										|| ex.getMessage().contains(
												"The request failed. The request failed. outlook.office365.com")
										|| ex.getMessage().contains(
												"The request failed. Network is unreachable: no further information")) {
									if (stop) {
										break;
									}
									connectionHandleoffice_output("");
									System.out.println("line no 20877");
									i--;
								} else if (true) {
									// duplicate folder error
								}
								ex.printStackTrace();
							}
						} else {
							continue;
						}
//								} 
//								else {
//									try {
////										if (!map.containsKey(aa[i]) && map.containsKey(aa[i - 1])) {
//										if (!map.containsValue(aa[i]) && map.containsKey(aa[i - 1])) {
//											folderid = map.get(aa[i - 1]);
//											System.out.println("exist");
//
//											folder1 = new Folder(service);
//											folder1.setDisplayName(aa[i]);
//											
//											
//											folder1.setFolderClass(folder.getContainerClass());
//											
//											
//											
////											System.out.println("this is folder displsy name " + foldername);
////											if (path.contains("Tasks") || foldername.contains("Tasks")
////													|| foldername.contains("Task") || foldername.contains("task")) {
////												folder1.setFolderClass("IPF.Task");
////												folder1.save(folderid);
////											} else if (path.contains("Notes") || foldername.contains("Notes")
////													|| foldername.contains("Note") || foldername.contains("note")
////													|| foldername.contains("Note")) {
////												folder1.setFolderClass("IPF.Note");
////												folder1.save(folderid);
////											} else if (path.contains("Contacts") || foldername.contains("contacts")
////													|| foldername.contains("Contacts") || foldername.contains("Contact")
////													|| foldername.contains("contact")) {
////												folder1.setFolderClass("IPF.Contact");
////												folder1.save(folderid);
////											} else if (path.contains("Calendar") || foldername.contains("Calendars")
////													|| foldername.contains("calendars")
////													|| foldername.contains("Birthday")
////													|| foldername.contains("Birthdays")
////													|| foldername.contains("United States holidays")
////													|| foldername.contains("holidays")
////													|| foldername.contains("Holidays")) {
////												folder1.setFolderClass("IPF.Appointment");
////												folder1.save(folderid);
////
////											} else if (foldername.equals("Journal")) {
////												folder1.setFolderClass("IPF.Journal");
////												folder1.save(folderid);
////											} else {
////												folder1.setFolderClass("IPF.Note");
////												
////											}
//											
//											
//											
//											folder1.save(folderid);
//											folderid = folder1.getId();
//											if (path.contains("Accounts")) {
//
//												if (subFolder.getDisplayName().toString().equals(aa[i]) && i > 1) {
//													try {
//														addmessage(subFolder, folderid);
//													} catch (Exception e) {
//														if (e.getMessage().contains(
//																"The request failed. The server cannot service this request right now. Try again later")) {
//															i--;
//															if (stop) {
//																break;
//															}
//															connectionHandleoffice_output("");
//															e.printStackTrace();
//															StringWriter sw = new StringWriter();
//															PrintWriter pw = new PrintWriter(sw);
//															e.printStackTrace(pw);
//															mf.logger.warning(sw.toString());
//														} else
//															e.printStackTrace();
//														StringWriter sw = new StringWriter();
//														PrintWriter pw = new PrintWriter(sw);
//														e.printStackTrace(pw);
//														mf.logger.warning(sw.toString());
//													}
//												}
//											} else {
//												if (subFolder.getDisplayName().toString().equals(aa[i])) {
//													try {
//														addmessage(subFolder, folderid);
//													} catch (Exception ex) {
//														if (ex.getMessage()
//																.contains("The request failed. outlook.office365.com")
//																|| ex.getMessage().contains(
//																		"The request failed. The request failed. Connection reset")
//																|| ex.getMessage().contains("Connection reset")
//																|| ex.getMessage().contains("outlook.office365.com")
//																|| ex.getMessage().contains(
//																		"The request failed. java.net.SocketException: Connection reset")
//																|| ex.getMessage().contains(
//																		"The request failed. The request failed. No such host is known (outlook.office365.com)")
//																|| ex.getMessage().contains(
//																		"The request failed. The request failed. outlook.office365.com")
//																|| ex.getMessage().contains(
//																		"The request failed. Network is unreachable: no further information")) {
//															if (stop) {
//																break;
//															}
//															connectionHandleoffice_output("");
//															System.out.println("line no 20964");
//															i--;
//														} else if (ex.getMessage().contains(
//																"The request failed. The server cannot service this request right now. Try again later")) {
//															i--;
//															if (stop) {
//																break;
//															}
//															connectionHandleoffice_output("");
//															ex.printStackTrace();
//															StringWriter sw = new StringWriter();
//															PrintWriter pw = new PrintWriter(sw);
//															ex.printStackTrace(pw);
//															mf.logger.warning(sw.toString());
//														} else {
//
//															ex.printStackTrace();
//															StringWriter sw = new StringWriter();
//															PrintWriter pw = new PrintWriter(sw);
//															ex.printStackTrace(pw);
//															mf.logger.warning(sw.toString());
//														}
//													}
//												}
//
//											}
//
//											map.put(aa[i], folderid);
//											System.out.println(
//													"this is folder id line no 20905 " + aa[i] + " id " + folderid);
//										} else {
//											folderid = map.get(aa[i]);
//
////											if (subFolder.getName().toString().equals(aa[i])) {
////												folderid = map.get(aa[i - 1]);
////												folder1 = new Folder(service);
////												folder1.setDisplayName(aa[i]);
////												String foldername = folder1.getDisplayName();
////												if (foldername.contains("Tasks") || foldername.contains("Task")
////														|| foldername.contains("task")) {
////													folder1.setFolderClass("IPF.Task");
////													folder1.save(folderid);
////												} else if (foldername.contains("Notes") || foldername.contains("Note")
////														|| foldername.contains("note") || foldername.contains("Note")) {
////													folder1.setFolderClass("IPF.Note");
////													folder1.save(folderid);
////												} else if (path.contains("Contacts") || foldername.contains("contacts")
////														|| foldername.contains("Contacts")
////														|| foldername.contains("Contact")
////														|| foldername.contains("contact")) {
////													folder1.setFolderClass("IPF.Contact");
////													folder1.save(folderid);
////												} else if (path.contains("Calendar") || foldername.contains("Calendars")
////														|| foldername.contains("calendars")
////														|| foldername.contains("Birthday")
////														|| foldername.contains("Birthdays")
////														|| foldername.contains("United States holidays")
////														|| foldername.contains("holidays")
////														|| foldername.contains("Holidays")) {
////													folder1.setFolderClass("IPF.Appointment");
////													folder1.save(folderid);
////
////												} else if (foldername.equals("Journal")) {
////													folder1.setFolderClass("IPF.Journal");
////													folder1.save(folderid);
////												} else {
////													folder1.setFolderClass("IPF.Note");
////													folder1.save(folderid);
////												}
////												// folder1.save(folderid);
////												if (subFolder.getName().toString().equals(aa[i])) {
////													try {
////														addmessage(subFolder, folder1.getId());
////													} catch (Exception ex) {
////														if (ex.getMessage()
////																.contains("The request failed. outlook.office365.com")
////																|| ex.getMessage().contains(
////																		"The request failed. The request failed. Connection reset")
////																|| ex.getMessage().contains("Connection reset")
////																|| ex.getMessage().contains("outlook.office365.com")
////																|| ex.getMessage().contains(
////																		"The request failed. java.net.SocketException: Connection reset")
////																|| ex.getMessage().contains(
////																		"The request failed. The request failed. No such host is known (outlook.office365.com)")
////																|| ex.getMessage().contains(
////																		"The request failed. The request failed. outlook.office365.com")
////																|| ex.getMessage().contains(
////																		"The request failed. Network is unreachable: no further information")) {
////															connectionHandleoffice_output("");
////															i--;
////														} else if (ex.getMessage().contains(
////																"The request failed. The server cannot service this request right now. Try again later")) {
////															i--;
////															connectionHandleoffice_output("");
////															mf.logger.warning("Exception during adding message : "
////																	+ ex.getMessage());
////															ex.printStackTrace();
////														} else {
////
////															mf.logger.warning("Exception during adding message : "
////																	+ ex.getMessage());
////															ex.printStackTrace();
////														}
////													}
////
////												}
////
////												map.put(aa[i], folderid);
////											}
//
//											// System.out.println("this is folder id "+aa[i]+" id "+folderid);
//
//										}
//									} catch (Exception ex) {
//
//										if (ex.getMessage().contains(
//												"The request failed. The server cannot service this request right now. Try again later")) {
//											i--;
//											System.out.println("line no 21076");
//											if (stop) {
//												break;
//											}
//											connectionHandleoffice_output("");
//
//										} else if (ex.getMessage().contains("The request failed. outlook.office365.com")
//												|| ex.getMessage().contains(
//														"The request failed. The request failed. Connection reset")
//												|| ex.getMessage().contains("Connection reset")
//												|| ex.getMessage().contains("outlook.office365.com")
//												|| ex.getMessage().contains(
//														"The request failed. java.net.SocketException: Connection reset")
//												|| ex.getMessage().contains(
//														"The request failed. The request failed. No such host is known (outlook.office365.com)")
//												|| ex.getMessage().contains(
//														"The request failed. The request failed. outlook.office365.com")
//												|| ex.getMessage().contains(
//														"The request failed. Network is unreachable: no further information")) {
//											if (stop) {
//												break;
//											}
//											connectionHandleoffice_output("");
//											System.out.println("line no 21093");
//										} else if (ex.getMessage()
//												.contains("A folder with the specified name already exists.")) {
//
//											StringWriter sw = new StringWriter();
//											// PrintWriter pw = new PrintWriter(sw);
//											// ex.printStackTrace(pw);
//											mf.logger.warning(sw.toString());
//											System.out.println("line no 21101");
//											continue;
//										} else {
//											ex.printStackTrace();
//											StringWriter sw = new StringWriter();
//											PrintWriter pw = new PrintWriter(sw);
//											ex.printStackTrace(pw);
//											mf.logger.warning(sw.toString());
//										}
//
//									}
//								}
					}
//						} catch (Exception ex) {
//
//							if (ex.getMessage().contains("The request failed. outlook.office365.com")
//									|| ex.getMessage()
//											.contains("The request failed. The request failed. Connection reset")
//									|| ex.getMessage().contains("Connection reset")
//									|| ex.getMessage().contains("outlook.office365.com")
//									|| ex.getMessage()
//											.contains("The request failed. java.net.SocketException: Connection reset")
//									|| ex.getMessage().contains(
//											"The request failed. The request failed. No such host is known (outlook.office365.com)")
//									|| ex.getMessage()
//											.contains("The request failed. The request failed. outlook.office365.com")
//									|| ex.getMessage().contains(
//											"The request failed. Network is unreachable: no further information")) {
//								if (stop) {
//									break;
//								}
//								connectionHandleoffice_output("");
//								System.out.println("line no 21130");
//							} else if (ex.getMessage().contains(
//									"The request failed. The server cannot service this request right now. Try again later")) {
//
//								if (stop) {
//									break;
//								}
//								connectionHandleoffice_output("");
////								l--;
//								ex.printStackTrace();
//								StringWriter sw = new StringWriter();
//								PrintWriter pw = new PrintWriter(sw);
//								ex.printStackTrace(pw);
//								mf.logger.warning(sw.toString());
//							}
//
//							ex.printStackTrace();
//							StringWriter sw = new StringWriter();
//							PrintWriter pw = new PrintWriter(sw);
//							ex.printStackTrace(pw);
//							mf.logger.warning(sw.toString());
//						}

					Folder = pathx2;

					lbl_progressreport.setText("Getting Folder " + folder.getDisplayName().replaceAll("[\\[\\]]", ""));

				}
//				}
				// read sub-folders
				String oldPathx2 = pathx2;
				if (subFolderx1.getSubFolders().size() > 0) {

					getFolderforexchange(subFolderx1, pathx2, subFold);
				}
				pathx2 = oldPathx2;
//				path = removefolder(path);
			}

		} catch (Exception ex) {
			while (!checkInternet()) {
				Progressbar.setText("connecting to server...");
				lbl_progressreport.setText("Please chech your internet");
			}
			if (ex.getMessage().contains("The request failed. outlook.office365.com")
					|| ex.getMessage().contains("The request failed. The request failed. Connection reset")
					|| ex.getMessage().contains("Connection reset") || ex.getMessage().contains("outlook.office365.com")
					|| ex.getMessage().contains("The request failed. java.net.SocketException: Connection reset")
					|| ex.getMessage().contains(
							"The request failed. The request failed. No such host is known (outlook.office365.com)")
					|| ex.getMessage().contains("The request failed. The request failed. outlook.office365.com")
					|| ex.getMessage().contains("The request failed. Network is unreachable: no further information")) {
				try {
					connectionHandleoffice_output("");
					System.out.println("line no 22044");
				} catch (Exception e) {
					e.printStackTrace();
				}
				mf.logger.warning("Exception during adding message : " + ex.getMessage());
			} else if (ex.getMessage().contains(
					"The request failed. The server cannot service this request right now. Try again later")) {
				try {
					connectionHandleoffice_output("");
				} catch (Exception e) {
					e.printStackTrace();
				} catch (Error e) {
					e.printStackTrace();
				}
				mf.logger.warning("Exception during adding message : " + ex.getMessage());
				ex.printStackTrace();
			}
		}
	}

	public void addmessage(FolderInfo subFolder, FolderId folderidx2) throws Exception, Error {
		int mcount = 0;
		boolean selected = true;
		listduplicacy.clear();
		if (subFolder.getContents().size() > 0) {
			mf.listduplicacy.clear();
			try {
//				long foldercount = subFolder.getMessageCount();
				MessageInfoCollection messageInfoCollection2 = subFolder.getContents();
				int foldercount = messageInfoCollection2.size();
				System.out.println("total message of " + subFolder.getDisplayName() + " is : " + foldercount);

//				Iterator<MapiMessage> it = storage.enumerateMessages(subFolder).iterator();

				listduplicacy.clear();
				listdupliccal.clear();
				listdupliccontact.clear();

				for (int i11 = 0; i11 < foldercount; i11++) {
					if (stop) {
						break;
					}
					if (i11 == 50 && demo) {
						break;
					}

//					if (push) {
//						pushstart();
//					}

					try {

						MessageInfo messageInfo = (MessageInfo) messageInfoCollection2.get_Item(i11);

						MapiMessage message1 = pst.extractMessage(messageInfo);
						System.out.println(message1.getSubject() + " message sub " + i11);
						MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
						MailConversionOptions de = new MailConversionOptions();
						MailMessage mess = message1.toMailMessage(de);
						MapiMessage message = MapiMessage.fromMailMessage(mess, d);
						boolean isDraft = mess.isDraft();
						if (chckbxMigrateOrBackup.isSelected()) {
							if (message1.getAttachments().size() > 0) {
								message1.getAttachments().clear();
								mess.getAttachments().clear();
								message.getAttachments().clear();
							}
						}
						Date c = mess.getDate();
						String date = c.toString();
						String[] datearr = null;
						datearr = date.split("\\s");
						String[] time = datearr[3].split(":");
						int i = Integer.parseInt(time[0]);
						Calendar calendar = new GregorianCalendar();
						TimeZone timeZone = calendar.getTimeZone();
						TimeZone timezone = TimeZone.getTimeZone(timeZone.getID());
						int ms = timezone.getOffset(Calendar.ZONE_OFFSET);
						if (i < 1) {
							time[0] = String.valueOf(i + 12 + (convertMillisintohour(ms)));
						} else {
							time[0] = String.valueOf(i + (convertMillisintohour(ms)));
						}
						int i1 = Integer.parseInt(time[1]);
						time[1] = String.valueOf(i1 + (convertMillisintomin(ms)));
						date = date.replace(datearr[3], time[0] + ":" + time[1] + ":" + time[2]);
						Calendar call = Calendar.getInstance();
						call.setTime(c);
						call.set(Calendar.HOUR_OF_DAY, Integer.parseInt(time[0]));
						call.set(Calendar.MINUTE, Integer.parseInt(time[1]));
						call.set(Calendar.SECOND, Integer.parseInt(time[2]));
						call.set(Calendar.MILLISECOND, 0);
						call.setTimeZone(TimeZone.getDefault());

						SimpleDateFormat sdf = new SimpleDateFormat("MMM dd HH:mm:ss z yyyy");

						// SimpleDateFormat sdfn = new SimpleDateFormat("dd-M-yyyy hh:mm:ss");

						String strDate = sdf.format(call.getTime());
						Date c1 = sdf.parse(strDate);
						mess.setDate(c1);
						MapiMessage message5 = MapiMessage.fromMailMessage(mess, d);
						message5.save(System.getProperty("java.io.tmpdir") + "subject" + ".eml",
								SaveOptions.getDefaultEml());
						Date reciveddate = message5.getDeliveryTime();
						long Recivedtime_in_milisecond = reciveddate.getTime();

						try {
							Item ite = new Item(service);
							EmailMessage item = new EmailMessage(service);
//							microsoft.exchange.webservices.data.core.service.item.Appointment ap = null;
							String iCalFileName = System.getProperty("java.io.tmpdir") + "subject" + ".eml";
							FileInputStream fs = new FileInputStream(iCalFileName);
							byte[] bytes = new byte[fs.available()];
							int numBytesToRead = fs.available();
							int numBytesRead = 0;
							while ((numBytesToRead > 0)) {
								int n = fs.read(bytes, numBytesRead, numBytesToRead);
								if ((n == 0)) {
									break;
								}
								numBytesRead = (numBytesRead + n);
								numBytesToRead = (numBytesToRead - n);
							}
							fs.close();
							ite.setMimeContent(new MimeContent("UTF-8", bytes));
							item.setMimeContent(new MimeContent("UTF-8", bytes));
							if (!isDraft) {
								if (subFolder.getDisplayName().contains("Drafts")) {
								} else {
									ExtendedPropertyDefinition PR_MESSAGE_FLAGS_msgflag_read = new ExtendedPropertyDefinition(
											3591,
											microsoft.exchange.webservices.data.core.enumeration.property.MapiPropertyType.Integer);
									item.setExtendedProperty(PR_MESSAGE_FLAGS_msgflag_read, 1);

									ite.setExtendedProperty(PR_MESSAGE_FLAGS_msgflag_read, 1);
								}
							} else {
								if (subFolder.getDisplayName().equalsIgnoreCase("Sent Items")) {
									ExtendedPropertyDefinition PR_MESSAGE_FLAGS_msgflag_read = new ExtendedPropertyDefinition(
											3591,
											microsoft.exchange.webservices.data.core.enumeration.property.MapiPropertyType.Integer);
									item.setExtendedProperty(PR_MESSAGE_FLAGS_msgflag_read, 1);
									ite.setExtendedProperty(PR_MESSAGE_FLAGS_msgflag_read, 1);
								} else {

								}
							}
							if (message1.getMessageClass().equalsIgnoreCase("IPM.Contact")) {

								String contactsubject = "NA";
								try {
									contactsubject = message1.getSubject().toString();
								} catch (Exception e) {
									contactsubject = "NA";
								}

//								if(chckbx_Mail_Filter.isSelected()) {
//									
//									
//									if(checkDateExist(Recivedtime_in_milisecond)) {
//										
//										
//									}
//									
//									
//									
//								}
								MapiContact con = (MapiContact) message1.toMapiMessageItem();

								if (chckbxRemoveDuplicacy.isSelected()) {
									String input = duplicacymapiContact(con);

									if (!listdupliccontact.contains(input)) {

										listdupliccontact.add(input);

										copyContactPSTToOffice(message1, service, contactsubject, folderidx2,
												subFolder.getContainerClass());
										count_destination++;

									}
								} else {

									copyContactPSTToOffice(message1, service, contactsubject, folderidx2,
											subFolder.getContainerClass());
									count_destination++;

								}
								count_destination_total++;

							} else if (message1.getMessageClass().equalsIgnoreCase("IPM.Appointment")
									|| message1.getMessageClass().equalsIgnoreCase("IPM.Schedule.Meeting.Request")) {
								String subject = "non";
								try {
									subject = message1.getSubject().toString();
								} catch (Exception ex) {
									subject = "non";
								}

								MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();

								if (chckbxRemoveDuplicacy.isSelected()) {
									String input = duplicacymapiCal(cal);

									if (!listdupliccal.contains(input)) {

										listdupliccal.add(input);

										copyAppointmentPSTToOffice(message1, service, subject, folderidx2,
												subFolder.getContainerClass());
										count_destination++;

									}
								} else {

									copyAppointmentPSTToOffice(message1, service, subject, folderidx2,
											subFolder.getContainerClass());
									count_destination++;

								}
								count_destination_total++;
							} else if (message1.getMessageClass().equalsIgnoreCase("IPM.StickyNote")) {
								String subject = "non";
								try {
									subject = message1.getSubject().toString().toString();
								} catch (Exception ex) {
									subject = "non";
									// ex.printStackTrace();
								}
								copyStickyNotesPSTToOffice(message1, service, subject, folderidx2,
										subFolder.getContainerClass());
								count_destination++;
								count_destination_total++;
							} else if (message1.getMessageClass().equalsIgnoreCase("IPM.Task")) {
								String subject = "non";
								try {
									subject = message1.getSubject().toString().toString();
								} catch (Exception ex) {
									subject = "non";
									// ex.printStackTrace();
								}
								copyStickyTaskPSTToOffice(message1, service, subject, folderidx2,
										subFolder.getContainerClass());
								count_destination++;
								count_destination_total++;
							} else {
								Date Receiveddate = mess.getDate();

								System.out.println("this message time  : " + Receiveddate);

								long Rec_d_in_mill = Receiveddate.getTime();
								if (chckbx_Mail_Filter.isSelected()) {
									selected = false;
									if (checkDateExist(Rec_d_in_mill)) {
										item.save(folderidx2);
										count_destination++;
									}
								}
								if (chckbxRemoveDuplicacy.isSelected()) {
									selected = false;
									String input = duplicacymapi(message1);
									if (!listduplicacy.contains(input)) {
										System.out.println("Not a duplicate message");
										listduplicacy.add(input);
										item.save(folderidx2);
										count_destination++;
									}
								}
								if (selected) {
									item.save(folderidx2);
									count_destination++;
								}
								count_destination_total++;
							}
						} catch (Exception e) {
							while (!checkInternet()) {
								Progressbar.setText("connecting to server...");
								lbl_progressreport.setText("Please chech your internet");
							}
							e.printStackTrace();
							if (e.getMessage().contains("The request failed. outlook.office365.com")
									|| e.getMessage()
											.contains("The request failed. The request failed. Connection reset")
									|| e.getMessage().contains("Connection reset")
									|| e.getMessage().contains("outlook.office365.com")
									|| e.getMessage()
											.contains("The request failed. java.net.SocketException: Connection reset")
									|| e.getMessage().contains(
											"The request failed. The request failed. No such host is known (outlook.office365.com)")
									|| e.getMessage()
											.contains("The request failed. The request failed. outlook.office365.com")
									|| e.getMessage().contains(
											"The request failed. Network is unreachable: no further information")) {
								connectionHandleoffice_output("");
								i11--;
								mf.logger.warning("Exception during adding message line no 22035: " + e.getMessage());

								e.printStackTrace();
							} else if (e.getMessage().contains(
									"The request failed. The server cannot service this request right now. Try again later")) {

								System.out.println("21252");
								i11--;
								connectionHandleoffice_output("");
								mf.logger.warning("Exception during adding message line no 22044: " + e.getMessage());
								e.printStackTrace();
							}
							mf.logger.warning("Exception during adding message line no 22047 folder name : "
									+ subFolder.getDisplayName() + e.getMessage());

							e.printStackTrace();
						}
						System.out.println(subFolder.getDisplayName() + "  total message save :  " + ++mcount);
						lbl_progressreport.setText("<html>Total Message Saved Count  " + "<b>" + count_destination
								+ " / " + count_destination_total + "</b>" + subFolder.getDisplayName()
								+ "   Extracting messsage " + message.getSubject());
					} catch (Exception ex) {
						if (ex.getMessage().contains("The request failed. outlook.office365.com")
								|| ex.getMessage().contains("The request failed. The request failed. Connection reset")
								|| ex.getMessage().contains("Connection reset")
								|| ex.getMessage().contains("outlook.office365.com")
								|| ex.getMessage()
										.contains("The request failed. java.net.SocketException: Connection reset")
								|| ex.getMessage().contains(
										"The request failed. The request failed. No such host is known (outlook.office365.com)")
								|| ex.getMessage()
										.contains("The request failed. The request failed. outlook.office365.com")
								|| ex.getMessage().contains(
										"The request failed. Network is unreachable: no further information")) {
							connectionHandleoffice_output("");
							ex.printStackTrace();
						} else if (ex.getMessage().contains(
								"The request failed. The server cannot service this request right now. Try again later")) {
							i11--;
							connectionHandleoffice_output("");
							mf.logger.warning("Exception during adding message lie no 22034: " + ex.getMessage());
							ex.printStackTrace();
						}
						ex.printStackTrace();
					}
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	public void copyContactPSTToOffice(MapiMessage mapiMessage, ExchangeService service, String subject,
			FolderId childFolderId, String subFolderClass) throws Exception, Error {
		MapiContact mapiContact = (MapiContact) mapiMessage.toMapiMessageItem();

		Contact contact = new Contact(service);
		try {
			contact.setAssistantName(mapiContact.getProfessionalInfo().getAssistant());
		} catch (Exception e1) {
			// e1.printStackTrace();
		}
		try {
			contact.setBirthday(mapiContact.getEvents().getBirthday());
		} catch (Exception e1) {
			// e1.printStackTrace();
		}
		try {
			contact.setBody(new MessageBody(mapiContact.getBody()));
		} catch (Exception e1) {
			// e1.printStackTrace();
		}
		try {
			contact.setBusinessHomePage(mapiContact.getPersonalInfo().getBusinessHomePage());
		} catch (Exception e1) {
			// e1.printStackTrace();
		}
		try {
			contact.setCompanyName(mapiContact.getProfessionalInfo().getCompanyName());
		} catch (Exception e1) {
			// e1.printStackTrace();
		}
		try {
			byte[] contactPicture = mapiContact.getPhoto().getData();
			contact.setContactPicture(contactPicture);
		} catch (Exception e) {
		} catch (Error e) {
		}
//		contact.setContactPicture("AvinashProfilePicture");
//		contact.setCulture("Culture");
		try {
			contact.setDepartment(mapiContact.getProfessionalInfo().getDepartmentName());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			contact.setDisplayName(mapiContact.getNameInfo().getDisplayName());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			contact.setGivenName(mapiContact.getNameInfo().getGivenName());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			contact.setInitials(mapiContact.getNameInfo().getInitials());
		} catch (Exception e) {
			// e.printStackTrace();
		}
//		contact.setInReplyTo(mapiContact.getEvents().getDisplayName());
//		contact.setIsReminderSet(true);
		try {
			contact.setCompanyName(mapiContact.getProfessionalInfo().getCompanyName());
		} catch (Exception e) {
			// e.printStackTrace();
		}

		contact.setItemClass("IPM.Contact");
		try {
			contact.setJobTitle(mapiContact.getProfessionalInfo().getTitle());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			contact.setManager(mapiContact.getProfessionalInfo().getManagerName());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			contact.setMiddleName(mapiContact.getNameInfo().getMiddleName());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			contact.setMileage(mapiContact.getMileage());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			contact.setNickName(mapiContact.getNameInfo().getNickname());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			contact.setOfficeLocation(mapiContact.getProfessionalInfo().getOfficeLocation());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			contact.setProfession(mapiContact.getProfessionalInfo().getProfession());
		} catch (Exception e) {
			// e.printStackTrace();
		}
//		contact.setReminderDueBy(new Date());
		// System.out.println("Sensitivity : " + mapiContact.getSensitivity());
		contact.setSensitivity(Sensitivity.Normal);
		try {
			contact.setSpouseName(mapiContact.getPersonalInfo().getSpouseName());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			contact.setSubject(mapiContact.getSubject());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		// System.out.println("this is contact subject " + mapiContact.getSubject());
		try {
			contact.setSurname(mapiContact.getNameInfo().getSurname());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			contact.setWeddingAnniversary(mapiContact.getEvents().getWeddingAnniversary());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			if (mapiContact.getElectronicAddresses().getEmail1().getEmailAddress() != null) {
				contact.getEmailAddresses().setEmailAddress(EmailAddressKey.EmailAddress1,
						new EmailAddress(mapiContact.getElectronicAddresses().getEmail1().getEmailAddress()));
			}
			if (mapiContact.getElectronicAddresses().getEmail2().getEmailAddress() != null) {
				contact.getEmailAddresses().setEmailAddress(EmailAddressKey.EmailAddress2,
						new EmailAddress(mapiContact.getElectronicAddresses().getEmail2().getEmailAddress()));
			}
			if (mapiContact.getElectronicAddresses().getEmail3().getEmailAddress() != null) {
				contact.getEmailAddresses().setEmailAddress(EmailAddressKey.EmailAddress3,
						new EmailAddress(mapiContact.getElectronicAddresses().getEmail3().getEmailAddress()));
			}
		} catch (ServiceLocalException e) {
			// e.printStackTrace();
		}
		PhysicalAddressEntry physicaladdress = new PhysicalAddressEntry();
		try {
			physicaladdress.setCity(mapiContact.getPhysicalAddresses().getHomeAddress().getCity());
			physicaladdress.setCountryOrRegion(mapiContact.getPhysicalAddresses().getHomeAddress().getCountry());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			physicaladdress.setPostalCode(mapiContact.getPhysicalAddresses().getHomeAddress().getPostalCode());
			physicaladdress.setState(mapiContact.getPhysicalAddresses().getHomeAddress().getStateOrProvince());
			physicaladdress.setStreet(mapiContact.getPhysicalAddresses().getHomeAddress().getStreet());
		} catch (Exception e) {
			// e.printStackTrace();
		}

		contact.getPhysicalAddresses().setPhysicalAddress(PhysicalAddressKey.Home, physicaladdress);
		physicaladdress = new PhysicalAddressEntry();
		try {
			physicaladdress.setCity(mapiContact.getPhysicalAddresses().getWorkAddress().getCity());
			physicaladdress.setCountryOrRegion(mapiContact.getPhysicalAddresses().getWorkAddress().getCountry());
			physicaladdress.setPostalCode(mapiContact.getPhysicalAddresses().getWorkAddress().getPostalCode());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			physicaladdress.setState(mapiContact.getPhysicalAddresses().getWorkAddress().getStateOrProvince());
			physicaladdress.setStreet(mapiContact.getPhysicalAddresses().getWorkAddress().getStreet());
			contact.getPhysicalAddresses().setPhysicalAddress(PhysicalAddressKey.Business, physicaladdress);
		} catch (ServiceLocalException e) {
			// e.printStackTrace();
		} catch (Exception e) {
			// e.printStackTrace();
		}
		physicaladdress = new PhysicalAddressEntry();
		try {
			physicaladdress.setCity(mapiContact.getPhysicalAddresses().getOtherAddress().getCity());
			physicaladdress.setCountryOrRegion(mapiContact.getPhysicalAddresses().getOtherAddress().getCountry());
			physicaladdress.setPostalCode(mapiContact.getPhysicalAddresses().getOtherAddress().getPostalCode());
			physicaladdress.setState(mapiContact.getPhysicalAddresses().getOtherAddress().getStateOrProvince());
			physicaladdress.setStreet(mapiContact.getPhysicalAddresses().getOtherAddress().getStreet());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			contact.getPhysicalAddresses().setPhysicalAddress(PhysicalAddressKey.Other, physicaladdress);
		} catch (ServiceLocalException e) {
			// e.printStackTrace();
		}
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.AssistantPhone,
				mapiContact.getTelephones().getAssistantTelephoneNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.BusinessFax,
				mapiContact.getElectronicAddresses().getBusinessFax().getFaxNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.BusinessPhone,
				mapiContact.getTelephones().getBusinessTelephoneNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.BusinessPhone2,
				mapiContact.getTelephones().getBusiness2TelephoneNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.Callback,
				mapiContact.getTelephones().getCallbackTelephoneNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.CarPhone,
				mapiContact.getTelephones().getCarTelephoneNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.CompanyMainPhone,
				mapiContact.getTelephones().getCompanyMainTelephoneNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.HomeFax,
				mapiContact.getElectronicAddresses().getHomeFax().getFaxNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.HomePhone,
				mapiContact.getTelephones().getHomeTelephoneNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.HomePhone2,
				mapiContact.getTelephones().getHome2TelephoneNumber());
		try {
			contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.Isdn, mapiContact.getTelephones().getIsdnNumber());
		} catch (ServiceLocalException e) {
			// e.printStackTrace();
		}
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.MobilePhone,
				mapiContact.getTelephones().getMobileTelephoneNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.OtherFax,
				mapiContact.getElectronicAddresses().getPrimaryFax().getFaxNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.OtherTelephone,
				mapiContact.getTelephones().getOtherTelephoneNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.Pager,
				mapiContact.getTelephones().getPagerTelephoneNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.PrimaryPhone,
				mapiContact.getTelephones().getPrimaryTelephoneNumber());
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.RadioPhone,
				mapiContact.getTelephones().getRadioTelephoneNumber());
		try {
			contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.Telex,
					mapiContact.getTelephones().getTelexNumber());
		} catch (ServiceLocalException e) {
			// e.printStackTrace();
		}
		contact.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.TtyTddPhone,
				mapiContact.getTelephones().getTtyTddPhoneNumber());
		contact.getImAddresses().setImAddressKey(ImAddressKey.ImAddress1,
				mapiContact.getPersonalInfo().getInstantMessagingAddress());
		try {
			MapiAttachmentCollection mapiAttachment = mapiContact.getAttachments();
			for (MapiAttachment str : mapiAttachment) {
				str.save(System.getProperty("java.io.tmpdir") + File.separator + str.getFileName());
				contact.getAttachments()
						.addFileAttachment(System.getProperty("java.io.tmpdir") + File.separator + str.getFileName());
			}
		} catch (ServiceLocalException e) {
			// e.printStackTrace();
		}

		if (!subFolderClass.equalsIgnoreCase("IPF.Contact")) {
			// new function to be written
			// get wellknown contactfolder and create custom/locat
			// then create pst file name folder and append contact to it
			// every time this id will be passed will be based on each file
			System.out.println("entered");
			if (!generalFolderMap.containsKey("Contact")) {
				WellKnownFolderName con = WellKnownFolderName.Contacts;
				firstFolderGeneralMaker(locdat, con);
			}
			String pstFileNam = new File(filepath).getName().replaceAll("[\\[\\]]", "") + " " + "Contacts";
			String generalFolderPath = "Contact" + File.separator + pstFileNam;
			if (!generalFolderMap.containsKey(generalFolderPath)) {

				secondGeneralFolderMaker(generalFolderPath, "Contact", pstFileNam, "IPF.Contact");
			}
			FolderId conPar = generalFolderMap.get(generalFolderPath);
			contact.save(conPar);
		} else {
			contact.save(childFolderId);
		}

//		message_count++;
//		lblNewLabel_15.setText(String.valueOf(message_count));
	}

	public void copyStickyNotesPSTToOffice(MapiMessage mapiMessage, ExchangeService service, String subject,
			FolderId childFolderId, String subFolderClass) throws Exception, Error {
		EmailMessage emailMessage = new EmailMessage(service);
		try {
			emailMessage.setSubject(mapiMessage.getSubject());
		} catch (Exception ex) {
			emailMessage.setSubject("");
		} catch (Error ex) {
			emailMessage.setSubject("");
		}
		try {
			MessageBody messageBody = new MessageBody(mapiMessage.getBody());
			emailMessage.setBody(messageBody);
		} catch (Exception ex) {
			emailMessage.setBody(new MessageBody(""));
		} catch (Error ex) {
			emailMessage.setBody(new MessageBody(""));
		}
		emailMessage.setItemClass("IPM.StickyNote");

		emailMessage.save(childFolderId);

	}

	public void copyAppointmentPSTToOffice(MapiMessage mapiMessage, ExchangeService service, String subject,
			FolderId childFolderId, String subFolderClass) throws Exception, Error {
		MapiCalendar mapiCalendar = null;

		try {
			mapiCalendar = (MapiCalendar) mapiMessage.toMapiMessageItem();

			System.out.println(mapiMessage.getSubject());

		} catch (Exception ex) {
			System.out.println("error line no 21672");
//			mf.logger.info(ex.pr);
			// ex.printStackTrace();
		} catch (Error ex) {
//		ex.printStackTrace();
		}

		microsoft.exchange.webservices.data.core.service.item.Appointment appoinment = new microsoft.exchange.webservices.data.core.service.item.Appointment(
				service);
		appoinment.setIsAllDayEvent(true);
		appoinment.setBody(new MessageBody(BodyType.HTML, mapiCalendar.getBodyHtml()));

		try {
			appoinment.setBody(new MessageBody(mapiCalendar.getBody()));
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			appoinment.setStart(mapiCalendar.getStartDate());
			appoinment.setEnd(mapiCalendar.getEndDate());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			appoinment.setSubject(mapiCalendar.getSubject());
		} catch (Exception e) {
			// e.printStackTrace();
		}
		try {
			MapiAttachmentCollection mapiAttachmeentCollection = mapiCalendar.getAttachments();
			for (MapiAttachment attachment : mapiAttachmeentCollection) {
				attachment.save(new FileOutputStream(
						new File(System.getProperty("java.io.tmpdir") + File.separator + attachment.getFileName())));
				appoinment.getAttachments().addFileAttachment(
						System.getProperty("java.io.tmpdir") + File.separator + attachment.getFileName());
			}
		} catch (FileNotFoundException e) {
			// e.printStackTrace();
		} catch (ServiceLocalException e) {
			// e.printStackTrace();
		}
		try {
			appoinment.setLocation(mapiCalendar.getLocation());
		} catch (Exception e) {
			// e.printStackTrace();
		}

		try {
			StringList stringList = new StringList();
			String[] categori = mapiCalendar.getCategories();
			for (String cataeg : categori) {
				stringList.add(cataeg);
			}
			appoinment.setCategories(stringList);
		} catch (Exception ex) {
		} catch (Error ex) {
		}
//		if(!subFolderClass.equalsIgnoreCase("IPF.Appointment")) {
//			// new function to be written 
//			if(!generalFolderMap.containsKey("Appointment")) {
//				WellKnownFolderName calc=WellKnownFolderName.Calendar;
//				firstFolderGeneralMaker(locdat,calc);
//			}
//			String pstFileNam=new File(filepath).getName().replaceAll("[\\[\\]]", "")+" "+"Calenders";
//			String generalFolderPath="Appointment"+File.separator+pstFileNam;
//			if(!generalFolderMap.containsKey(generalFolderPath)) {
//				
//				secondGeneralFolderMaker(generalFolderPath,"Appointment",pstFileNam,"IPF.Appointment");
//			}
//			FolderId calcPar=generalFolderMap.get(generalFolderPath);
//			appoinment.save(calcPar, SendInvitationsMode.SendToNone);
//		}else {
//			appoinment.save(childFolderId, SendInvitationsMode.SendToNone);
//		}
//	

		if (!subFolderClass.equalsIgnoreCase("IPF.Appointment")) {
			// new function to be written
			if (!generalFolderMap.containsKey("Appointment")) {
				WellKnownFolderName calc = WellKnownFolderName.Calendar;
				firstFolderGeneralMaker(locdat, calc);
			}
			String pstFileNam = new File(filepath).getName().replaceAll("[\\[\\]]", "") + " " + "Calenders";
			String generalFolderPath = "Appointment" + File.separator + pstFileNam;
			if (!generalFolderMap.containsKey(generalFolderPath)) {

				secondGeneralFolderMaker(generalFolderPath, "Appointment", pstFileNam, "IPF.Appointment");
			}
			FolderId calcPar = generalFolderMap.get(generalFolderPath);

			try {
				appoinment.save(calcPar, SendInvitationsMode.SendToNone);
			} catch (Exception e) {
				if (e.getMessage().contains("The request failed schema validation:")) {
					appoinment.setSubject("had unicode codes");
					appoinment.save(calcPar, SendInvitationsMode.SendToNone);
				} else {
					throw e;
				}
			}
		} else {
			System.out.println(mapiMessage.getMessageClass() + " message class" + subFolderClass + "  folder class");
			try {
				appoinment.save(childFolderId, SendInvitationsMode.SendToNone);
			} catch (Exception e) {
				if (e.getMessage().contains("The request failed schema validation:")) {
					appoinment.setSubject("had unicode codes");
					appoinment.save(childFolderId, SendInvitationsMode.SendToNone);
				} else {
					throw e;
				}
			}
		}

//		MapiCalendar mapiCalendar = null;
//		try {
//			mapiCalendar = (MapiCalendar) mapiMessage
//					.toMapiMessageItem();
//		} catch (Exception ex) {
//		}
//		MapiCalendarAttendees attendee = mapiCalendar.getAttendees();
//		MapiRecipientCollection attendeeRecipent = attendee
//				.getAppointmentRecipients();
//		microsoft.exchange.webservices.data.core.service.item.Appointment appoinment = new microsoft.exchange.webservices.data.core.service.item.Appointment (service);
//		appoinment.setSubject(mapiCalendar.getSubject());
//		appoinment.setBody(new MessageBody(BodyType.HTML,
//				mapiCalendar.getBodyHtml()));
//		appoinment.setIsAllDayEvent(mapiCalendar.isAllDay());
//		appoinment.setAllowNewTimeProposal(
//				mapiCalendar.getAppointmentCounterProposal());
////		TimeZoneDefinition timezonedefination = new TimeZoneDefinition();
////	
////		timezonedefination = new TimeZoneDefinition();
////		timezonedefination
////				.setId(mapiCalendar.getEndDateTimeZone().getKeyName());
//		
//		appoinment.setStart(mapiCalendar.getStartDate());
//		appoinment.setEnd(mapiCalendar.getEndDate());
//		appoinment.setLocation(mapiCalendar.getLocation());
//		MapiAttachmentCollection mapiAttachmeentCollection = mapiCalendar
//				.getAttachments();
//		for (MapiAttachment attachment : mapiAttachmeentCollection) {
//			attachment.save(new FileOutputStream(
//					new File(System.getProperty("java.io.tmpdir")
//							+ File.separator
//							+ attachment.getFileName())));
//			appoinment.getAttachments().addFileAttachment(
//					System.getProperty("java.io.tmpdir")
//							+ File.separator
//							+ attachment.getFileName());
//		}
//		appoinment.setLocation(mapiCalendar.getLocation());
//		MapiCalendarEventRecurrence mapiCalendarEventRecurrence = mapiCalendar
//				.getRecurrence();
//		if (mapiCalendarEventRecurrence
//				.getRecurrencePattern() != null) {
//			System.out.println(
//					"+++++++++++++++++RECCURENCE+++++++++++++++++");
//			System.out.println("start :"
//					+ mapiCalendarEventRecurrence.getClipEnd());
//			System.out.println("end : "
//					+ mapiCalendarEventRecurrence.getClipStart());
//			System.out.println(
//					"getPatternType : " + mapiCalendarEventRecurrence
//							.getRecurrencePattern().getPatternType());
//			System.out.println("reccurence getOccurrenceCount : "
//					+ mapiCalendarEventRecurrence.getRecurrencePattern()
//							.getOccurrenceCount());
//			System.out
//					.println("timeZone :" + mapiCalendarEventRecurrence
//							.getTimeZoneStruct().getKeyName());
//			appoinment.setRecurrence(
//					new Recurrence.DailyPattern(new Date(), 5));
//			if (mapiCalendarEventRecurrence.getRecurrencePattern()
//					.getPatternType() == 1) {
//			} else if (mapiCalendarEventRecurrence
//					.getRecurrencePattern().getPatternType() == 2) {
//			} else if (mapiCalendarEventRecurrence
//					.getRecurrencePattern().getPatternType() == 3) {
//			} else {
//			}
//		}
//		try {
//			StringList stringList = new StringList();
//			String[] categori = mapiCalendar.getCategories();
//			for (String cataeg : categori) {
//				stringList.add(cataeg);
//			}
//			appoinment.setCategories(stringList);
//		} catch (Exception ex) {
//		} catch (Error ex) {
//		}
//		if (attendeeRecipent.size() == 0) {
//			appoinment.save(childFolderId,
//					SendInvitationsMode.SendToNone);
//		} else {
//			for (MapiRecipient attendeId : attendeeRecipent) {
//				try {
//					if (attendeId.getRecipientType() == 2) {
//						appoinment.getOptionalAttendees()
//								.add(attendeId.getEmailAddress());
//					} else {
//						appoinment.getRequiredAttendees()
//								.add(attendeId.getEmailAddress());
//					}
//				} catch (Exception ex) {
//				} catch (Error ex) {
//				}
//			}
//			try {
//				appoinment.save(childFolderId);
//			} catch (Exception e) {
//				e.printStackTrace();
//			}
//		}

	}

	public void copyStickyTaskPSTToOffice(MapiMessage mapiMessageWithoutAnyConversion, ExchangeService service,
			String subject, FolderId childFolderId, String subFolderClass) throws Exception {
		MapiTask mapiTask = (MapiTask) mapiMessageWithoutAnyConversion.toMapiMessageItem();
		microsoft.exchange.webservices.data.core.service.item.Task task = new microsoft.exchange.webservices.data.core.service.item.Task(
				service);
		task.setSubject(subject);
//		MessageBody messageBody = null;

		try {
			MessageBody messageBody = new MessageBody(mapiMessageWithoutAnyConversion.getBody());
			task.setBody(messageBody);
		} catch (Exception ex) {
			task.setBody(new MessageBody(""));
		} catch (Error ex) {
			task.setBody(new MessageBody(""));
		}
//		try {
//			messageBody = new MessageBody(BodyType.HTML, mapiTask.getBodyHtml());
//			System.out.println("HTML code : " + mapiTask.getBodyHtml());
//			System.out.println("S code : " + mapiTask.getBody());
//		} catch (Exception e) {
//			messageBody = new MessageBody(BodyType.HTML, "");
//		} catch (Error e) {
//			messageBody = new MessageBody(BodyType.HTML, "");
//		}
//		task.setBody(messageBody);
		StringList stringList;
		String[] categories;
		try {
			stringList = new StringList();
			categories = mapiTask.getCategories();
			for (String categ : categories) {
				stringList.add(categ);
			}
			task.setCategories(stringList);
		} catch (Exception e10) {

		}
		try {
			stringList = new StringList();
			categories = mapiTask.getCompanies();
			for (String categ : categories) {
				stringList.add(categ);
			}
			task.setCompanies(stringList);
		} catch (Exception e10) {

		}
		try {
//		task.setCompleteDate(mapiTask.getDateCompleted());
		} catch (Exception e9) {

		}
//	task.setContacts(stringList);
//	task.setCulture();
		try {
			task.setDueDate(mapiTask.getDueDate());
		} catch (Exception e8) {
			e8.printStackTrace();
		}
//	task.setImportance(null);
//	task.setInReplyTo(ParentFolder);
		try {
			task.setIsReminderSet(mapiTask.getReminderSet());
		} catch (Exception e7) {
			e7.printStackTrace();
		}
		try {
			task.setMileage(mapiTask.getMileage());
		} catch (Exception e6) {
			e6.printStackTrace();
		}
		try {
			task.setPercentComplete((double) mapiTask.getPercentComplete());
		} catch (Exception e5) {
			e5.printStackTrace();
		}
		try {
			task.setRecurrence(null);
		} catch (Exception e4) {
			e4.printStackTrace();
		}
		try {
			task.setSensitivity(Sensitivity.Confidential);
		} catch (Exception e3) {
			task.setSensitivity(Sensitivity.Normal);
		}
		try {
			task.setStartDate(mapiTask.getStartDate());
		} catch (Exception e2) {
			task.setStartDate(null);
		}
		try {
			task.setStatus(TaskStatus.Completed);
		} catch (Exception e1) {
			task.setStatus(null);
		}
		try {
			task.setSubject(mapiTask.getSubject());
		} catch (Exception e) {
			task.setSubject("");
		}
//	task.setTotalWork(null)
//	task.setInReplyTo(ParentFolder);
		if (!subFolderClass.equalsIgnoreCase("IPF.Task")) {
			// new function to be written
			if (!generalFolderMap.containsKey("Task")) {
				WellKnownFolderName tas = WellKnownFolderName.Tasks;
				firstFolderGeneralMaker(locdat, tas);
			}
			String pstFileNam = new File(filepath).getName().replaceAll("[\\[\\]]", "") + " " + "Tasks";
			String generalFolderPath = "Task" + File.separator + pstFileNam;
			if (!generalFolderMap.containsKey(generalFolderPath)) {

				secondGeneralFolderMaker(generalFolderPath, "Task", pstFileNam, "IPF.Task");
			}
			FolderId taskPar = generalFolderMap.get(generalFolderPath);
			task.save(taskPar);
		} else {
			task.save(childFolderId);
		}

	}

	public void connectionHandleoffice_output(String id) throws Exception, Error {

		if (filetype.equals("OFFICE 365")) {

			service.close();
			service = connectRefreshOutput(username_p3.trim(), password_p3.trim());
//			FolderView folderview1 = new FolderView(Integer.MAX_VALUE);
//			findfoldersresults = service_output.findFolders(WellKnownFolderName.MsgFolderRoot, folderview1);
		}
	}

	public static ExchangeService loginWithRefreshTokenEWS(String username) {
		try {
			String refreshToken = ews.getRefreshToken();
			service = ews.loginRefreshTokenEWS(username, refreshToken);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return service;
	}

	public ExchangeService connectRefreshOutput(String email, String password) throws Exception, Error {
		boolean connection = true;
//		ExchangeService servicee = null;
		label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));
		while (connection == true) {
			if (stop) {
				break;
			}
			lbl_progressreport.setText("Reconnecting to Office 365 server...");
			try {
//				service = new ExchangeService();
//				service.setUrl(new URI("https://outlook.office365.com/EWS/Exchange.asmx"));
				if (modern_Authentication.isSelected()) {

					service = ows.loginRefreshTokenEWS(username_p3, ows.getRefreshToken());

//					BearerTokenCredentials credentials1 = new BearerTokenCredentials(refresh_Token.refreshoutput());
//					service.setCredentials(credentials1);
				} else {
					ExchangeCredentials credentials = new WebCredentials(email, password);
					service.setCredentials(credentials);
				}
				service.setTraceEnabled(true);
				service.findFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(1));
				connection = false;
			} catch (Exception ex) {
				mf.logger.warning("Exception during reconnecting : " + ex.getMessage());
				System.out.println("connecting to server.......... ");

				connection = true;
			} catch (Error ex) {
				mf.logger.warning("Exception during reconnecting : " + ex.getMessage());
				System.out.println("connecting to server.......... ");
				connection = true;
			}
		}
		label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));
		lbl_progressreport.setText("Connection extablished Retriving Messasge");
		return service;
	}

	public boolean checkDateExist(Long emailMilliseconds) {
		Calendar calendarstartdate = dateChooser_mail_fromdate.getCalendar();
		calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
		calendarstartdate.set(Calendar.MINUTE, 00);
		calendarstartdate.set(Calendar.SECOND, 00);
		Calendar calendarenddate = dateChooser_mail_tilldate.getCalendar();
		calendarenddate.set(Calendar.HOUR_OF_DAY, 23);
		calendarenddate.set(Calendar.MINUTE, 59);
		calendarenddate.set(Calendar.SECOND, 59);
		Long startDateMillisecond = calendarstartdate.getTimeInMillis();
		Long endateMillisecond = calendarenddate.getTimeInMillis();
		if (emailMilliseconds >= startDateMillisecond && emailMilliseconds <= endateMillisecond) {
			return true;
		}
		return false;
	}

	public void thirdFolderMaker(HashMap<String, FolderId> map, FolderInfo folderInfo2, FolderId secndFoldId) {

		try {
			if (stop) {
				return;
			}
			Folder firstTopOfFolder = new Folder(service);
			firstTopOfFolder.setDisplayName(folderInfo2.getDisplayName().replaceAll("[\\[\\]]", "").trim());
			firstTopOfFolder.save(secndFoldId);
			FolderId base = firstTopOfFolder.getId();
			System.out.println(folderInfo2.getDisplayName() + "  folderInfo2.getDisplayName()");
			map.put(folderInfo2.getDisplayName(), base);

		} catch (Exception e) {

			while (!checkInternet()) {
				Progressbar.setText("connecting to server...");
				lbl_progressreport.setText("Please chech your internet");
			}
			mf.logger.info("The request failed. The server cannot service this request right now. Try again later"
					+ e.getMessage());

			if (e.getMessage().contains(
					"The request failed. The server cannot service this request right now. Try again later")) {

				if (stop) {
					return;
				}
				try {
					connectionHandleoffice_output("");
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (Error e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

				thirdFolderMaker(map, folderInfo2, secndFoldId);
				e.printStackTrace();
			} else if (e.getMessage().contains("The request failed. outlook.office365.com")
					|| e.getMessage().contains("The request failed. The request failed. Connection reset")
					|| e.getMessage().contains("Connection reset") || e.getMessage().contains("outlook.office365.com")
					|| e.getMessage().contains("The request failed. java.net.SocketException: Connection reset")
					|| e.getMessage().contains(
							"The request failed. The request failed. No such host is known (outlook.office365.com)")
					|| e.getMessage().contains("The request failed. The request failed. outlook.office365.com")
					|| e.getMessage().contains("The request failed. Network is unreachable: no further information")) {

				try {
					connectionHandleoffice_output("");
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (Error e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				if (stop) {
					return;
				}

				thirdFolderMaker(map, folderInfo2, secndFoldId);
				e.printStackTrace();

			}
			e.printStackTrace();
		}
	}

	public Folder secondFolderMaker(String filepath) {
		Folder nameOfFile = null;
		try {

			if (stop) {
				return null;
			}
			nameOfFile = new Folder(service);
			nameOfFile.setDisplayName(new File(filepath).getName().replaceAll("[\\[\\]]", "").trim());
			nameOfFile.save(rootfolderid);

			return nameOfFile;

		} catch (Exception e) {

			while (!checkInternet()) {
				Progressbar.setText("connecting to server...");
				lbl_progressreport.setText("Please chech your internet");
			}

			mf.logger.info("The request failed. The server cannot service this request right now. Try again later"
					+ e.getMessage());

			if (e.getMessage().contains(
					"The request failed. The server cannot service this request right now. Try again later")) {

				if (stop) {
					return null;
				}
				try {
					connectionHandleoffice_output("");
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (Error e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				e.printStackTrace();
				return secondFolderMaker(filepath);

			} else if (e.getMessage().contains("The request failed. outlook.office365.com")
					|| e.getMessage().contains("The request failed. The request failed. Connection reset")
					|| e.getMessage().contains("Connection reset") || e.getMessage().contains("outlook.office365.com")
					|| e.getMessage().contains("The request failed. java.net.SocketException: Connection reset")
					|| e.getMessage().contains(
							"The request failed. The request failed. No such host is known (outlook.office365.com)")
					|| e.getMessage().contains("The request failed. The request failed. outlook.office365.com")
					|| e.getMessage().contains("The request failed. Network is unreachable: no further information")) {

				try {
					connectionHandleoffice_output("");
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (Error e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				if (stop) {
					return null;
				}
				e.printStackTrace();
				return secondFolderMaker(filepath);

			}
			e.printStackTrace();
		}
		return nameOfFile;
	}

	public void firstFolderMaker(String locdat) {
		Folder folderroot = null;
		try {

			if (stop) {
				return;
			}
			System.out.println(locdat + "  exist");
			folderroot = new Folder(service);
			folderroot.setDisplayName(locdat);

			if (filterselected.equals("mailboxsel")) {
				folderroot.save(WellKnownFolderName.MsgFolderRoot);
				rootfolderid = folderroot.getId();
			} else if (filterselected.equals("archivesel")) {
				folderroot.save(WellKnownFolderName.ArchiveMsgFolderRoot);
				rootfolderid = folderroot.getId();
			} else if (filterselected.equals("publicfoldersel")) {
				folderroot.save(WellKnownFolderName.PublicFoldersRoot);
				rootfolderid = folderroot.getId();
			}

		} catch (Exception e1) {
			StringWriter sw = new StringWriter();
			e1.printStackTrace(new PrintWriter(sw));
			String exceptionAsString = sw.toString();

			if (e1.getMessage().contains("Access is denied")) {
//				String warn = "You Have not access to Public folder ";
				JOptionPane.showMessageDialog(main_multiplefile.this,
						"You Don't have Access to Public folder.Please first Allow the Access of Public folder From Setting,While Right Now Data will be migrated to MailBox",
						messageboxtitle, JOptionPane.ERROR_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));
				try {
					folderroot.save(WellKnownFolderName.MsgFolderRoot);
					rootfolderid = folderroot.getId();
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				e1.printStackTrace();
				return;
			} else if (e1.getMessage().contains("The request failed. outlook.office365.com")
					|| e1.getMessage().contains("The request failed. The request failed. Connection reset")
					|| e1.getMessage().contains("Connection reset") || e1.getMessage().contains("outlook.office365.com")
					|| e1.getMessage().contains("The request failed. java.net.SocketException: Connection reset")
					|| e1.getMessage().contains(
							"The request failed. The request failed. No such host is known (outlook.office365.com)")
					|| e1.getMessage().contains("The request failed. The request failed. outlook.office365.com")
					|| e1.getMessage().contains("The request failed. Network is unreachable: no further information")) {
				e1.printStackTrace();
//				StringWriter sw = new StringWriter();
//				PrintWriter pw = new PrintWriter(sw);
//				e1.printStackTrace(pw);
				mf.logger.warning(exceptionAsString);

				try {
					connectionHandleoffice_output("");
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (Error e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				firstFolderMaker(locdat);
				if (stop) {
					return;
				}
			} else if (e1.getMessage().contains("A folder with the specified name already exists.")) {
				JOptionPane.showMessageDialog(main_multiplefile.this,
						"This Folder Name is already exist please change custom Folder name and try again ",
						messageboxtitle, JOptionPane.ERROR_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));
				textField_customfolder.setEditable(true);
				except = true;
				e1.printStackTrace();
				return;
			} else if (exceptionAsString.contains(" The specified folder could not be found in the store.")) {
				try {
					JOptionPane.showMessageDialog(main_multiplefile.this,
							"You Don't have Access to Archive folder.Please first Allow the Access of Archive Folder From Setting,While Right Now Data will be migrated to MailBox",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));
					folderroot.save(WellKnownFolderName.MsgFolderRoot);
					rootfolderid = folderroot.getId();
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

			} else {
				e1.printStackTrace();
			}
		}
	}

	public void firstFolderGeneralMaker(String locdats, WellKnownFolderName foundation) {
		Folder folderroot1 = null;
		try {

			if (stop) {
				return;
			}
			folderroot1 = new Folder(service);
			folderroot1.setDisplayName(locdats);
			System.out.println(locdats + " exists");

			if (foundation == WellKnownFolderName.Contacts) {
				System.out.println(locdats + " exists  Contact");

			} else if (foundation == WellKnownFolderName.Calendar) {
				System.out.println(locdats + " exists  Appointment");

			} else if (foundation == WellKnownFolderName.Tasks) {
				System.out.println(locdats + " exists  Task");
			}

			folderroot1.save(foundation);
			FolderId rootfoldergeneralId = folderroot1.getId();

			if (foundation == WellKnownFolderName.Contacts) {

				generalFolderMap.put("Contact", rootfoldergeneralId);
			} else if (foundation == WellKnownFolderName.Calendar) {

				generalFolderMap.put("Appointment", rootfoldergeneralId);
			} else if (foundation == WellKnownFolderName.Tasks) {

				generalFolderMap.put("Task", rootfoldergeneralId);
			}

		} catch (Exception e1) {
			StringWriter sw = new StringWriter();
			e1.printStackTrace(new PrintWriter(sw));
			String exceptionAsString = sw.toString();

			if (e1.getMessage().contains("Access is denied")) {
//				String warn = "You Have not access to Public folder ";
				JOptionPane.showMessageDialog(main_multiplefile.this,
						"You Don't have Access to Public folder.Please first Allow the Access of Public folder From Setting,While Right Now Data will be migrated to MailBox",
						messageboxtitle, JOptionPane.ERROR_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));

				e1.printStackTrace();
				return;
			} else if (e1.getMessage().contains("The request failed. outlook.office365.com")
					|| e1.getMessage().contains("The request failed. The request failed. Connection reset")
					|| e1.getMessage().contains("Connection reset") || e1.getMessage().contains("outlook.office365.com")
					|| e1.getMessage().contains("The request failed. java.net.SocketException: Connection reset")
					|| e1.getMessage().contains(
							"The request failed. The request failed. No such host is known (outlook.office365.com)")
					|| e1.getMessage().contains("The request failed. The request failed. outlook.office365.com")
					|| e1.getMessage().contains("The request failed. Network is unreachable: no further information")) {
				e1.printStackTrace();
//				StringWriter sw = new StringWriter();
//				PrintWriter pw = new PrintWriter(sw);
//				e1.printStackTrace(pw);
				mf.logger.warning(exceptionAsString);

				try {
					connectionHandleoffice_output("");
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (Error e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				firstFolderGeneralMaker(locdat, foundation);
				if (stop) {
					return;
				}
			} else if (e1.getMessage().contains("A folder with the specified name already exists.")) {
				JOptionPane.showMessageDialog(main_multiplefile.this,
						"This Folder Name is already exist please change custom Folder name and try again ",
						messageboxtitle, JOptionPane.ERROR_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));
				textField_customfolder.setEditable(true);
				except = true;
				e1.printStackTrace();
				return;
			} else if (exceptionAsString.contains(" The specified folder could not be found in the store.")) {
//				try {
//					JOptionPane.showMessageDialog(main_multiplefile.this,
//							"You Don't have Access to Archive folder.Please first Allow the Access of Archive Folder From Setting,While Right Now Data will be migrated to MailBox",
//							messageboxtitle, JOptionPane.ERROR_MESSAGE,
//							new ImageIcon(Main_Frame.class.getResource("/information.png")));
//					folderroot.save(WellKnownFolderName.MsgFolderRoot);
//					rootfolderid = folderroot.getId();
//				} catch (Exception e) {
//					// TODO Auto-generated catch block
//					e.printStackTrace();
//				}

			} else {
				e1.printStackTrace();
			}
		}
	}

	public void secondGeneralFolderMaker(String savePath, String ParentGeneral, String name, String folClass) {

		try {
			if (stop) {
				return;
			}
			Folder firstsTopOfFolder = new Folder(service);
			firstsTopOfFolder.setDisplayName(name);
			firstsTopOfFolder.setFolderClass(folClass);
			FolderId ancestor = generalFolderMap.get(ParentGeneral);
			System.out.println(ParentGeneral + " ParentGeneral");
			firstsTopOfFolder.save(ancestor);
			FolderId base = firstsTopOfFolder.getId();
			generalFolderMap.put(savePath, base);

		} catch (Exception e) {

			while (!checkInternet()) {
				Progressbar.setText("connecting to server...");
				lbl_progressreport.setText("Please chech your internet");
			}
			mf.logger.info("The request failed. The server cannot service this request right now. Try again later"
					+ e.getMessage());

			if (e.getMessage().contains(
					"The request failed. The server cannot service this request right now. Try again later")) {

				if (stop) {
					return;
				}
				try {
					connectionHandleoffice_output("");
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (Error e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

//					 thirdFolderMaker(map,folderInfo2,secndFoldId);
				e.printStackTrace();
			} else if (e.getMessage().contains("The request failed. outlook.office365.com")
					|| e.getMessage().contains("The request failed. The request failed. Connection reset")
					|| e.getMessage().contains("Connection reset") || e.getMessage().contains("outlook.office365.com")
					|| e.getMessage().contains("The request failed. java.net.SocketException: Connection reset")
					|| e.getMessage().contains(
							"The request failed. The request failed. No such host is known (outlook.office365.com)")
					|| e.getMessage().contains("The request failed. The request failed. outlook.office365.com")
					|| e.getMessage().contains("The request failed. Network is unreachable: no further information")) {

				try {
					connectionHandleoffice_output("");
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (Error e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				if (stop) {
					return;
				}

				secondGeneralFolderMaker(savePath, ParentGeneral, name, folClass);
				e.printStackTrace();

			}
			e.printStackTrace();
		}

	}

	public void Buttonclick(String btnName, boolean enter) {

		if (btnName.equals("btn_previous")) {
			btn_previous_p2.setGradientColor1(new Color(70, 130, 180));
			btn_previous_p2.setGradientColor2(new Color(70, 130, 180));
			btn_previous_p2.setForeground(new Color(255, 255, 255));

			btn_previous_p2.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseEntered(MouseEvent e) {
//					
				}

				@Override
				public void mouseExited(MouseEvent e) {

					if (btn_previous_p2.isEnabled() & fir) {
						btn_previous_p2.setGradientColor1(new Color(70, 130, 180));
						btn_previous_p2.setGradientColor2(new Color(70, 130, 180));
						btn_previous_p2.setForeground(new Color(255, 255, 255));

						second = false;
						third = false;
						fourth = false;
						fifth = false;
					}

				}

			});

			btnNotClick(btnNewButton_2);
			btnNotClick(btn_converter_1);
			btnNotClick(btn_next_pane2);
			btnNotClick(btn_Next);

		}
		if (btnName.equals("btn_Next")) {
			btn_Next.setGradientColor1(new Color(70, 130, 180));
			btn_Next.setGradientColor2(new Color(70, 130, 180));
			btn_Next.setForeground(new Color(255, 255, 255));

			btn_Next.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseEntered(MouseEvent e) {
//					
				}

				@Override
				public void mouseExited(MouseEvent e) {

					if (btn_Next.isEnabled() & second) {
						btn_Next.setGradientColor1(new Color(70, 130, 180));
						btn_Next.setGradientColor2(new Color(70, 130, 180));
						btn_Next.setForeground(new Color(255, 255, 255));

						fir = false;
						third = false;
						fourth = false;
						fifth = false;

					}

				}

			});

			btnNotClick(btnNewButton_2);
			btnNotClick(btn_converter_1);
			btnNotClick(btn_next_pane2);
			btnNotClick(btn_previous_p2);
		}
		if (btnName.equals("btn_next_pane2")) {
			btn_next_pane2.setGradientColor1(new Color(70, 130, 180));
			btn_next_pane2.setGradientColor2(new Color(70, 130, 180));
			btn_next_pane2.setForeground(new Color(255, 255, 255));

			btn_next_pane2.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseEntered(MouseEvent e) {
//					
				}

				@Override
				public void mouseExited(MouseEvent e) {

					if (btn_next_pane2.isEnabled() & third) {
						btn_next_pane2.setGradientColor1(new Color(70, 130, 180));
						btn_next_pane2.setGradientColor2(new Color(70, 130, 180));
						btn_next_pane2.setForeground(new Color(255, 255, 255));

						fir = false;
						second = false;
						fourth = false;
						fifth = false;
					}

				}

			});

			btnNotClick(btnNewButton_2);
			btnNotClick(btn_converter_1);
			btnNotClick(btn_Next);
			btnNotClick(btn_previous_p2);
		}

		if (btnName.equals("btn_converter_1")) {
			btn_converter_1.setGradientColor1(new Color(70, 130, 180));
			btn_converter_1.setGradientColor2(new Color(70, 130, 180));
			btn_converter_1.setForeground(new Color(255, 255, 255));

			btn_converter_1.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseEntered(MouseEvent e) {
//					
				}

				@Override
				public void mouseExited(MouseEvent e) {

					if (btn_converter_1.isEnabled() & fourth) {
						btn_converter_1.setGradientColor1(new Color(70, 130, 180));
						btn_converter_1.setGradientColor2(new Color(70, 130, 180));
						btn_converter_1.setForeground(new Color(255, 255, 255));

						fir = false;
						second = false;
						third = false;
						fifth = false;

					}

				}

			});

			btnNotClick(btnNewButton_2);
			btnNotClick(btn_next_pane2);
			btnNotClick(btn_Next);
			btnNotClick(btn_previous_p2);
		}
		if (btnName.equals("btnNewButton_2")) {
			btnNewButton_2.setGradientColor1(new Color(70, 130, 180));
			btnNewButton_2.setGradientColor2(new Color(70, 130, 180));
			btnNewButton_2.setForeground(new Color(255, 255, 255));

			btnNewButton_2.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseEntered(MouseEvent e) {
//					
				}

				@Override
				public void mouseExited(MouseEvent e) {

					if (btnNewButton_2.isEnabled() & fifth) {
						btnNewButton_2.setGradientColor1(new Color(70, 130, 180));
						btnNewButton_2.setGradientColor2(new Color(70, 130, 180));
						btnNewButton_2.setForeground(new Color(255, 255, 255));

						fir = false;
						second = false;
						third = false;
						fourth = false;

					}

				}

			});

			btnNotClick(btn_converter_1);
			btnNotClick(btn_next_pane2);
			btnNotClick(btn_Next);
			btnNotClick(btn_previous_p2);
		}

	}

	public void btnNotClick(GradientButton notsellected) {

		notsellected.setGradientColor1(new Color(255, 255, 255));
		notsellected.setGradientColor2(new Color(255, 255, 255));
		notsellected.setForeground(new Color(80, 80, 80));

	}
	
	
	
	
	
	public void selecteddestination() {
		try {
			chckbxCustomFolderName.setSelected(false);
			if (chckbxCustomFolderName.isSelected()) {
				textField_customfolder.setEnabled(true);
				textField_customfolder.setEditable(true);
			} else {

				textField_customfolder.setText("");
				textField_customfolder.setEnabled(false);
				textField_customfolder.setEditable(false);
			}

			panel_3_.setVisible(false);

			panel_3_2.setVisible(false);

			panel_3_1_2.setVisible(false);

			panel_3_1_2.setVisible(false);
			lbl_splitpst.setVisible(false);
			chckbx_splitpst.setSelected(false);
			chckbx_splitpst.setVisible(false);
			panel_3_1_2_1.setVisible(false);
			chckbxSavePdfAttachment.setVisible(false);
			label_15.setVisible(false);
			textField_domain_name_p3.setText("");
			output = false;
			tf_Destination_Location.setText(System.getProperty("user.home") + File.separator + "Desktop");
			textField_username_p3.setText("");
			chckbxSaveInSame.setSelected(false);
			try {
				passwordField_p3.setText("");
				lbl_progressreport.setText("");
				btn_converter_1.setEnabled(false);
				lbl_splitpst.setVisible(false);
				chckbx_splitpst.setSelected(false);
				panel_progress.setVisible(true);
				tf_portNo_p3.setVisible(false);
				lblPortNo.setVisible(false);
				comboBox.setVisible(false);
				btn_signout_p3.setVisible(false);
				chckbx_convert_pdf_to_pdf.setVisible(false);
				label_pdf_to_pdf.setVisible(false);
				lblMakeSureYou.setVisible(true);
				lblEnableImap_p3.setVisible(true);
				lblTurnOffTwo_p3.setVisible(true);
				label_16.setVisible(false);
				chckbxSaveMboxIn.setVisible(false);
				chckbxRestoreToDefault.setVisible(false);
				panel_5.setVisible(false);
				panel_8.setVisible(false);
				mailbox.setVisible(false);
				archive.setVisible(false);
				publicfolder.setVisible(false);
//				if (arg0.getSource() == comboBox_fileDestination_type) {
//
//					JComboBox cb = (JComboBox) arg0.getSource();
//
//					filetype = (String) cb.getSelectedItem();
//					if (cb.getSelectedItem() == null) {
//						return;
//					}
//				}
			} catch (Exception e) {
				System.out.println("here we riched 2582");
				// TODO Auto-generated catch block
				// e.printStackTrace();
			}

//if(!cb.getSelectedItem()==null) {
			try {
				if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")
						|| filetype.equalsIgnoreCase("YAHOO MAIL") || filetype.equalsIgnoreCase("Icloud")
						|| filetype.equalsIgnoreCase("GoDaddy email")
						|| filetype.equalsIgnoreCase("Hostgator email")
						|| filetype.equalsIgnoreCase("Amazon WorkMail")
						|| filetype.equalsIgnoreCase("OFFICE 365") || filetype.equalsIgnoreCase("AOL")
						|| filetype.equalsIgnoreCase("Live Exchange")
						|| filetype.equalsIgnoreCase("Yandex Mail") || filetype.equalsIgnoreCase("Zoho Mail")
						|| filetype.equalsIgnoreCase("HOTMAIL") || filetype.equalsIgnoreCase("IMAP")) {
					lblEnableImap_p3.setText("<HTML><U>To Enable IMAP</U><HTML>");
					lbl_splitpst.setVisible(false);
					chckbx_splitpst.setSelected(false);
					chckbxSaveInSame.setVisible(false);
					label_13.setVisible(false);
					lblEnableImap_p3.setVisible(false);
					lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
							+ " account , you'll need to generate and use an app password.</U></HTML>");
					lblMakeSureYou.setText("Please  Click on The Link");
					lblNewLabel_5.setVisible(true);

					if (!filetype.equalsIgnoreCase("OFFICE 365")) {
						lblNewLabel_1.setVisible(true);
						passwordField_p3.setVisible(true);
						chckbxShowPassword_p3.setVisible(true);
					}

					if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("OFFICE 365")
							|| filetype.equalsIgnoreCase("G-SUITE"))) {
						basic_Authentication.setSelected(true);
						System.out.println("here we riched");
						textField_username_p3.setEnabled(true);
						passwordField_p3.setEnabled(true);
						tf_portNo_p3.setEnabled(true);
						chckbxShowPassword_p3.setEnabled(true);
						lblPortNo.setEnabled(true);
						lblNewLabel_5.setEnabled(true);
						lblNewLabel_1.setEnabled(true);
						lblNewLabel.setEnabled(true);
						lblemailAddress.setEnabled(true);

					}
					if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")) {
						modern_Authentication.setVisible(true);
						basic_Authentication.setVisible(true);
						basic_Authentication.setEnabled(true);
						basic_Authentication.setSelected(true);
						textField_username_p3.setEnabled(true);
						passwordField_p3.setEnabled(true);
						chckbxShowPassword_p3.setEnabled(true);
						lblNewLabel.setEnabled(true);
						lblemailAddress.setEnabled(true);
						lblNewLabel_1.setEnabled(true);
						lblNewLabel_5.setEnabled(true);

					} else {
						modern_Authentication.setVisible(false);
						basic_Authentication.setVisible(false);
//				lblNewLabel.setEnabled(false);
//				lblemailAddress.setEnabled(false);
//				lblNewLabel_1.setEnabled(false);
//				lblNewLabel_5.setEnabled(false);
					}

					if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("G-SUITE")
							|| filetype.equalsIgnoreCase("Zoho Mail")) {
						lblEnableImap_p3.setVisible(true);
						lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
								+ " account , you'll need to generate and use an app password or"
								+ System.lineSeparator() + " turn on less secure app</U></HTML>");
						textField_username_p3.setEnabled(true);
						passwordField_p3.setEnabled(true);
					}

					if (filetype.equalsIgnoreCase("Live Exchange")) {
						panel_3_.setVisible(true);
						basic_Authentication.setSelected(true);
						CardLayout card = (CardLayout) panel_3_.getLayout();
						card.show(panel_3_, "panel_3_1_2");
						lbl_Domain.setText("IP or Computer Name");
						panel_3_1_2_1.setVisible(true);
						lblMakeSureYou.setVisible(false);
						lblEnableImap_p3.setVisible(false);
						lblTurnOffTwo_p3.setVisible(false);
						lblNewLabel_5.setVisible(false);
						lblTurnOffTwo_p3.setText("");
						lblMakeSureYou.setText("");
						lblEnableImap_p3.setText("");

					} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {
						panel_3_.setVisible(true);
						basic_Authentication.setSelected(true);
						CardLayout card = (CardLayout) panel_3_.getLayout();
						card.show(panel_3_, "panel_3_1_2");
						lbl_Domain.setText("Amazon Domain Name");
						panel_3_1_2_1.setVisible(true);
						tf_portNo_p3.setVisible(true);
						lblPortNo.setVisible(true);
						lblTurnOffTwo_p3.setText("");
						lblMakeSureYou.setText("");
						lblEnableImap_p3.setText("");
						lblNewLabel_5.setVisible(false);
						lblMakeSureYou.setVisible(false);
						lblEnableImap_p3.setVisible(false);
						lblTurnOffTwo_p3.setVisible(false);
						lblNewLabel_5.setVisible(false);

					} else if (filetype.equalsIgnoreCase("Hostgator email")) {
						panel_3_.setVisible(true);
						CardLayout card = (CardLayout) panel_3_.getLayout();
						basic_Authentication.setSelected(true);
						card.show(panel_3_, "panel_3_1_2");
						lbl_Domain.setText("Hostgator HOST");
						panel_3_1_2_1.setVisible(true);
						tf_portNo_p3.setVisible(true);
						lblPortNo.setVisible(true);
						lblTurnOffTwo_p3.setText("");
						lblMakeSureYou.setText("");
						lblEnableImap_p3.setText("");
						lblNewLabel_5.setVisible(false);
						lblMakeSureYou.setVisible(false);
						lblEnableImap_p3.setVisible(false);
						lblTurnOffTwo_p3.setVisible(false);
						lblNewLabel_5.setVisible(false);

					} else if (filetype.equalsIgnoreCase("IMAP")) {
						panel_3_.setVisible(true);
						CardLayout card = (CardLayout) panel_3_.getLayout();
						card.show(panel_3_, "panel_3_1_2");
						lbl_Domain.setText("IMAP HOST");
						textField_username_p3.setEnabled(true);
						passwordField_p3.setEnabled(true);
						basic_Authentication.setSelected(true);
						panel_3_1_2_1.setVisible(true);
						tf_portNo_p3.setVisible(true);
						chckbxSaveInSame.setVisible(false);
						label_13.setVisible(false);
						lblPortNo.setVisible(true);
						lblTurnOffTwo_p3.setText("");
						panel.setVisible(false);
						lblMakeSureYou.setText("");
						lblEnableImap_p3.setText("");
						lblNewLabel_5.setVisible(false);
						lblMakeSureYou.setVisible(false);
						lblEnableImap_p3.setVisible(false);
						lblTurnOffTwo_p3.setVisible(false);
						lblNewLabel_5.setVisible(false);

					} else if (filetype.equalsIgnoreCase("GoDaddy email")) {
						panel_3_.setVisible(true);
						basic_Authentication.setSelected(true);
						CardLayout card = (CardLayout) panel_3_.getLayout();
						card.show(panel_3_, "panel_3_1_2");
						lblTurnOffTwo_p3.setText("");
						textField_username_p3.setEnabled(true);
						passwordField_p3.setEnabled(true);
						lblMakeSureYou.setText("");
						lblEnableImap_p3.setText("");
						lblNewLabel_5.setVisible(false);
						lblMakeSureYou.setVisible(false);
						lblEnableImap_p3.setVisible(false);
						lblTurnOffTwo_p3.setVisible(false);
						lblNewLabel_5.setVisible(false);

					} else {

						panel_3_.setVisible(true);
						panel.setVisible(true);

						if (filetype.equalsIgnoreCase("OFFICE 365")) {
							modern_Authentication.setSelected(true);
							modern_Authentication.setVisible(true);
//					added by false to true
							textField_username_p3.setEnabled(true);
							passwordField_p3.setEnabled(false);
							chckbxShowPassword_p3.setEnabled(false);
							basic_Authentication.setVisible(true);
							lblNewLabel_1.setVisible(false);
							passwordField_p3.setVisible(false);
							chckbxShowPassword_p3.setVisible(false);
							mailbox.setVisible(true);
							archive.setVisible(true);
							publicfolder.setVisible(true);
							basic_Authentication.setEnabled(false);
						} else if (!(filetype.equalsIgnoreCase("GMAIL")
								|| filetype.equalsIgnoreCase("G-SUITE"))) {
							basic_Authentication.setSelected(false);
							modern_Authentication.setVisible(false);
						}

						if (filetype.equalsIgnoreCase("OFFICE 365")
								|| filetype.equalsIgnoreCase("Live Exchange")) {

							lblEnableImap_p3.setText("<HTML><U>To Enable IMAP</U><HTML>");

							lblEnableImap_p3.setVisible(false);
							lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
									+ " account , you'll need to generate and use an app password.</U></HTML>");
							lblMakeSureYou.setText("Please  Click on The Link");

							panel.setVisible(false);

							lblNewLabel_5.setVisible(false);

//					panel.setVisible(true);
							if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")
									|| fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
									|| fileoptionm.equalsIgnoreCase("OLM File (.olm)")) {
								chckbxRestoreToDefault.setVisible(true);
								task_box.setVisible(true);
								panel_taskfilter.setVisible(true);
							}
							lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
									+ " account , you'll need to generate and use an app password.</U></HTML>");
						}

						CardLayout card = (CardLayout) panel_3_.getLayout();
						card.show(panel_3_, "panel_3_1_2");
					}
					lbl_splitpst.setVisible(false);
					chckbx_splitpst.setSelected(false);
				} else {
					panel_progress.setVisible(true);
					panel_3_.setVisible(true);
					chckbx_splitpst.setVisible(false);
					lbl_splitpst.setVisible(false);
					if (filetype.equalsIgnoreCase("pst")) {

						chckbx_splitpst.setVisible(true);
						lbl_splitpst.setVisible(true);
					}
					if (filetype.equalsIgnoreCase("pdf")) {
						chckbxSavePdfAttachment.setVisible(true);
						label_15.setVisible(true);
						chckbx_convert_pdf_to_pdf.setVisible(true);
						label_pdf_to_pdf.setVisible(true);

					}
					if (filetype.equalsIgnoreCase("DOCX") || filetype.equalsIgnoreCase("DOC")
							|| filetype.equalsIgnoreCase("DOCM") || filetype.equalsIgnoreCase("DOCM")
							|| filetype.equalsIgnoreCase("TIFF") || filetype.equalsIgnoreCase("TXT")
							|| filetype.equalsIgnoreCase("GIF") || filetype.equalsIgnoreCase("JPG")
							|| filetype.equalsIgnoreCase("PNG") || filetype.equalsIgnoreCase("Json")
							|| filetype.equalsIgnoreCase("JPG")) {
						chckbxSavePdfAttachment.setVisible(true);
					}

					if (filetype.equalsIgnoreCase("EML") || filetype.equalsIgnoreCase("MSG")
							|| filetype.equalsIgnoreCase("EMLX") || filetype.equalsIgnoreCase("HTML")
							|| filetype.equalsIgnoreCase("MHTML")) {
						chckbxSavePdfAttachment.setVisible(true);
						label_15.setVisible(true);
					}
					if (filetype.equalsIgnoreCase("VCF") || filetype.equalsIgnoreCase("ICS")) {
						chckbxSavePdfAttachment.setVisible(true);
					}

					System.out.println("this is fileoptionm" + fileoptionm);

					if (fileoptionm.equalsIgnoreCase("MBOX") && filetype.equalsIgnoreCase("PST")) {
						label_16.setVisible(true);
						chckbxSaveMboxIn.setVisible(true);

					}

					if (fileoptionm.equalsIgnoreCase("Maildir") && filetype.equalsIgnoreCase("PST")) {
						chckbxRestoreToDefault.setVisible(true);
						panel_8.setVisible(true);
					}

					CardLayout card = (CardLayout) panel_3_.getLayout();
					card.show(panel_3_, "panel_3_1_1");

					if (!(fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")
							|| fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
							|| fileoptionm.equalsIgnoreCase("OLM File (.olm)"))) {
						chckbxMaintainFolderStructure.setVisible(false);
						label_14.setVisible(false);
						task_box.setVisible(false);
						panel_taskfilter.setVisible(false);
						chckbxRestoreToDefault.setVisible(true);
					} else {
						task_box.setVisible(true);
						panel_taskfilter.setVisible(true);
					}

					panel_3_2.setVisible(true);
					btn_converter_1.setVisible(true);
					btn_converter_1.setEnabled(true);

					if (filetype.equalsIgnoreCase("Opera Mail")) {

						String str = null;

						if (OS.contains("windows")) {
							str = System.getenv("APPDATA").replace("Roaming", "Local") + File.separator
									+ "Opera Mail" + File.separator + "Opera Mail" + File.separator + "Mail"
									+ File.separator + "store";
						} else {
							str = System.getProperty("user.home") + File.separator + "Library" + File.separator
									+ "Application Support" + File.separator + "Opera Mail" + File.separator
									+ "mail";
						}

						System.out.println(str);

						if (new File(str).exists()) {

							tf_Destination_Location.setText(str);

						} else {
							String warn = filetype + " Not Installed Do you want to proceed ?";
							int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle,
									JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
							if (ans == JOptionPane.YES_OPTION) {

							} else {
								SwingUtilities.invokeLater(new Runnable() {

									public void run() {
										comboBox_fileDestination_type.setSelectedItem("PST");
									}
								});
							}

						}

					} else if (filetype.equalsIgnoreCase("Thunderbird")) {

						String str = null;

						if (OS.contains("windows")) {
							str = System.getenv("APPDATA") + File.separator + "Thunderbird" + File.separator
									+ "Profiles";
						} else {
							str = System.getProperty("user.home") + File.separator + "Library" + File.separator
									+ "Thunderbird" + File.separator + "Profiles";
						}

						if (new File(str).exists()) {

							File[] f = new File(str).listFiles();
							for (File fl : f) {
								if (fl != null) {
									if (fl.isDirectory()) {
										String filename = fl.getName();
										String extension = filename.substring(filename.lastIndexOf(".") + 1,
												filename.length());
										String ext = "default";
										if (ext.equals(extension)) {
											// System.out.println(file);

											String defaultfolder = fl.getName();

											str = str + File.separator + defaultfolder + File.separator + "Mail"
													+ File.separator + "Local Folders";

											tf_Destination_Location.setText(str);
											break;
										} else {

										}
									}
								}
							}
						} else {
							String warn = filetype + " Not Installed Do you want to proceed ?";
							int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle,
									JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
							if (ans == JOptionPane.YES_OPTION) {

							} else {
								SwingUtilities.invokeLater(new Runnable() {

									public void run() {
										comboBox_fileDestination_type.setSelectedItem("PST");
									}
								});
							}

						}

					}

					if (!(filetype.equalsIgnoreCase("PST") || filetype.equalsIgnoreCase("Thunderbird")
							|| filetype.equalsIgnoreCase("Opera Mail") || filetype.equalsIgnoreCase("OST")
							|| filetype.equalsIgnoreCase("MBOX") || filetype.equalsIgnoreCase("CSV"))) {
						comboBox.setVisible(true);
						panel_5.setVisible(true);
					}
					if (filetype.equalsIgnoreCase("HOTMAIL")) {
						basic_Authentication.setSelected(true);
						textField_username_p3.setEnabled(true);
						passwordField_p3.setEnabled(true);
					}
				}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				// e.printStackTrace();
				System.out.println("here we riched line no 2984");
			}
		} catch (HeadlessException e) {
			// TODO Auto-generated catch block
			// e.printStackTrace();
			System.out.println("here we riched line no 2989");
		}
	}
	
	
	
	public String GetSelectedDest() {
		
		String filrtype="PST";
		
		
		
		return filetype;
	}
	
	
	
	
}
