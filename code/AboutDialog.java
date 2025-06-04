package email.code;

import java.awt.Color;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;

import components.GradientButton;
import components.Label;
import javax.swing.border.TitledBorder;
import java.awt.SystemColor;
import javax.swing.border.CompoundBorder;
import javax.swing.border.BevelBorder;
import javax.swing.border.LineBorder;

@SuppressWarnings("serial")
public class AboutDialog extends JDialog {
	static String aboutTitle;
	private final JPanel contentPanel = new JPanel();
	String labeltext = "";
	Main_Frame mf;

	/**
	 * Launch the application.
	 */

	/**
	 * Create the dialog.
	 * 
	 * @param modal
	 * @param parent
	 */

	public AboutDialog(JFrame parent, boolean modal, String labeltext) {

		super(parent, true);
		this.labeltext = labeltext;
		aboutTitle = Main_Frame.projectTitle;
		setTitle(aboutTitle);

		setIconImage(Toolkit.getDefaultToolkit().getImage(AboutDialog.class.getResource("/128x128.png")));
		setResizable(false);
		setBounds(100, 100, 351, 362);
		getContentPane().setLayout(null);
		contentPanel.setBackground(SystemColor.inactiveCaption);
		contentPanel.setBounds(0, 0, 335, 337);
		contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
		getContentPane().add(contentPanel);
		contentPanel.setLayout(null);

		JPanel panel = new JPanel();
		panel.setBorder(new TitledBorder(new CompoundBorder(new BevelBorder(BevelBorder.RAISED, new Color(240, 240, 240), new Color(255, 255, 255), new Color(105, 105, 105), new Color(160, 160, 160)), new LineBorder(new Color(180, 180, 180))), "<html><b>Product Details", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 120, 215)));
		panel.setBackground(SystemColor.menu);
		panel.setBounds(5, 38, 325, 249);
		contentPanel.add(panel);
		panel.setLayout(null);

		JLabel lblNewLabel_2 = new JLabel("Edition:              Standard ");
		lblNewLabel_2.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		lblNewLabel_2.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		lblNewLabel_2.setBounds(22, 59, 255, 24);
		panel.add(lblNewLabel_2);

		JLabel lblVersionStandard = new JLabel("Version:              " + All_Data.version);
		lblVersionStandard.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		lblVersionStandard.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		lblVersionStandard.setBounds(22, 87, 255, 24);
		panel.add(lblVersionStandard);

		JLabel lblLicencseStandard = new JLabel("Licensed To:       " + labeltext);
		lblLicencseStandard.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		lblLicencseStandard.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		lblLicencseStandard.setBounds(22, 117, 255, 21);
		panel.add(lblLicencseStandard);

		JLabel websitelink = new JLabel("");
		websitelink.setBounds(30, 173, 163, 20);
		panel.add(websitelink);
		websitelink.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		websitelink.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		websitelink.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent arg0) {
				try {
					Desktop.getDesktop().browse(new URI("https://www.DevopixTech.com"));
				} catch (URISyntaxException | IOException ex) {
					// It looks like there's a problem
				}
			}
		});
		websitelink.setForeground(new Color(0, 0, 205));
		websitelink.setText("Home");
		websitelink.setCursor(new Cursor(Cursor.HAND_CURSOR));

		JLabel lblSupportInformation = new JLabel("<html><u>Support Information");
		lblSupportInformation.setBounds(20, 147, 181, 22);
		panel.add(lblSupportInformation);
		lblSupportInformation.setFont(new Font("Segoe UI", Font.BOLD, 12));

		JLabel supportlink = new JLabel("");
		supportlink.setBounds(30, 194, 164, 16);
		panel.add(supportlink);
		supportlink.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		supportlink.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		supportlink.setForeground(new Color(0, 0, 205));
		supportlink.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				try {
					Desktop.getDesktop().browse(
							new URI("http://messenger.providesupport.com/messenger/0pi295uz3ga080c7lxqxxuaoxr.html"));
				} catch (URISyntaxException | IOException ex) {
					// It looks like there's a problem
				}
			}
		});
		supportlink.setText("Live Chat");
		supportlink.setCursor(new Cursor(Cursor.HAND_CURSOR));
		supportlink.setCursor(new Cursor(Cursor.HAND_CURSOR));

		JLabel saleslink = new JLabel("DevopixTechSoftwareSolution@gmail.com");
		saleslink.setBounds(30, 211, 269, 23);
		panel.add(saleslink);
		saleslink.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				try {
					Desktop.getDesktop().mail(new URI("mailto:DevopixTechSoftwareSolution@gmail.com" + ""));
				} catch (URISyntaxException | IOException ex) {
					// It looks like there's a problem
				}
			}
		});
		saleslink.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
		saleslink.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		saleslink.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		saleslink.setForeground(new Color(0, 0, 205));
		
//		JLabel more_products = new JLabel("Explore more products");
//		more_products.setForeground(Color.BLUE);
//		more_products.setBounds(10, 204, 245, 24);
//		panel.add(more_products);
//		
//		more_products.addMouseListener(new MouseAdapter() {
//		
//			public void mouseClicked(MouseEvent e) {
//				
//				try {
//					Desktop.getDesktop().browse(new URI(All_Data.exploremoreproducts));
//				} catch (URISyntaxException | IOException ex) {
//					// It looks like there's a problem
//				}
//			
//			
//
//				
//			}
//		});
//		more_products.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
//		more_products.setFont(new Font("Segoe UI", Font.PLAIN, 12));
//		more_products.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		
		//		JLabel lblNewLabel = new JLabel("");
		//		lblNewLabel.setIcon(new ImageIcon(AboutDialog.class.getResource("/about.png")));
		//		lblNewLabel.setBounds(0, 0, 214, 337);
		//		contentPanel.add(lblNewLabel);
		
				JLabel lblNewLabel_1 = new JLabel("<html><u>"+aboutTitle);
				lblNewLabel_1.setBounds(10, 26, 305, 28);
				panel.add(lblNewLabel_1);
				lblNewLabel_1.setFont(new Font("Segoe UI Semibold", Font.BOLD, 13));
		
		
		{
			GradientButton okButton = new GradientButton("Close");
			okButton.setFont(new Font("Tahoma", Font.BOLD, 11));
			okButton.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseEntered(MouseEvent e) {
					okButton.setGradientColor1(new Color(70, 130, 180));
					okButton.setGradientColor2(new Color(70, 130, 180));
					okButton.setForeground(new Color(255, 255, 255));
				}

				@Override
				public void mouseExited(MouseEvent e) {
					okButton.setGradientColor1(new Color(255, 255, 255));
					okButton.setGradientColor2(new Color(255, 255, 255));
					okButton.setForeground(new Color(80, 80, 80));
				}
				
			});
			okButton.setShadowColor(new Color(0, 0, 255));
			okButton.setForeground(UIManager.getColor("CheckBox.shadow"));
			okButton.setRolloverEnabled(false);
			okButton.setRequestFocusEnabled(false);
			okButton.setOpaque(false);
			okButton.setFocusable(false);
			okButton.setFocusTraversalKeysEnabled(false);
			okButton.setFocusPainted(false);
			okButton.setDefaultCapable(false);
			okButton.setContentAreaFilled(false);
			okButton.setBorderPainted(false);
			
//			okButton.setContentAreaFilled(false);
//			okButton.setBorderPainted(false);
//			okButton.addMouseListener(new MouseAdapter() {
//				@Override
//				public void mouseEntered(MouseEvent e) {
//					okButton.setIcon(new ImageIcon(AboutDialog.class.getResource("/ok_about-hvr.png")));
//				}
//
//				@Override
//				public void mouseExited(MouseEvent e) {
//					okButton.setIcon(new ImageIcon(AboutDialog.class.getResource("/ok_about.png")));
//				}
//			});
//			okButton.setFocusPainted(false);
//			okButton.setIcon(new ImageIcon(AboutDialog.class.getResource("/ok_about.png")));
			okButton.setBounds(108, 290, 122, 35);
			contentPanel.add(okButton);
			okButton.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent arg0) {
					dispose();
				}
			});
			okButton.setActionCommand("OK");
			getRootPane().setDefaultButton(okButton);
		}

//		JButton userlicButton = new JButton("");
//		userlicButton.addActionListener(new ActionListener() {
//			public void actionPerformed(ActionEvent e) {
//				try {
//					Desktop.getDesktop()
//							.browse(new URI("https://www.arysontechnologies.com/pdf/Eula%20for%20Aryson.pdf"));
//				} catch (URISyntaxException | IOException ex) {
//					// It looks like there's a problem
//				}
//			}
//		});
//		userlicButton.addMouseListener(new MouseAdapter() {
//			@Override
//			public void mouseEntered(MouseEvent arg0) {
//				userlicButton.setIcon(new ImageIcon(AboutDialog.class.getResource("/view-user-license-hvr.png")));
//			}
//
//			@Override
//			public void mouseExited(MouseEvent e) {
//				userlicButton.setIcon(new ImageIcon(AboutDialog.class.getResource("/view-user-license.png")));
//			}
//		});
//		userlicButton.setIcon(new ImageIcon(AboutDialog.class.getResource("/view-user-license.png")));
//		userlicButton.setFocusPainted(false);
//		userlicButton.setBorderPainted(false);
//		userlicButton.setContentAreaFilled(false);
//		userlicButton.setBounds(51, 291, 117, 27);
//		contentPanel.add(userlicButton);

//		JLabel lblProductInformation = new JLabel("Product Details");
//		lblProductInformation.setFont(new Font("Segoe UI", Font.BOLD, 12));
//		lblProductInformation.setBounds(10, 36, 138, 22);
//		contentPanel.add(lblProductInformation);

		Label lblNewLabel_4 = new Label("Copyright \u00A9DevopixTech Software Solution");
		lblNewLabel_4.setGradientColor1(new Color(154, 205, 50));
		lblNewLabel_4.setGradientColor2(new Color(255, 215, 0));
		lblNewLabel_4.setShadowColor(Color.BLACK);
		lblNewLabel_4.setFont(new Font("Segoe UI", Font.BOLD, 12));
		lblNewLabel_4.setBounds(0, 5, 335, 35);
		contentPanel.add(lblNewLabel_4);
	}
}
