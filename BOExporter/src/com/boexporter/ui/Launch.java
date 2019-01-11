package com.boexporter.ui;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JTextField;

import com.boexporter.logic.ExportBOUsers;
import com.boexporter.tools.Configure;


public class Launch extends JFrame {
	
	private static JButton buttonExportData;
	private static JButton buttonReset;

	 public static JLabel labelBOCMSUrl;
	 public static JLabel labelBOUserName;
	 public static JLabel labelBOPassword;
	 public static JLabel labelExportExcelPath;
	    
	 private static JTextField textBOCMSUrl;
	 private static JTextField textBOUserName;
	 private static JPasswordField pwdtextBOPassword;
	 private static JTextField textExportExcelPath;

    private static final long serialVersionUID = 1L;
    static String path = null;

    public static void main(final String[] args) {

        // ��������
        final JFrame frame = new JFrame();
        frame.setLayout(new BorderLayout());
        frame.setTitle("BO User and Report Exporter");
        frame.setSize(500, 500);
        
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        //������ʾ
        Dimension displaySize = Toolkit.getDefaultToolkit().getScreenSize(); // �����ʾ����С����  
        Dimension frameSize = frame.getSize();             // ��ô��ڴ�С����  
        if (frameSize.width > displaySize.width)  
        frameSize.width = displaySize.width;           // ���ڵĿ�Ȳ��ܴ�����ʾ���Ŀ��  

        frame.setLocation((displaySize.width - frameSize.width) / 2,  
        (displaySize.height - frameSize.height) / 2); // ���ô��ھ�����ʾ����ʾ  

        // ��Ӱ�ť
        buttonExportData = new JButton("Export");
       
        buttonExportData.setBounds(80, 70, 60, 25);
  
        buttonExportData.setVisible(true);
        
        buttonReset  = new JButton("Reset");
        buttonReset.setBounds(80, 70, 60, 25);
        
        buttonReset.setVisible(true);
        
      
        
        

        // ��������
        final Font font = new Font("����", Font.BOLD, 20);
        buttonExportData.setFont(font);
        buttonReset.setFont(font);
        
        // ���ѡ��
        labelBOCMSUrl = new JLabel("CMS URL: ");
        labelBOUserName = new JLabel("User Name: ");
        labelBOPassword  = new JLabel("Password: ");
        labelExportExcelPath  = new JLabel("Export File Name");
        
        //��ȡ�����ļ�
        Configure config = new Configure("config.properties");
        
        
        // �����
        
        

        
        textBOCMSUrl = new JTextField(config.getValue("bo.cmsurl"));
        textBOUserName = new JTextField(config.getValue("bo.username"));
        pwdtextBOPassword = new JPasswordField(20);
        String passwordStr = config.getValue("bo.password");
        pwdtextBOPassword.setText(passwordStr);
        textExportExcelPath = new JTextField(config.getValue("bo.exportpath"));
   

        // �����ʾ��

        
        //����Panel
        JPanel panelGrid = new JPanel();
        panelGrid.setLayout(new GridLayout(5, 2));
        // ������
        panelGrid.add(labelBOCMSUrl);
        panelGrid.add(textBOCMSUrl);
        panelGrid.add(labelBOUserName);
        panelGrid.add(textBOUserName);
        panelGrid.add(labelBOPassword);
        panelGrid.add(pwdtextBOPassword);
        panelGrid.add(labelExportExcelPath);
        panelGrid.add(textExportExcelPath);
        panelGrid.add(buttonExportData);
        panelGrid.add(buttonReset);
        panelGrid.setVisible(true);
        
        frame.setContentPane(panelGrid);

        // ����ı���
       
        frame.setVisible(true);

        // Ϊ��ť����¼�
        buttonExportData.addActionListener(new ActionListener() {

            @Override
            public void actionPerformed(final ActionEvent e) {
            	buttonExportData.setEnabled(false);
            	buttonExportData.setEnabled(false);
            	
            	String boUser = textBOUserName.getText();
        		String boPassword = new String(pwdtextBOPassword.getPassword());
        		String boCmsName = textBOCMSUrl.getText();
        		String excelPath = textExportExcelPath.getText();
        		ExportBOUsers exporter = new ExportBOUsers(boCmsName,boUser,boPassword,excelPath);
        		// exporter.saveBoUsersToExcel();
        		try {
					boolean result = exporter.saveBoUsersAndReportListToExcel();
					
					if(result) {
	            		
	            		String message = "Done";
	            		
	            		JOptionPane.showOptionDialog(null, message, "��Ϣ", JOptionPane.DEFAULT_OPTION, JOptionPane.WARNING_MESSAGE,null,null,null); 
	            		
	            	}else {
	            		String message = "Failed";
	            		JOptionPane.showOptionDialog(null, message, "!!����!!", JOptionPane.DEFAULT_OPTION, JOptionPane.ERROR_MESSAGE,null,null,null); 
	            	}
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
					String message = "Faild, Please Check Log";
            		JOptionPane.showOptionDialog(null, message, "!!����!!", JOptionPane.DEFAULT_OPTION, JOptionPane.ERROR_MESSAGE,null,null,null); 
            	
				}
        		buttonExportData.setEnabled(true);
            	buttonExportData.setEnabled(true);
                }
            
        });
        
        // ���ð�ť����¼�
        
        buttonReset.addActionListener(new ActionListener() {

            @Override
            public void actionPerformed(final ActionEvent e) {
            	
            	textBOUserName.setText("");
        		pwdtextBOPassword.setText("");
        		textBOCMSUrl.setText("");
        		textExportExcelPath.setText("");
            }
        		
            
        });

      

       
    }

}




