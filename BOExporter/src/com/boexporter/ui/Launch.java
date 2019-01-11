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

        // 画出窗口
        final JFrame frame = new JFrame();
        frame.setLayout(new BorderLayout());
        frame.setTitle("BO User and Report Exporter");
        frame.setSize(500, 500);
        
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        //居中显示
        Dimension displaySize = Toolkit.getDefaultToolkit().getScreenSize(); // 获得显示器大小对象  
        Dimension frameSize = frame.getSize();             // 获得窗口大小对象  
        if (frameSize.width > displaySize.width)  
        frameSize.width = displaySize.width;           // 窗口的宽度不能大于显示器的宽度  

        frame.setLocation((displaySize.width - frameSize.width) / 2,  
        (displaySize.height - frameSize.height) / 2); // 设置窗口居中显示器显示  

        // 添加按钮
        buttonExportData = new JButton("Export");
       
        buttonExportData.setBounds(80, 70, 60, 25);
  
        buttonExportData.setVisible(true);
        
        buttonReset  = new JButton("Reset");
        buttonReset.setBounds(80, 70, 60, 25);
        
        buttonReset.setVisible(true);
        
      
        
        

        // 设置字体
        final Font font = new Font("宋体", Font.BOLD, 20);
        buttonExportData.setFont(font);
        buttonReset.setFont(font);
        
        // 添加选项
        labelBOCMSUrl = new JLabel("CMS URL: ");
        labelBOUserName = new JLabel("User Name: ");
        labelBOPassword  = new JLabel("Password: ");
        labelExportExcelPath  = new JLabel("Export File Name");
        
        //读取配置文件
        Configure config = new Configure("config.properties");
        
        
        // 输入框
        
        

        
        textBOCMSUrl = new JTextField(config.getValue("bo.cmsurl"));
        textBOUserName = new JTextField(config.getValue("bo.username"));
        pwdtextBOPassword = new JPasswordField(20);
        String passwordStr = config.getValue("bo.password");
        pwdtextBOPassword.setText(passwordStr);
        textExportExcelPath = new JTextField(config.getValue("bo.exportpath"));
   

        // 添加显示框

        
        //设置Panel
        JPanel panelGrid = new JPanel();
        panelGrid.setLayout(new GridLayout(5, 2));
        // 添加组件
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

        // 添加文本框
       
        frame.setVisible(true);

        // 为按钮添加事件
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
	            		
	            		JOptionPane.showOptionDialog(null, message, "信息", JOptionPane.DEFAULT_OPTION, JOptionPane.WARNING_MESSAGE,null,null,null); 
	            		
	            	}else {
	            		String message = "Failed";
	            		JOptionPane.showOptionDialog(null, message, "!!错误!!", JOptionPane.DEFAULT_OPTION, JOptionPane.ERROR_MESSAGE,null,null,null); 
	            	}
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
					String message = "Faild, Please Check Log";
            		JOptionPane.showOptionDialog(null, message, "!!错误!!", JOptionPane.DEFAULT_OPTION, JOptionPane.ERROR_MESSAGE,null,null,null); 
            	
				}
        		buttonExportData.setEnabled(true);
            	buttonExportData.setEnabled(true);
                }
            
        });
        
        // 重置按钮添加事件
        
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




