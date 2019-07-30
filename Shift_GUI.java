/*
 * Develoed by Barış Bülbül on 23.07.2019 14:57.
 * Name of the current project excelread.
 * Name of the current module ExcelRead.
 * Last modified 22.07.2019 11:11.
 * StjBarisB.
 * Copyright (c) 2019. All rights reserved.
 */

package com.gt.shift;

import org.apache.poi.ss.usermodel.*;

import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;
import java.util.Vector;

public class Shift_GUI {
    private JFileChooser jChooser;
    private JButton SENDButton;
    private JPanel mainPanel;
    private JButton SEARCHButton;
    private JTextArea enterYourMassageTextArea;
    private JTable table1;
    private JComboBox<String> toMail;
    private Vector<String> headers = new Vector<>();
    private Vector<Vector> data = new Vector<>();
    private DefaultTableModel model = null;
    private static int tableWidth = 0;
    private static int tableHeight = 0;

    private Shift_GUI() {
        jChooser = new JFileChooser();
        String[] mails = new String[]{"asd"};
        for (String x :
                mails) {
            toMail.addItem(x);
        }
        toMail.setEditable(true);
        SEARCHButton.addActionListener(actionEvent -> {
            jChooser.showOpenDialog(null);
            File excelFile = jChooser.getSelectedFile();
            if (excelFile.getName().endsWith("xls") || excelFile.getName().endsWith("xlsx")) {
                fillData(excelFile);
                model = new DefaultTableModel(data, headers);
                model.removeRow(0);
                tableWidth = model.getColumnCount() * 150;
                tableHeight = model.getRowCount() * 25;
                table1.setPreferredSize(new Dimension(tableWidth, tableHeight));
                table1.setModel(model);
            } else {
                JOptionPane.showMessageDialog(null,
                        "Please Select Only Excel File.",
                        "Error", JOptionPane.ERROR_MESSAGE);
            }
        });
        SENDButton.addActionListener(actionEvent -> sendMail());
    }

    private void fillData(File excelFile) {
        String mailMassage = null;
        SimpleDateFormat dateFormat = new SimpleDateFormat("M/dd/yy");
        Date date = new Date();
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new File(String.valueOf(excelFile)));
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        assert workbook != null;
        Sheet sheet = workbook.getSheetAt(0);
        headers.clear();
        DataFormatter dataFormatter = new DataFormatter();
        for (int i = 0; i < sheet.getRow(0).getPhysicalNumberOfCells() + 1; i++) {
            Cell cell = sheet.getRow(0).getCell(i);
            String cellValue = dataFormatter.formatCellValue(cell);
            headers.add(cellValue);
        }
        data.clear();
        for (Row row : sheet) {
            Vector<String> datas = new Vector<>();
            for (Cell cell : row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                datas.add(cellValue);
                if (cellValue.equalsIgnoreCase(dateFormat.format(date)) && cell.getColumnIndex() == 1) {
                    int x = cell.getRowIndex();
                    int y = cell.getColumnIndex() + 2;
                    String shiftName = String.valueOf(workbook.getSheetAt(0).getRow(x).getCell(y));
                    String shiftNu = String.valueOf(workbook.getSheetAt(0).getRow(x).getCell(y + 1));
                    mailMassage = "Günün(" + cellValue + ")talihlisi " + shiftNu.substring(0, shiftNu.length() - 2)
                            + ".nöbetini tutan " + shiftName + "'dır.";
                }
                if (cell.getColumnIndex() == 5 && cellValue.equalsIgnoreCase(dateFormat.format(date)))
                    mailMassage = dateFormat.format(date) + "'tarihi tatildir.";
            }
            data.add(datas);
        }
        try {
            workbook.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        if (mailMassage != null) {
            JOptionPane.showMessageDialog(null, mailMassage);
            enterYourMassageTextArea.setText(mailMassage);
        } else {
            mailMassage = dateFormat.format(date) + "'tarihinde kimse nöbetci değildir.";
            JOptionPane.showMessageDialog(null, mailMassage);
            enterYourMassageTextArea.setText(mailMassage);
        }
    }

    private void sendMail() {
        String userName = "********************";
        String password = "********************";
        Properties prop = new Properties();
        prop.put("mail.smtp.host", "**********************");
        prop.put("mail.smtp.port", "**************");
        prop.put("mail.smtp.auth", "****");
        prop.put("mail.smtp.starttls.enable", "*****");
        Session session = Session.getInstance(prop,
                new javax.mail.Authenticator() {
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(
                                userName, password);
                    }
                });
        try {
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress("*****************************"));
            message.setRecipients(
                    Message.RecipientType.TO,
                    InternetAddress.parse(String.valueOf(toMail.getSelectedItem()))
            );
            message.setSubject("Günün Nöbetcisi Hk.");
            message.setText(enterYourMassageTextArea.getText());
            Transport.send(message);
            JOptionPane.showMessageDialog(null, "E-Mail Gönderildi");
            enterYourMassageTextArea.setText("");
            toMail.setSelectedItem(null);
        } catch (
                MessagingException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws ClassNotFoundException, UnsupportedLookAndFeelException, InstantiationException, IllegalAccessException {
        UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
        JFrame jFrame = new JFrame("Shift Application");
        jFrame.setContentPane(new Shift_GUI().mainPanel);
        jFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        jFrame.pack();
        jFrame.setLocationRelativeTo(null);
        jFrame.setSize(800, 400);
        jFrame.setResizable(true);
        jFrame.setVisible(true);
    }
}
