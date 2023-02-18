package com.pdf_to_word;

import javax.swing.*;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JFileChooser;
import java.io.File;

import org.apache.commons.io.FilenameUtils;

import com.spire.pdf.FileFormat;
import com.spire.pdf.PdfDocument;

public class UI {

    private JFrame window;
    private JButton openButton;

    public UI() {
        window = new JFrame("Convert Word To PDF");
        openButton = new JButton("Open Word To Convert To PDF");
        openButton.setBounds(200, 100, 300, 50);

        window.setLayout(null);

        window.setLayout(null);
        window.setSize(800, 800);
        window.setVisible(true);
        window.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        window.add(openButton);

        openButton.addActionListener(new ActionListener() {

            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();

                fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);

                int returnValue = fileChooser.showDialog(fileChooser, null);

                try {
                    if (JFileChooser.APPROVE_OPTION == returnValue) {
                        File[] files = fileChooser.getSelectedFile().listFiles();

                        for (File f : files) {
                            String extension = FilenameUtils.getExtension(f.getAbsolutePath());

                            if (extension.toLowerCase().compareTo("pdf") == 0) {

                                String saveName = FilenameUtils.getBaseName(f.getAbsolutePath()) + ".docx";
                                String placeToSave = FilenameUtils.getFullPath(f.getAbsolutePath()) + saveName;

                                PdfDocument doc = new PdfDocument();
                                doc.loadFromFile(f.getAbsolutePath());
                                doc.saveToFile(placeToSave, FileFormat.DOCX);
                                doc.close();
                            }
                        }

                        JOptionPane.showMessageDialog(window, "Pdfs Converted To Word Successfully", "Success",
                                JOptionPane.INFORMATION_MESSAGE);
                    }

                } catch (Exception ex) {
                    System.out.println("Error while Opening");
                    JOptionPane.showMessageDialog(window, "Error while converting, maybe select Properly", "Errorr",
                            JOptionPane.ERROR_MESSAGE);
                }

            }

        });

    }
}
