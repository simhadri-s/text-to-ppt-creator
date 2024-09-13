package com.ppt.powerpoint;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.context.ConfigurableApplicationContext;

import javax.swing.*;
import javax.swing.border.EmptyBorder;

import java.awt.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;

@SpringBootApplication
public class PowerpointApplication {

    @Autowired
    private createApp createApp;

    public static void main(String[] args) {
        SpringApplicationBuilder builder = new SpringApplicationBuilder(PowerpointApplication.class);
        builder.headless(false);
        ConfigurableApplicationContext context = builder.run(args);

        SwingUtilities.invokeLater(() -> {
            PowerpointApplication app = context.getBean(PowerpointApplication.class);
            app.createAndShowGUI();
        });
    }

    private void createAndShowGUI() {
        JFrame mainFrame = new JFrame("Text To PPT Creator");
        mainFrame.setSize(800, 600);
        mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        ImageIcon windowLogoIcon = new ImageIcon(getClass().getResource("/logo.png"));
        Image windowImage = windowLogoIcon.getImage();
        mainFrame.setIconImage(windowImage);
        mainFrame.setLayout(new BorderLayout());
        JPanel mainPanel = new JPanel(new BorderLayout());
        JLabel inputFieldLabel = new JLabel("Enter the Title of the PPT:");
        inputFieldLabel.setFont(new Font("Serif", Font.BOLD, 20));
        JTextField inputTitle = new JTextField(20);
        inputTitle.setPreferredSize(new Dimension(200, 40));
        JLabel inputBoxLabel = new JLabel("Paste the content here:");
        inputBoxLabel.setFont(new Font("Serif", Font.BOLD, 20));
        JTextArea inputField = new JTextArea(20, 40);
        JScrollPane scrollPane = new JScrollPane(inputField);
        JButton createButton = new JButton("Create PPT");
        createButton.setPreferredSize(new Dimension(200, 45));
        createButton.setFont(new Font("Serif", Font.BOLD, 20));
        createButton.setCursor(new Cursor(Cursor.HAND_CURSOR));
        createButton.addActionListener(e -> {
            String title = inputTitle.getText();
            String text = inputField.getText();
            text = text.replaceAll("\\*\\*", "");
            text = text.replaceAll("\n---\n", "");
            text = text.replaceAll("###", "");
            try{
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch(ClassNotFoundException | InstantiationException | IllegalAccessException | UnsupportedLookAndFeelException f) {
                f.printStackTrace();
            }

            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Save File");
            int userSelection = fileChooser.showSaveDialog(null);

            String outputPath;
            if(userSelection == JFileChooser.APPROVE_OPTION) {
                File  fileOutputPath = fileChooser.getSelectedFile();
                outputPath = fileOutputPath.getAbsolutePath();
                if(!outputPath.endsWith(".pptx")){
                    outputPath=outputPath+".pptx";
                }
                try {
                    createApp.createPPT(title, text, outputPath);
                    JOptionPane.showMessageDialog(mainFrame, "PPT created successfully");
                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(mainFrame, "Error creating PPT: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
            
        });
        
        JPanel textFieldJPanel = new JPanel(new BorderLayout());
        JPanel textBoxJPanel = new JPanel(new BorderLayout());
        textFieldJPanel.add(inputFieldLabel,BorderLayout.NORTH);
        textFieldJPanel.add(inputTitle, BorderLayout.SOUTH);
        textBoxJPanel.add(inputBoxLabel, BorderLayout.NORTH);
        textBoxJPanel.add(scrollPane, BorderLayout.CENTER);
        mainPanel.add(textFieldJPanel, BorderLayout.NORTH);
        mainPanel.add(textBoxJPanel, BorderLayout.CENTER);
        mainPanel.setBorder(new EmptyBorder(10, 10, 10 , 10));
        mainPanel.add(createButton, BorderLayout.SOUTH);


        mainFrame.add(mainPanel);
        mainFrame.setVisible(true);

        // Ensure application closes when the frame is closed
        mainFrame.addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                System.exit(0);
            }
        });
    }
}

