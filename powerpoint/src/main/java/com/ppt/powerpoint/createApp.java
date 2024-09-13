package com.ppt.powerpoint;

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.Color;
import java.awt.Dimension;

@Service
public class createApp {

    private final textExtractor objTextExtractor = new textExtractor();

    public void createSlideMethod(String textHeadString, String textBodyString, XMLSlideShow ppt) {
        XSLFSlide slide = ppt.createSlide();

        if (textHeadString != null) {
            XSLFTextBox headBox = slide.createTextBox();
            headBox.setAnchor(new java.awt.Rectangle(100, 20, 800, 80));
            XSLFTextParagraph headingParagraph = headBox.addNewTextParagraph();
            XSLFTextRun headingTextRun = headingParagraph.addNewTextRun();
            textHeadString = textHeadString.toUpperCase();
            headingTextRun.setText(textHeadString);
            headingTextRun.setBold(true);
            headingTextRun.setFontSize(40.0);
            headingTextRun.setFontColor(new Color(128, 128, 128));
            headingTextRun.setFontFamily("Arial");
        }

        XSLFTextBox textBox = slide.createTextBox();
        textBox.setAnchor(new java.awt.Rectangle(100, 120, 400, 300));

        String[] lineText = objTextExtractor.pointLocator(objTextExtractor.bodyExtractor(textBodyString)).split("\n");

        for (String  text : lineText) {
            XSLFTextParagraph bullet = textBox.addNewTextParagraph();
            String actText = objTextExtractor.bodyExtractor(text);
            if (actText.charAt(0) != '/' && actText.charAt(1) != '-'){
                bullet.setBullet(true);
            } else {
                actText = actText.replaceFirst("/-", "-");
                bullet.setBullet(false);
            }
            bullet.setIndent(0.3 * 72);
            XSLFTextRun bulletTextRun = bullet.addNewTextRun();
            bulletTextRun.setText(actText);
            bulletTextRun.setFontSize(18.0);
        }
    }

    public void createPPT(String titleText, String textBodyString, String pptSavePath) throws IOException {
        
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            ppt.setPageSize(new Dimension(960, 540));

            XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);

            XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE);

            // Create a slide with the title layout
            XSLFSlide slide1 = ppt.createSlide(titleLayout);

            // Set the title and subtitle
            XSLFTextShape title = slide1.getPlaceholder(0);
            title.setText(titleText);

            for(XSLFTextParagraph paragraph: title.getTextParagraphs()){
                for(XSLFTextRun run: paragraph.getTextRuns()) {
                    run.setFontSize(45.0);
                    run.setBold(true);
                    String upperCaseString = run.getRawText().toUpperCase();
                    run.setText(upperCaseString);
                }
            }
            
            XSLFTextShape subTitle = slide1.getPlaceholder(1);
            subTitle.setText("-By Kannadathi");
            
            for(XSLFTextParagraph paragraph: subTitle.getTextParagraphs()){
                paragraph.setTextAlign(TextAlign.RIGHT);
                for(XSLFTextRun run: paragraph.getTextRuns()) {
                    run.setFontSize(20.0);;
                }
            }

            String[] paragraphs = objTextExtractor.paragraphSeperator(textBodyString);

            for (String paragraphText : paragraphs) {
                String textHeadString = objTextExtractor.headingExtractor(paragraphText);
                createSlideMethod(textHeadString, paragraphText, ppt);
            }

            try (FileOutputStream out = new FileOutputStream(pptSavePath)) {
                ppt.write(out);
            }
        }
    }
}

class textExtractor {
    String headingExtractor(String textToExtract) {
        int newLineIndex = textToExtract.indexOf("\n");
        if (newLineIndex == -1 || !Character.isDigit(textToExtract.charAt(0))) {
            return null;
        }
        String extractedHeading = textToExtract.substring(0, newLineIndex);
        return extractedHeading.replaceFirst("\\d+\\.", "").trim();
    }

    String bodyExtractor(String textToExtract) {
        int newLineIndex = textToExtract.indexOf("\n");
        if (newLineIndex == -1) {
            return spaceRemover(textToExtract);
        }

        return spaceRemover(textToExtract.substring(newLineIndex + 1));
    }

    String pointLocator(String textToExtract) {
        String extractedText= textToExtract.replaceAll("     ", "/-");
        extractedText = extractedText.replaceAll("   -", "-");
        return extractedText;
    }

    String spaceRemover(String textToRemove) {
        String removedText = textToRemove.trim();
        removedText = removedText.replaceFirst("-", "");
        return removedText;
    }

    String[] paragraphSeperator(String text) {
        return text.split("\n\n");
    }
    
}

