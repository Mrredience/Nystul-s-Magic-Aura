import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;

import java.io.*;

public class NystulSMagicAura {
    public static void main(String[] args) {
        // 指定 Word 文件路径
        String inputFilePath = "input.docx"; // 或 "input.doc"
        String outputFilePath = "output.txt";

        try {
            if (inputFilePath.endsWith(".docx")) {
                convertDocxToTxt(inputFilePath, outputFilePath);
            } else if (inputFilePath.endsWith(".doc")) {
                convertDocToTxt(inputFilePath, outputFilePath);
            } else {
                System.out.println("不支持的文件格式。");
            }
        } catch (IOException e) {
            System.err.println("转换过程中出错: " + e.getMessage());
        }
    }

    // 将 .docx 文件转换为 .txt
    public static void convertDocxToTxt(String inputFilePath, String outputFilePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             XWPFDocument document = new XWPFDocument(fis);
             BufferedWriter writer = new BufferedWriter(new FileWriter(outputFilePath))) {

            document.getParagraphs().forEach(paragraph -> {
                try {
                    writer.write(paragraph.getText());
                    writer.newLine();
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            });

            System.out.println("文件成功转换为 TXT: " + outputFilePath);
        }
    }

    // 将 .doc 文件转换为 .txt
    public static void convertDocToTxt(String inputFilePath, String outputFilePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             HWPFDocument document = new HWPFDocument(fis);
             WordExtractor extractor = new WordExtractor(document);
             BufferedWriter writer = new BufferedWriter(new FileWriter(outputFilePath))) {

            String[] paragraphs = extractor.getParagraphText();
            for (String paragraph : paragraphs) {
                writer.write(paragraph);
                writer.newLine();
            }

            System.out.println("文件成功转换为 TXT: " + outputFilePath);
        }
    }
}
