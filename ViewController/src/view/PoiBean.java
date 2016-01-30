package view;

import java.awt.Desktop;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class PoiBean {
    public PoiBean() {
    }
    public String generateDoc() throws FileNotFoundException, IOException {
        String titleContent = "Title";
        String mainContent = "Main content";
        
        XWPFDocument document = new XWPFDocument(); //��������  docx ���������
        
        XWPFParagraph titileParagraph = document.createParagraph(); //���������� ��������� � ���������
        XWPFRun titleRun = titileParagraph.createRun(); //����� ��� ���������� ������ � ��������
        titleRun.setText(titleContent);  //������� ������
        titleRun.setFontSize(18);   //��������� ������� ������
        titileParagraph.setAlignment(ParagraphAlignment.CENTER); //������������
        
        XWPFParagraph mainParagraph = document.createParagraph();
        XWPFRun mainRun = mainParagraph.createRun(); 
        mainRun.setText(mainContent);
        mainRun.setFontSize(15);
        mainParagraph.setAlignment(ParagraphAlignment.BOTH);
        
        FileOutputStream out = new FileOutputStream(new File("C:\\Work\\createdocument.docx")); //���������� ��������� � �������� �������
        document.write(out);
        out.close();
        
        Desktop dt = Desktop.getDesktop(); // �������� ��������� ������������ ��������
        dt.open(new File("C:\\Work\\createdocument.docx"));
        
        return null;
    }
}
