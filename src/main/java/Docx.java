import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class Docx {
    private String name;
    private String dir;
    private XWPFDocument doc;

    public Docx(String name, String dir) {
        this.name = name;
        this.dir = dir;
    }

    public void Read() {
        try {
            doc = new XWPFDocument(OPCPackage.open(name));
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public void Write() {
        try {
            doc.write(new FileOutputStream(dir + name));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void Replace(String key, String tag) {
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.contains("<" + key + ">")) {
                        text = text.replace("<" + key + ">", tag);//your content
                        r.setText(text, 0);
                    }
                }
            }
        }

        for (XWPFTable tbl : doc.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            if (text != null && text.contains("<" + key + ">")) {
                                text = text.replace("<" + key + ">", tag);
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
    }


    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

}
