package readData;

import jxl.*;
import jxl.read.biff.BiffException;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class JxlRead {
    public JxlRead() {}

    public void getExcel(String filePath) {
        try {
            InputStream is = new FileInputStream(filePath);
            Workbook wb = Workbook.getWorkbook(is);
            Sheet sheet = wb.getSheet(0);
            Cell cell = sheet.getCell(0,0);
            String str = cell.getContents();
            if (cell.getType() == CellType.LABEL) {
                LabelCell labelCell = (LabelCell)cell;
                str = labelCell.getString();
            }
            System.out.println(str);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        JxlRead jxlRead = new JxlRead();
        jxlRead.getExcel("C:\\Users\\yu\\Desktop\\word\\文档\\中国统计年鉴 2017(EXCEL+目录)\\CH0102.xls");
    }

}
