package readData;

import jxl.*;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class JxlRead {
    public JxlRead() {}
    String str;
    String filename;
    int i;
    File file = new File("G:/Java/");
    File[] files = file.listFiles();
    /*
    * 读取 excel 文件的 A1 单元格
    * 注意读取是流的需要关闭
    * **/
    public void getExcel() {
        for (i = 0; i <files.length; i++) {
            //System.out.println(files[i].getName());
            try {
                InputStream is = new FileInputStream("G:/Java/" + files[i].getName());
                filename = files[i].getName();
                Workbook wb = Workbook.getWorkbook(is);
                Sheet sheet = wb.getSheet(0);
                Cell cell = sheet.getCell(0, 0);
                str = cell.getContents();
                if (cell.getType() == CellType.LABEL) {
                    LabelCell labelCell = (LabelCell) cell;
                    str = labelCell.getString();
                }
                is.close();
                renameFile("G:/Java/", filename + "", str + ".xls");
                //System.out.println(str);
            } catch (IOException e) {
                e.printStackTrace();
            } catch (BiffException e) {
                e.printStackTrace();
            }
        }
    }
    /*
    * 文件重命名
    * path：重命名路径
    * oldname：需要重命名的文件
    * newname：重命名后的文件
    * **/
    public void renameFile(String path, String oldname, String newname){
        if (!oldname.equals(newname)) {
            File oldfile = new File(path + oldname);
            // System.out.println("path + old " + oldfile);
            File newfile = new File(path + newname);
            // System.out.println("path + new " + newfile);
            // System.out.println(oldfile.renameTo(newfile));
            // 新建文件
            // try {
            //     newfile.createNewFile();
            // } catch (IOException e) {
            //     e.printStackTrace();
            // }
            //System.out.println(newfile + "111111");
            if (!oldfile.exists()) {
                return;
            }
            if (newfile.exists()) {
                System.out.println(newfile + "已存在");
            }else {
                oldfile.renameTo(newfile);
                System.out.println(oldfile.renameTo(newfile));
            }
        }else {
            System.out.println("文件名相同");
        }
    }

    public static void main(String[] args) {
        JxlRead jxlRead = new JxlRead();
        //jxlRead.getExcel("G:/Java/");
        jxlRead.getExcel();
        //System.out.println(jxlRead.str + ".xls");
        //jxlRead.readFile();
    }

}
