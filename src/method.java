import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;


import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;

import static org.apache.poi.ss.usermodel.CellType.*;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;

public class method extends Thread implements Runnable {

    File dirfile;
    File file;
    InputStream inputStream;
    Workbook loadworkbook;
    Workbook writeworkbook;
    Workbook scienceclassworkbook;
    Workbook artsclassworkbook;
    HSSFCellStyle hssfCellStyle;
    String[] sciencehead = new String[]{"姓名", "县", "校", "年级", "班次", "总分", "总分_校排名", "理综", "理综_校排名", "理综_班排名", "语文", "语文_校排名", "语文_班排名", "数学", "数学_校排名", "数学_班排名", "英语", "英语_校排名", "英语_班排名", "物理", "物理_校排名", "物理_班排名", "化学", "化学_校排名", "化学_班排名", "生物", "生物_校排名", "生物_班排名"};
    String[] artshead = new String[]{"姓名", "县", "校", "年级", "班次", "总分", "总分_校排名", "文综", "文综_校排名", "文综_班排名", "语文", "语文_校排名", "语文_班排名", "数学", "数学_校排名", "数学_班排名", "英语", "英语_校排名", "英语_班排名", "政治", "政治_校排名", "政治_班排名", "历史", "历史_校排名", "历史_班排名", "地理", "地理_校排名", "地理_班排名"};
    String[] sciencehead2 = new String[]{"姓名", "年级", "班次", "总分", "总分_校排名", "总分_班排名", "理综", "理综_校排名", "理综_班排名", "语文", "语文_校排名", "语文_班排名", "数学", "数学_校排名", "数学_班排名", "英语", "英语_校排名", "英语_班排名", "物理", "物理_校排名", "物理_班排名", "化学", "化学_校排名", "化学_班排名", "生物", "生物_校排名", "生物_班排名"};
    String[] artshead2 = new String[]{"姓名", "年级", "班次", "总分", "总分_校排名", "总分_班排名", "文综", "文综_校排名", "文综_班排名", "语文", "语文_校排名", "语文_班排名", "数学", "数学_校排名", "数学_班排名", "英语", "英语_校排名", "英语_班排名", "政治", "政治_校排名", "政治_班排名", "历史", "历史_校排名", "历史_班排名", "地理", "地理_校排名", "地理_班排名"};
    int[] scienceclass = new int[]{1, 3, 4, 6, 7, 9, 10, 11, 13, 14, 15, 16, 18, 19, 22, 24, 25, 26, 27};
    int[] artclass = new int[]{2, 5, 8, 11, 17, 20, 21};
    End end;

    public method(File file, End end) {
        this.end = end;
        this.file = file;
        dirfile = new File(file.getParentFile(), "分科成绩");
    }

    private void main() throws Exception {
        loadworkbook = new HSSFWorkbook(inputStream);
        writeworkbook = new HSSFWorkbook();
        scienceclassworkbook = new HSSFWorkbook();
        artsclassworkbook = new HSSFWorkbook();
        hssfCellStyle = style(writeworkbook);
        create(scienceclass, sciencehead, writeworkbook, loadworkbook, "理科");
        create(artclass, artshead, writeworkbook, loadworkbook, "文科");
        hssfCellStyle = style(scienceclassworkbook);
        setClassworkbook(sciencehead2, scienceclassworkbook, writeworkbook.getSheetAt(0), scienceclass);
        hssfCellStyle = style(artsclassworkbook);
        setClassworkbook(artshead2, artsclassworkbook, writeworkbook.getSheetAt(1), artclass);
        save("文理综分科.xls", writeworkbook);
        save("理科班级.xls", scienceclassworkbook);
        save("文科班级.xls", artsclassworkbook);
    }

    private void save(String name, Workbook workbook) throws Exception {
        FileOutputStream fileOutputStream = new FileOutputStream(new File(dirfile, name));
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    private void setClassworkbook(String[] head, Workbook classworkbook, Sheet loadsheet, int... className) {
        for (int serise : className) {
            int serial = 1;
            Sheet sheet = classworkbook.createSheet(Integer.toString(serise));
            for (int roms = 0; roms < loadsheet.getPhysicalNumberOfRows(); roms++) {
                if (roms == 0) {
                    HSSFRow hssfRow = (HSSFRow) sheet.createRow(0);
                    HSSFCellStyle hssfCellStyle = style(classworkbook);
                    Font font = classworkbook.createFont();
                    font.setFontHeight((short) 230);
                    font.setBold(true);
                    hssfCellStyle.setFont(font);
                    hssfRow.setHeight((short) 1920);
                    for (int i = 0; i < head.length; i++) {
                        HSSFCell hssfCell = hssfRow.createCell(i);
                        hssfCell.setCellStyle(hssfCellStyle);
                        hssfCell.setCellValue(head[i]);
                        sheet.setColumnWidth(i, 1360);
                        if (i == 0) {
                            sheet.setColumnWidth(i, 2660);
                        }
                    }
                } else {
                    if (loadsheet.getRow(roms).getCell(4).getNumericCellValue() == serise) {
                        HSSFRow hssfRow = (HSSFRow) sheet.createRow(serial);
                        HSSFRow loadhssfRow = (HSSFRow) loadsheet.getRow(roms);
                        hssfRow.setHeight((short) 400);
                        for (int i = 0; i < loadsheet.getRow(0).getPhysicalNumberOfCells(); i++) {
                            for (int i2 = 0; i2 < head.length; i2++) {
                                if (loadsheet.getRow(0).getCell(i).getStringCellValue().equals(head[i2])) {
                                    HSSFCell cell = hssfRow.createCell(i2);
                                    cell.setCellStyle(hssfCellStyle);
                                    if (loadhssfRow.getCell(i).getCellType() == STRING) {
                                        cell.setCellValue(loadhssfRow.getCell(i).getStringCellValue());
                                    }
                                    if (loadhssfRow.getCell(i).getCellType() == NUMERIC) {
                                        cell.setCellValue(loadhssfRow.getCell(i).getNumericCellValue());
                                    }
                                }
                            }
                        }
                        hssfRow.getCell(3).setCellValue(hssfRow.getCell(6).getNumericCellValue() + hssfRow.getCell(9).getNumericCellValue() + hssfRow.getCell(12).getNumericCellValue() + hssfRow.getCell(15).getNumericCellValue());
                        Cell cell = hssfRow.createCell(5);
                        cell.setCellStyle(hssfCellStyle);
                        cell.setCellValue(serial);
                        serial++;
                    }
                }
            }
        }
    }

    private void create(int[] c, String[] head, Workbook classworkbook, Workbook loadworkbook, String name) {
        int serial = 1;
        ArrayList<Student> arrayList = new ArrayList<>();
        Sheet sheet = classworkbook.createSheet(name);
        Sheet loadsheet = loadworkbook.getSheetAt(0);
        for (int roms = 0; roms < loadsheet.getPhysicalNumberOfRows(); roms++) {
            if (roms == 0) {
                HSSFRow hssfRow = (HSSFRow) sheet.createRow(0);
                HSSFCellStyle hssfCellStyle = style(classworkbook);
                Font font = classworkbook.createFont();
                font.setFontHeight((short) 230);
                font.setBold(true);
                hssfCellStyle.setFont(font);
                hssfRow.setHeight((short) 1920);
                for (int i = 0; i < head.length; i++) {
                    HSSFCell hssfCell = hssfRow.createCell(i);
                    hssfCell.setCellStyle(hssfCellStyle);
                    hssfCell.setCellValue(head[i]);
                    sheet.setColumnWidth(i, 1360);
                    if (i == 0) {
                        sheet.setColumnWidth(i, 2660);
                    }
                }
            } else {
                boolean b = false;
                for (int cl : c) if (cl == Math.ceil(loadsheet.getRow(roms).getCell(7).getNumericCellValue())) b = true;
                if (b) {
                    HSSFRow hssfRow = (HSSFRow) sheet.createRow(serial);
                    HSSFRow loadhssfRow = (HSSFRow) loadsheet.getRow(roms);
                    hssfRow.setHeight((short) 400);
                    for (int i = 0; i < loadsheet.getRow(0).getPhysicalNumberOfCells(); i++) {
                        for (int i2 = 0; i2 < head.length; i2++) {
                            if (loadsheet.getRow(0).getCell(i).getStringCellValue().equals(head[i2])) {
                                HSSFCell cell = hssfRow.createCell(i2);
                                cell.setCellStyle(hssfCellStyle);
                                if (loadhssfRow.getCell(i).getCellType() == STRING) {
                                    cell.setCellValue(loadhssfRow.getCell(i).getStringCellValue());
                                }
                                if (loadhssfRow.getCell(i).getCellType() == NUMERIC) {
                                    cell.setCellValue(loadhssfRow.getCell(i).getNumericCellValue());
                                }
                            }
                        }
                    }
                    hssfRow.getCell(5).setCellValue(hssfRow.getCell(7).getNumericCellValue() + hssfRow.getCell(10).getNumericCellValue() + hssfRow.getCell(13).getNumericCellValue() + hssfRow.getCell(16).getNumericCellValue());
                    Student student = new Student();
                    student.setSer(serial);
                    student.setAll((int) hssfRow.getCell(5).getNumericCellValue());
                    arrayList.add(student);
                    serial++;
                }
            }
        }
        Sheet sersheet = classworkbook.createSheet(name + "_排序");
        Collections.sort(arrayList, new Comparator<Student>() {
            @Override
            public int compare(Student o1, Student o2) {
                return o2.getAll() - o1.getAll();
            }
        });
        copyRow(classworkbook, sersheet, sheet.getRow(0), sersheet.createRow(0), true);
        for (int i = 0; i < arrayList.size(); i++) {
            copyRow(classworkbook, sersheet, sheet.getRow(arrayList.get(i).getSer()), sersheet.createRow(i + 1), true);
            sersheet.getRow(i + 1).getCell(6).setCellValue(i + 1);
        }
        for (int i = 0; i < sersheet.getRow(0).getPhysicalNumberOfCells(); i++) {
            sersheet.setColumnWidth(i, 1360);
            if (i == 0) sersheet.setColumnWidth(0, 2660);
        }
        writeworkbook.removeSheetAt(writeworkbook.getSheetIndex(name));
    }

    private HSSFCellStyle style(Workbook workbook) {
        Font font = workbook.createFont();
        font.setFontHeight((short) 200);
        HSSFCellStyle cellStyle = (HSSFCellStyle) workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setFont(font);
        cellStyle.setWrapText(true);
        return cellStyle;
    }

    public void copyRow(Workbook wb, Sheet sheet, Row sourceRow, Row distRow, boolean isCopyValue) {
        distRow.setHeight(sourceRow.getHeight());
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress sourceRange = sheet.getMergedRegion(i);
            if (sourceRange.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress distRange = new CellRangeAddress(distRow.getRowNum(),
                        distRow.getRowNum() + (sourceRange.getLastRow() - sourceRange.getFirstRow()),
                        sourceRange.getFirstColumn(), sourceRange.getLastColumn());
                sheet.addMergedRegion(distRange);
            }
        }
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            copyCell(wb, sourceRow.getCell(i), distRow.createCell(i), isCopyValue);
        }
    }

    private void copyCell(Workbook wb, Cell sourceCell, Cell distCell, boolean isCopyValue) {
        distCell.setCellStyle(sourceCell.getCellStyle());
        if (sourceCell.getCellComment() != null) {
            distCell.setCellComment(sourceCell.getCellComment());
        }
        CellType sourceCellType = sourceCell.getCellType();
        distCell.setCellType(sourceCellType);
        if (isCopyValue) {
            switch (sourceCellType) {
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(sourceCell)) {
                        distCell.setCellValue(sourceCell.getDateCellValue());
                    } else {
                        distCell.setCellValue(sourceCell.getNumericCellValue());
                    }
                    break;
                case STRING:
                    distCell.setCellValue(sourceCell.getRichStringCellValue());
                    break;
                case BOOLEAN:
                    distCell.setCellValue(sourceCell.getBooleanCellValue());
                    break;
                default:
                    break;
            }
        }
    }

    @Override
    public void run() {
        try {
            inputStream = new FileInputStream(file);
            dirfile.mkdirs();
            main();
            end.end();
        } catch (Exception exception) {
            exception.printStackTrace();
        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    static class Student {
        private int ser;
        public int all;

        public void setSer(int i) {
            ser = i;
        }

        public void setAll(int i) {
            all = i;
        }

        public int getAll() {
            return all;
        }

        public int getSer() {
            return ser;
        }
    }

    public interface End {
        void end();
    }

}
