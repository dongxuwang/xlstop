package cn.bytecode.tool.controller;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;

import static org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined.*;
import static org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined.LIME;

@Controller
public class IndexController {

    private static String UPLOADED_FOLDER = System.getProperty("user.home") + File.separator + "exceltool";

    @GetMapping("/")
    public String index() {
        return "index";
    }


    @PostMapping("/")
    public String singleFileUpload(@RequestParam("file") MultipartFile file,
                                 RedirectAttributes redirectAttributes,
                                 HttpServletResponse response) throws IOException {

        if (file.isEmpty()) {
            redirectAttributes.addFlashAttribute("message", "请选择Excel文件上传");
            return "redirect:status";
        }

        //Get the workbook instance for XLS file
        HSSFWorkbook workbook = new HSSFWorkbook(file.getInputStream());

        //Get first sheet from the workbook
        HSSFSheet sheet = workbook.getSheetAt(0);

        int firstRowNum = sheet.getFirstRowNum();
        int rowNum = sheet.getLastRowNum();
        int firstCellNum = sheet.getRow(1).getFirstCellNum();
        int columnNum = sheet.getRow(1).getLastCellNum() - 1;

        System.out.println(firstRowNum + " | " + rowNum + " | " + firstCellNum + " | " + columnNum);

        Double[][] values = new Double[rowNum][columnNum];

        for (int i = 1; i <= rowNum; ++i) {
            for (int j = 1; j <= columnNum; ++j) {
                Cell cell = sheet.getRow(i).getCell(j);
                if (cell.getCellTypeEnum() != CellType.NUMERIC) {
                    redirectAttributes.addFlashAttribute("message", "文件主体内容含有非数字内容");
                    return "redirect:status";
                }
                values[i - 1][j - 1] = cell.getNumericCellValue();
            }
        }


        Integer[][] sortedArrayindexes = new Integer[rowNum][columnNum];

        for (int i = 0; i < rowNum; ++i) {
            ArrayIndexComparator comparator = new ArrayIndexComparator(values[i]);
            Integer[] indexes = comparator.createIndexArray();
            sortedArrayindexes[i] = indexes;
            Arrays.sort(indexes, comparator);
        }

        HSSFCellStyle[] styles = createSortedBgColors(workbook);


        setSortedBgColors(sheet, sortedArrayindexes, styles);

        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-disposition", "attachment; filename="+ LocalDateTime.now(ZoneId.of("GMT+08:00")).format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"))+ "_" + file.getOriginalFilename());
        workbook.write(response.getOutputStream());

        return null;
    }

    private void setSortedBgColors(HSSFSheet sheet, Integer[][] sortedArrayindexes, HSSFCellStyle[] styles) {
        for (int i = 0; i < sortedArrayindexes.length; ++i) {
            for (int j = 0; j < styles.length; ++j) {
                sheet.getRow(i + 1).getCell(sortedArrayindexes[i][j] + 1).setCellStyle(styles[j]);
            }
        }
    }

    private HSSFCellStyle[] createSortedBgColors(HSSFWorkbook workbook) {
        HSSFCellStyle cellStyle1 = workbook.createCellStyle();  //创建设置EXCEL表格样式对象 cellStyle
        HSSFCellStyle cellStyle2 = workbook.createCellStyle();  //创建设置EXCEL表格样式对象 cellStyle
        HSSFCellStyle cellStyle3 = workbook.createCellStyle();  //创建设置EXCEL表格样式对象 cellStyle
        HSSFCellStyle cellStyle4 = workbook.createCellStyle();  //创建设置EXCEL表格样式对象 cellStyle
        HSSFCellStyle cellStyle5 = workbook.createCellStyle();  //创建设置EXCEL表格样式对象 cellStyle
        cellStyle1.setFillForegroundColor(RED.getIndex());// 设置背景色
        cellStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle2.setFillForegroundColor(ORANGE.getIndex());// 设置背景色
        cellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle3.setFillForegroundColor(YELLOW.getIndex());// 设置背景色
        cellStyle3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle4.setFillForegroundColor(SKY_BLUE.getIndex());// 设置背景色
        cellStyle4.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle5.setFillForegroundColor(LIME.getIndex());// 设置背景色
        cellStyle5.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return new HSSFCellStyle[]{cellStyle1, cellStyle2, cellStyle3, cellStyle4, cellStyle5};
    }

    @GetMapping("status")
    public String status() {
        return "status";
    }
}
