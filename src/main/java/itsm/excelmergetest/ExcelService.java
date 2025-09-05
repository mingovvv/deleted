package itsm.excelmergetest;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.*;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;


@Slf4j
@Service
@RequiredArgsConstructor
public class ExcelService {


    public byte[] generateExcel(ClassPathResource attachedExcelFile) {

        try {

            log.info("generateExcel start");
            ClassPathResource file = new ClassPathResource("/static/request_infrastructure_excel_template.xlsx");

            try (ByteArrayOutputStream baos = new ByteArrayOutputStream(); InputStream is = file.getInputStream()) {

                JxlsHelper jxlsHelper = JxlsHelper.getInstance();
                jxlsHelper.processTemplate(is, baos, new Context());
                log.info("jxlsHelper processTemplate");

                return insertAllSheetsOfBIntoMiddleOfA(baos.toByteArray(), attachedExcelFile);

            }

        } catch (Exception e) {
            throw new RuntimeException("Excel 생성 실패", e);
        }

    }

    public byte[] downloadExcel() {
        ClassPathResource attachedExcelFile = new ClassPathResource("/static/attached.xlsx");
        return generateExcel(attachedExcelFile);
    }

    private static byte[] insertAllSheetsOfBIntoMiddleOfA(byte[] aBytes, Resource bExcel) throws IOException {

        try (XSSFWorkbook aWb = new XSSFWorkbook(new ByteArrayInputStream(aBytes));
             InputStream bIs = bExcel.getInputStream();
             XSSFWorkbook bWb = new XSSFWorkbook(bIs);
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            int aCount = aWb.getNumberOfSheets();

            // 가운데 탭 모두 제거(역순 삭제로 인덱스 보존)
            if (aCount >= 3) {
                for (int i = aCount - 2; i >= 1; i--) {
                    aWb.removeSheetAt(i);
                }
            }

            // B의 모든 시트를 1번 위치부터 추가
            for (int j = 0; j < bWb.getNumberOfSheets(); j++) {
                XSSFSheet src = bWb.getSheetAt(j);
                String newName = uniqueSheetName(aWb, src.getSheetName());
                XSSFSheet dest = aWb.createSheet(newName);
                copyEntireSheet(src, dest);
                aWb.setSheetOrder(newName, 1 + j); // 첫 시트(0) 뒤에 차례대로 삽입
            }

            aWb.setActiveSheet(0);
            aWb.setFirstVisibleTab(0);
            aWb.setForceFormulaRecalculation(true);

            aWb.write(out);
            return out.toByteArray();
        }
    }

    private static String uniqueSheetName(XSSFWorkbook wb, String base) {
        String safe = WorkbookUtil.createSafeSheetName(base);
        if (wb.getSheet(safe) == null) return safe;
        for (int i = 1; i < 1000; i++) {
            String cand = safe + " (" + i + ")";
            if (wb.getSheet(cand) == null) return cand;
        }
        return safe + " (" + System.currentTimeMillis() + ")";
    }

    private static void copyEntireSheet(XSSFSheet src, XSSFSheet dest) {
        XSSFWorkbook destWb = dest.getWorkbook();
        Map<Integer, XSSFCellStyle> styleMap = new HashMap<>();
        int maxCol = 0;

        for (Row r : src) {
            XSSFRow sRow = (XSSFRow) r;
            XSSFRow dRow = dest.createRow(sRow.getRowNum());
            dRow.setHeight(sRow.getHeight());

            for (Cell c : sRow) {
                XSSFCell sCell = (XSSFCell) c;
                XSSFCell dCell = dRow.createCell(sCell.getColumnIndex());

                XSSFCellStyle sStyle = sCell.getCellStyle();
                if (sStyle != null) {
                    int key = sStyle.getIndex();
                    XSSFCellStyle cached = styleMap.computeIfAbsent(key, k -> {
                        XSSFCellStyle cs = destWb.createCellStyle();
                        cs.cloneStyleFrom(sStyle);
                        return cs;
                    });
                    dCell.setCellStyle(cached);
                }

                switch (sCell.getCellType()) {
                    case STRING -> dCell.setCellValue(sCell.getRichStringCellValue());
                    case NUMERIC -> {
                        if (DateUtil.isCellDateFormatted(sCell)) {
                            dCell.setCellValue(sCell.getDateCellValue());
                        } else {
                            dCell.setCellValue(sCell.getNumericCellValue());
                        }
                    }
                    case BOOLEAN -> dCell.setCellValue(sCell.getBooleanCellValue());
                    case FORMULA -> dCell.setCellFormula(sCell.getCellFormula());
                    case BLANK -> dCell.setBlank();
                    case ERROR -> dCell.setCellErrorValue(sCell.getErrorCellValue());
                }
                maxCol = Math.max(maxCol, sCell.getColumnIndex());
            }
        }

        for (int i = 0; i < src.getNumMergedRegions(); i++) {
            dest.addMergedRegion(src.getMergedRegion(i));
        }
        for (int c = 0; c <= maxCol; c++) {
            dest.setColumnWidth(c, src.getColumnWidth(c));
        }
        dest.setDisplayGridlines(src.isDisplayGridlines());
        dest.setPrintGridlines(src.isPrintGridlines());
        dest.getPrintSetup().setLandscape(src.getPrintSetup().getLandscape());
    }

}
