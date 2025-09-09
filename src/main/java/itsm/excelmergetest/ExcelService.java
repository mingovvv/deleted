package itsm.excelmergetest;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObjectData;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
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

                return insertAllSheetsOfBIntoMiddleOfA2(baos.toByteArray(), attachedExcelFile);
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

//                copyEntireSheet(src, dest);
                copyEntireSheet2(src, dest);
                copyPictureAndCharts(src, dest);

                aWb.setSheetOrder(newName, 1 + j);
            }

            aWb.setActiveSheet(0);
            aWb.setFirstVisibleTab(0);
            aWb.setForceFormulaRecalculation(true);

            aWb.write(out);
            return out.toByteArray();
        }
    }

    private static byte[] insertAllSheetsOfBIntoMiddleOfA2(byte[] aBytes, Resource bExcel) throws IOException {
        try (XSSFWorkbook aWb = new XSSFWorkbook(new ByteArrayInputStream(aBytes));
             InputStream bIs = bExcel.getInputStream();
             XSSFWorkbook bWb = new XSSFWorkbook(bIs);
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            int aCount = aWb.getNumberOfSheets();
            if (aCount >= 3) {
                for (int i = aCount - 2; i >= 1; i--) {
                    aWb.removeSheetAt(i);
                }
            }

            try (SXSSFWorkbook sxssf = new SXSSFWorkbook(aWb, 100, true, true)) {
                for (int j = 0; j < bWb.getNumberOfSheets(); j++) {
                    XSSFSheet src = bWb.getSheetAt(j);
                    String newName = uniqueSheetName(aWb, src.getSheetName());

                    SXSSFSheet dest = sxssf.createSheet(newName);

                    copyEntireSheet3(src, dest, 1000, 50);
                    copyPicture(src, dest);

                    sxssf.getXSSFWorkbook().setSheetOrder(newName, 1 + j);
                }

                sxssf.getXSSFWorkbook().setActiveSheet(0);
                sxssf.getXSSFWorkbook().setFirstVisibleTab(0);
                sxssf.getXSSFWorkbook().setForceFormulaRecalculation(true);

                sxssf.write(out);
                sxssf.dispose();
            }

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

    private static void copyEntireSheet3(XSSFSheet src, SXSSFSheet dest, int flushEveryN, int keepRows) {
        final XSSFWorkbook baseWb = dest.getWorkbook().getXSSFWorkbook();
        final Map<Short, CellStyle> styleCache = new HashMap<>(256);

        int rowsSinceFlush = 0;
        int maxCol = -1;

        final int firstRow = Math.max(0, src.getFirstRowNum());
        final int lastRow  = src.getLastRowNum();

        for (int r = firstRow; r <= lastRow; r++) {
            XSSFRow sRow = src.getRow(r);
            if (sRow == null) continue;

            Row dRow = dest.createRow(r);

            dRow.setHeight(sRow.getHeight());
            try {
                dRow.setZeroHeight(sRow.getZeroHeight());
            } catch (Throwable ignore) { }

            short firstCol = (sRow.getFirstCellNum() >= 0) ? sRow.getFirstCellNum() : 0;
            short lastCol  = sRow.getLastCellNum();
            if (lastCol < 0) continue;
            if (lastCol - 1 > maxCol) maxCol = lastCol - 1;

            for (int c = firstCol; c < lastCol; c++) {
                XSSFCell sCell = sRow.getCell(c, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if (sCell == null) continue;

                Cell dCell = dRow.createCell(c);

                XSSFCellStyle sStyle = sCell.getCellStyle();
                if (sStyle != null) {
                    short idx = sStyle.getIndex();
                    CellStyle cached = styleCache.get(idx);
                    if (cached == null) {
                        CellStyle cs = baseWb.createCellStyle();
                        cs.cloneStyleFrom(sStyle);
                        styleCache.put(idx, cs);
                        cached = cs;
                    }
                    dCell.setCellStyle(cached);
                }

                switch (sCell.getCellType()) {
                    case STRING -> dCell.setCellValue(sCell.getStringCellValue());
                    case NUMERIC -> dCell.setCellValue(sCell.getNumericCellValue());
                    case BOOLEAN -> dCell.setCellValue(sCell.getBooleanCellValue());
                    case FORMULA -> dCell.setCellFormula(sCell.getCellFormula());
                    case ERROR -> dCell.setCellErrorValue(sCell.getErrorCellValue());
                    default -> {}
                }
            }

            rowsSinceFlush++;
            if (flushEveryN > 0 && rowsSinceFlush >= flushEveryN) {
                try {
                    dest.flushRows(Math.max(0, keepRows));
                } catch (IOException e) {
                    throw new RuntimeException("flushRows failed", e);
                }
                rowsSinceFlush = 0;
            }
        }

        for (int i = 0; i < src.getNumMergedRegions(); i++) {
            dest.addMergedRegion(src.getMergedRegion(i));
        }

        for (int c = 0; c <= maxCol; c++) {
            dest.setColumnWidth(c, src.getColumnWidth(c));
        }
        dest.setDisplayGridlines(false);
        dest.setPrintGridlines(src.isPrintGridlines());
        dest.getPrintSetup().setLandscape(src.getPrintSetup().getLandscape());
    }

    private static void copyEntireSheet2(XSSFSheet src, XSSFSheet dest) {
        final XSSFWorkbook destWb = dest.getWorkbook();

        final Map<Short, XSSFCellStyle> styleCache = new HashMap<>(256);

        int maxCol = -1;

        final int firstRow = Math.max(0, src.getFirstRowNum());
        final int lastRow  = src.getLastRowNum();

        for (int r = firstRow; r <= lastRow; r++) {
            XSSFRow sRow = src.getRow(r);
            if (sRow == null) continue;

            XSSFRow dRow = dest.getRow(r);
            if (dRow == null) dRow = dest.createRow(r);

            dRow.setHeight(sRow.getHeight());
            dRow.setZeroHeight(sRow.getZeroHeight());

            short firstCol = (sRow.getFirstCellNum() >= 0) ? sRow.getFirstCellNum() : 0;
            short lastCol  = sRow.getLastCellNum();
            if (lastCol < 0) continue;
            if (lastCol - 1 > maxCol) maxCol = lastCol - 1;

            for (int c = firstCol; c < lastCol; c++) {
                XSSFCell sCell = sRow.getCell(c, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if (sCell == null) continue;

                XSSFCell dCell = dRow.getCell(c);
                if (dCell == null) dCell = dRow.createCell(c);

                XSSFCellStyle sStyle = sCell.getCellStyle();
                if (sStyle != null) {
                    short styleIdx = sStyle.getIndex();
                    XSSFCellStyle cached = styleCache.get(styleIdx);
                    if (cached == null) {
                        cached = destWb.createCellStyle();
                        cached.cloneStyleFrom(sStyle);
                        styleCache.put(styleIdx, cached);
                    }
                    dCell.setCellStyle(cached);
                }

                switch (sCell.getCellType()) {
                    case STRING:
                        dCell.setCellValue(sCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        dCell.setCellValue(sCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        dCell.setCellValue(sCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        dCell.setCellFormula(sCell.getCellFormula());
                        break;
                    case ERROR:
                        dCell.setCellErrorValue(sCell.getErrorCellValue());
                        break;
                    default:
                }
            }
        }

        for (int i = 0; i < src.getNumMergedRegions(); i++) {
            dest.addMergedRegion(src.getMergedRegion(i));
        }
        for (int c = 0; c <= maxCol; c++) {
            dest.setColumnWidth(c, src.getColumnWidth(c));
        }
        dest.setDisplayGridlines(false);
        dest.setPrintGridlines(src.isPrintGridlines());
        dest.getPrintSetup().setLandscape(src.getPrintSetup().getLandscape());
    }

    private static void copyPictureAndCharts(XSSFSheet src, XSSFSheet dst) {
        XSSFDrawing srcDrawing = src.getDrawingPatriarch();
        if (srcDrawing == null) return;

        XSSFDrawing dstDrawing = dst.createDrawingPatriarch();
        Map<String, Integer> pictureIndexCache = new HashMap<>(64);

        List<XSSFChart> charts = srcDrawing.getCharts();
        int chartIdx = 0;

        for (CTTwoCellAnchor ctAnchor : srcDrawing.getCTDrawing().getTwoCellAnchorList()) {
            if (ctAnchor.isSetPic()) {
                copyPicture(ctAnchor.getFrom(), ctAnchor.getTo(), ctAnchor.getPic(),
                        srcDrawing, dstDrawing, dst.getWorkbook(), pictureIndexCache);
                continue;
            }
            if (ctAnchor.isSetGraphicFrame()) {
                CTGraphicalObjectData gd = ctAnchor.getGraphicFrame().getGraphic().getGraphicData();
                if ("http://schemas.openxmlformats.org/drawingml/2006/chart".equals(gd.getUri())) {
                    if (chartIdx < charts.size()) {
                        copyChart(ctAnchor.getFrom(), ctAnchor.getTo(), charts.get(chartIdx++), dstDrawing);
                    }
                }
            }
        }
    }

    private static void copyPicture(XSSFSheet src, SXSSFSheet dst) {
        XSSFDrawing srcDrawing = src.getDrawingPatriarch();
        if (srcDrawing == null) return;

        Drawing<?> dstDrawing = dst.createDrawingPatriarch();
        Workbook wb = dst.getWorkbook();

        for (CTTwoCellAnchor anchor : srcDrawing.getCTDrawing().getTwoCellAnchorList()) {
            if (!anchor.isSetPic()) continue;

            XSSFClientAnchor dstAnchor = new XSSFClientAnchor(
                    (int)((long)anchor.getFrom().getColOff()),
                    (int)((long)anchor.getFrom().getRowOff()),
                    (int)((long)anchor.getTo().getColOff()),
                    (int)((long)anchor.getTo().getRowOff()),
                    anchor.getFrom().getCol(), anchor.getFrom().getRow(),
                    anchor.getTo().getCol(),   anchor.getTo().getRow()
            );

            PictureData pd = extractPictureData(anchor.getPic(), srcDrawing);
            if (pd == null || pd.bytes == null) continue;

            int picIdx = wb.addPicture(pd.bytes, pd.poiType);
            dstDrawing.createPicture(dstAnchor, picIdx);
        }
    }

    private static void copyPicture(CTMarker from, CTMarker to, CTPicture ctPic, XSSFDrawing srcDrawing, XSSFDrawing dstDrawing, Workbook dstWb, Map<String, Integer> pictureIndexCache) {
        XSSFClientAnchor anchor = new XSSFClientAnchor(
                (int) ((long) from.getColOff()), (int) ((long) from.getRowOff()),
                (int) ((long) to.getColOff()),   (int) ((long) to.getRowOff()),
                from.getCol(), from.getRow(),
                to.getCol(),   to.getRow()
        );

        PictureData pd = extractPictureData(ctPic, srcDrawing);
        if (pd == null || pd.bytes == null) return;

        String key = pd.poiType + ":" + pd.bytes.length + ":" + Arrays.hashCode(pd.bytes);
        int idx = pictureIndexCache.computeIfAbsent(key, k -> dstWb.addPicture(pd.bytes, pd.poiType));
        dstDrawing.createPicture(anchor, idx);
    }

    private static void copyChart(CTMarker from, CTMarker to, XSSFChart srcChart, XSSFDrawing dstDrawing) {
        XSSFClientAnchor anchor = new XSSFClientAnchor(
                (int) ((long) from.getColOff()), (int) ((long) from.getRowOff()),
                (int) ((long) to.getColOff()),   (int) ((long) to.getRowOff()),
                from.getCol(), from.getRow(),
                to.getCol(),   to.getRow()
        );
        XSSFChart dstChart = dstDrawing.createChart(anchor);

        dstChart.getCTChartSpace().set(srcChart.getCTChartSpace());
        dstChart.getCTChart().set(srcChart.getCTChart());
    }

    private static PictureData extractPictureData(CTPicture ctPic, XSSFDrawing srcDrawing) {
        try {
            String rId = ctPic.getBlipFill().getBlip().getEmbed();
            if (rId == null) return null;

            PackageRelationship rel = srcDrawing.getPackagePart().getRelationship(rId);
            if (rel == null) return null;

            PackagePart part = srcDrawing.getPackagePart().getRelatedPart(rel);
            if (part == null) return null;

            try (InputStream is = part.getInputStream()) {
                byte[] bytes = is.readAllBytes();
                int poiType = detectPoiPictureType(part);
                return new PictureData(bytes, poiType);
            }
        } catch (Exception e) {
            log.warn("extractPictureData failed", e);
            return null;
        }
    }

    private static int detectPoiPictureType(PackagePart part) {
        String ct = part.getContentType();
        if (ct == null) return Workbook.PICTURE_TYPE_PNG;

        switch (ct) {
            case "image/jpeg":
            case "image/jpg":
                return Workbook.PICTURE_TYPE_JPEG;
            case "image/png":
                return Workbook.PICTURE_TYPE_PNG;
            default:
                return Workbook.PICTURE_TYPE_PNG;
        }
    }

    private static class PictureData {
        final byte[] bytes;
        final int poiType;
        PictureData(byte[] bytes, int poiType) {
            this.bytes = bytes;
            this.poiType = poiType;
        }
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
