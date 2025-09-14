package com.example;

import com.intellij.psi.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;

import static com.example.CommonAction.generateLogicDescriptions;

public class CoreExcelExporter {

    private final Set<String> usedSheetNames = new HashSet<>();

    public void exportControllerExcel(File outputFile, List<PsiClass> services) throws Exception {
        Workbook workbook = new XSSFWorkbook();

        // 스타일 생성
        CellStyle grayHeaderStyle = createGrayHeaderStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);
        CellStyle categoryStyle = createCategoryStyle(workbook, true);
        CellStyle firstColumnCategoryStyle = createFirstColumnCategoryStyle(workbook);

        for (PsiClass serviceClazz : services) {
            for (PsiMethod method : serviceClazz.getMethods()) {
//                if (!isCoreMethod(method)) continue;

                // 시트 이름으로 API 이름 사용
                String sheetName = sanitizeSheetName(method.getName());
                Sheet sheet = workbook.createSheet(sheetName);
                int rowNum = 0;

                // 메인 헤더 생성 (A1:G4)
                createMainHeaders(sheet, grayHeaderStyle, dataStyle, rowNum, method);
                rowNum = 4;

                // Logic 설명 섹션
                rowNum = createLogicSection(sheet, method, rowNum, dataStyle, categoryStyle, firstColumnCategoryStyle);

                // 파라미터 섹션
                rowNum = createParameterSection(sheet, method, rowNum, dataStyle, categoryStyle);

                // 반환 타입 섹션
                rowNum = createReturnTypeSection(sheet, method, rowNum, dataStyle, categoryStyle);

                // 컬럼 너비 조정 (A~G열)
                for (int i = 0; i < 7; i++) {
                    sheet.autoSizeColumn(i);
                    if (sheet.getColumnWidth(i) < 2500) {
                        sheet.setColumnWidth(i, 2500);
                    }
                    if (sheet.getColumnWidth(i) > 8000) {
                        sheet.setColumnWidth(i, 8000);
                    }
                }
            }
        }

        // 파일 저장
        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            workbook.write(fos);
        }
        workbook.close();
    }

    private String sanitizeSheetName(String name) {
        String sanitized = name.replaceAll("[\\\\/?*\\[\\]]", "");
        if (sanitized.length() > 31) {
            sanitized = sanitized.substring(0, 31);
        }

        String uniqueName = sanitized;
        int counter = 1;
        while (usedSheetNames.contains(uniqueName)) {
            uniqueName = sanitized.substring(0, Math.min(28, sanitized.length())) + "_" + counter;
            counter++;
        }

        usedSheetNames.add(uniqueName);
        return uniqueName;
    }

    private void createMainHeaders(Sheet sheet, CellStyle headerStyle, CellStyle dataStyle, int startRow, PsiMethod method) {
        // 첫 번째 행: API Name
        Row row1 = sheet.createRow(startRow);
        createCell(row1, 0, "API Name", headerStyle);
        createCell(row1, 1, method.getName(), dataStyle);
        createEmptyCells(row1, 2, 5, dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(startRow, startRow, 1, 5));

        // 두 번째 행: 서비스명
        Row row2 = sheet.createRow(startRow + 1);
        createCell(row2, 0, "서비스명", headerStyle);
        createCell(row2, 1, getServiceName(method), dataStyle);
        createEmptyCells(row2, 2, 5, dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(startRow + 1, startRow + 1, 1, 5));

        // 세 번째 행: 기능
        Row row3 = sheet.createRow(startRow + 2);
        createCell(row3, 0, "기능", headerStyle);
        createCell(row3, 1, getDescription(method), dataStyle);
        createEmptyCells(row3, 2, 5, dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(startRow + 2, startRow + 2, 1, 5));

        // 네 번째 행: 상세내용
        Row row4 = sheet.createRow(startRow + 3);
        createCell(row4, 0, "상세내용", headerStyle);
        createCell(row4, 1, getDetailDescription(method), dataStyle);
        createEmptyCells(row4, 2, 5, dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(startRow + 3, startRow + 3, 1, 5));
    }

    private void createEmptyCells(Row row, int fromCol, int toCol, CellStyle style) {
        for (int col = fromCol; col <= toCol; col++) {
            createCell(row, col, "", style);
        }
    }

    private int createLogicSection(Sheet sheet, PsiMethod method, int startRow, CellStyle dataStyle, CellStyle categoryStyle, CellStyle firstColumnCategoryStyle) {
        int currentRow = startRow;

        // Logic 설명 헤더
        Row headerRow = sheet.createRow(currentRow++);
        createCell(headerRow, 0, "Logic 설명", categoryStyle);
        createEmptyCells(headerRow, 1, 5, firstColumnCategoryStyle);
        CellRangeAddress region = new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 5);
        sheet.addMergedRegion(region);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);

        // 밑에
        Row row1 = sheet.createRow(currentRow++);
        createCell(row1, 0, "번호", categoryStyle);
        createCell(row1, 1, "내용", categoryStyle);
        createCell(row1, 2, "비고", categoryStyle);
        createEmptyCells(row1, 3, 5, categoryStyle);
        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 2, 5));

        // MOS_CORE 모듈 메소드인 경우 로직 분석하여 생성
//        if (isCoreMethod(method)) {
            List<String> logicDescriptions = generateLogicDescriptions(method, false);
            
            for (int i = 0; i < Math.max(logicDescriptions.size(), 2); i++) {
                Row dataRow = sheet.createRow(currentRow++);
//                createCell(dataRow, 0, "", dataStyle);
                createCell(dataRow, 0, String.valueOf(i + 1), dataStyle);
                createCell(dataRow, 1, i < logicDescriptions.size() ? logicDescriptions.get(i) : "", dataStyle);
                createCell(dataRow, 2, "", dataStyle);
                createEmptyCells(dataRow, 3, 5, dataStyle);
                sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 2, 5));
            }
        return currentRow;
    }

    private int createParameterSection(Sheet sheet, PsiMethod method, int startRow, CellStyle dataStyle, CellStyle categoryStyle) {
        int currentRow = startRow;

        // 파라미터 헤더
        Row headerRow = sheet.createRow(currentRow++);
        createCell(headerRow, 0, "파라미터", categoryStyle);
        CellRangeAddress region = new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 5);
        sheet.addMergedRegion(region);
        // 테두리 스타일 적용
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);

        // 파라미터 헤더
        Row row1 = sheet.createRow(currentRow++);
//        createCell(headerRow, 0, "파라미터", categoryStyle);
        createCell(row1, 0, "속성", categoryStyle);
        createCell(row1, 1, "타입", categoryStyle);
        createCell(row1, 2, "필수여부", categoryStyle);
        createCell(row1, 3, "설명", categoryStyle);
        createCell(row1, 4, "옵션", categoryStyle);
        createCell(row1, 5, "비고", categoryStyle);

        // 파라미터 데이터
        String[] paramInfo = getParameterInfoDetailed(method);
        String[] properties = paramInfo[0].split("\n");
        String[] types = paramInfo[1].split("\n");
        String[] required = paramInfo[2].split("\n");
        String[] descriptions = paramInfo[3].split("\n");

        int maxParams = Math.max(1, Math.max(properties.length, Math.max(types.length, Math.max(required.length, descriptions.length))));

        for (int i = 0; i < Math.max(maxParams, 2); i++) { // 최소 3개 행
            Row dataRow = sheet.createRow(currentRow++);
//            createCell(dataRow, 0, "", dataStyle);
            createCell(dataRow, 0, i < properties.length ? properties[i] : "", dataStyle);
            createCell(dataRow, 1, i < types.length ? types[i] : "", dataStyle);
            createCell(dataRow, 2, i < required.length ? required[i] : "", dataStyle);
            createCell(dataRow, 3, i < descriptions.length ? descriptions[i] : "", dataStyle);
            createCell(dataRow, 4, "", dataStyle); // 옵션
            createCell(dataRow, 5, "", dataStyle); // 비고
        }

        return currentRow;
    }

    private int createReturnTypeSection(Sheet sheet, PsiMethod method, int startRow, CellStyle dataStyle, CellStyle categoryStyle) {
        int currentRow = startRow;

        // 요청예시 헤더
        Row headerRow = sheet.createRow(currentRow++);
        createCell(headerRow, 0, "요청예시", categoryStyle);
        createEmptyCells(headerRow, 1, 5, categoryStyle);
//        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6));
        CellRangeAddress region = new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 5);
        sheet.addMergedRegion(region);
        // 테두리 스타일 적용
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);

        // 반환 타입 헤더
        Row row1 = sheet.createRow(currentRow++);
//        createCell(headerRow, 0, "반환 타입", categoryStyle);
        createCell(row1, 0, "타입", categoryStyle);
        createCell(row1, 1, "Content", categoryStyle);
        createEmptyCells(row1, 2, 5, categoryStyle);
        createCell(row1, 3, "비고", categoryStyle);
        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 2, 5));

        // 반환 타입 데이터 (3개 행)
        for (int i = 0; i < 3; i++) {
            Row dataRow = sheet.createRow(currentRow++);
//            createCell(dataRow, 0, "", dataStyle);
            createCell(dataRow, 0, i == 0 ? getReturnType(method) : "", dataStyle);
            createCell(dataRow, 1, i == 0 ? "Content" : "", dataStyle);
            createEmptyCells(dataRow, 2, 5, dataStyle);
            createCell(dataRow, 5, "", dataStyle);
            sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 2, 5));
        }

        return currentRow;
    }

    // 유틸리티 메소드들
    private void createCell(Row row, int column, String value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    private String getServiceName(PsiMethod method) {
        PsiClass containingClass = method.getContainingClass();
        if (containingClass != null) {
            String className = containingClass.getName();
            if (className != null && className.endsWith("Controller")) {
                return className.replace("Controller", "Service");
            }
            return className != null ? className : "Unknown Service";
        }
        return "Unknown Service";
    }

    private String getDescription(PsiMethod method) {
        PsiAnnotation operationAnnotation = method.getAnnotation("io.swagger.v3.oas.annotations.Operation");
        if (operationAnnotation != null) {
            String summary = getAnnotationValue(operationAnnotation, "summary");
            if (!summary.isEmpty()) {
                return summary;
            }
        }
        return method.getName() + " 기능";
    }

    private String getDetailDescription(PsiMethod method) {
        PsiAnnotation operationAnnotation = method.getAnnotation("io.swagger.v3.oas.annotations.Operation");
        if (operationAnnotation != null) {
            String description = getAnnotationValue(operationAnnotation, "description");
            if (!description.isEmpty()) {
                return removeHtmlTags(description);
            }
        }
        return method.getName() + " 메소드의 상세 설명";
    }

    private String getReturnType(PsiMethod method) {
        PsiType returnType = method.getReturnType();
        if (returnType != null) {
            return returnType.getPresentableText();
        }
        return "void";
    }

    private String[] getParameterInfoDetailed(PsiMethod method) {
        List<String> properties = new ArrayList<>();
        List<String> types = new ArrayList<>();
        List<String> required = new ArrayList<>();
        List<String> descriptions = new ArrayList<>();

        PsiParameter[] parameters = method.getParameterList().getParameters();
        for (PsiParameter param : parameters) {
            // Skip HttpServletRequest, HttpServletResponse, etc.
            String typeName = param.getType().getPresentableText();
            if (typeName.contains("HttpServlet") || typeName.contains("Principal") || 
                typeName.contains("Authentication")) {
                continue;
            }

            properties.add(param.getName() != null ? param.getName() : "");
            types.add(typeName);
            
            // Check if parameter is required (has @RequestParam(required=true) or no @RequestParam at all for @PathVariable)
            boolean isRequired = true;
            PsiAnnotation reqParam = param.getAnnotation("org.springframework.web.bind.annotation.RequestParam");
            if (reqParam != null) {
                String requiredValue = getAnnotationValue(reqParam, "required");
                isRequired = !requiredValue.equals("false");
            }
            required.add(isRequired ? "필수" : "선택");
            
            descriptions.add(param.getName() != null ? param.getName() + " 파라미터" : "");
        }

        return new String[]{
            String.join("\n", properties),
            String.join("\n", types),
            String.join("\n", required),
            String.join("\n", descriptions)
        };
    }

    /**
     * MOS_CORE 모듈의 메소드인지 확인
     * 파일 경로에 MOS_CORE가 포함되어 있는지 검사
     */
    private boolean isCoreMethod(PsiMethod method) {
        PsiClass containingClass = method.getContainingClass();
        if (containingClass == null) return false;

        PsiFile containingFile = containingClass.getContainingFile();
        if (containingFile == null) return false;

        String filePath = containingFile.getVirtualFile() != null ?
            containingFile.getVirtualFile().getPath() : "";

        return filePath.contains("core");
    }

    private String getAnnotationValue(PsiAnnotation annotation, String attributeName) {
        if (annotation == null) return "";

        PsiAnnotationMemberValue value = annotation.findDeclaredAttributeValue(attributeName);
        if (value == null) return "";

        if (value instanceof PsiLiteralExpression) {
            Object literalValue = ((PsiLiteralExpression) value).getValue();
            return literalValue != null ? literalValue.toString() : "";
        }

        String rawText = value.getText();

        // 문자열 정리
        rawText = rawText
                .replaceAll("^\"|\"$", "")       // 양쪽 따옴표 제거
                .replace(" + ", "")              // 문자열 연결 연산자 제거
                .replace("\\n", "\n")            // \n → 실제 개행
                .replace("\\\"", "\"")           // \" → "
                .replaceAll("<[^>]*>", "")       // 모든 HTML 태그 제거 (예: <b>, <pre>)
                .trim();                         // 앞뒤 공백 제거

        // 줄 단위로 정리: 의미 없는 공백 줄 제거
        StringBuilder cleaned = new StringBuilder();
        for (String line : rawText.split("\n")) {
            String trimmed = line.trim();
            if (!trimmed.isEmpty()) {
                cleaned.append(trimmed).append("\n");
            }
        }

        return cleaned.toString().trim();
    }

    // HTML 태그 제거 유틸리티 메소드
    private String removeHtmlTags(String text) {
        if (text == null || text.isEmpty()) {
            return text;
        }
        // HTML 태그 제거 (< > 사이의 모든 내용)
        return text.replaceAll("\"", "")
                .replaceAll("<[^>]*>", "")
                .replaceAll("(?m)^\\s*\\n", "")
                .trim();
    }

    private CellStyle createGrayHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setWrapText(true);
        return style;
    }

    private CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setWrapText(true);
        return style;
    }

    private CellStyle createCategoryStyle(Workbook workbook, boolean isTrue) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        if (isTrue) style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setWrapText(true);
        return style;
    }

    // 1번째 열 (A열)용 카테고리 스타일
    private CellStyle createFirstColumnCategoryStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setWrapText(true);
        return style;
    }

    // 1번째 열 (A열)용 데이터 스타일
    private CellStyle createFirstColumnDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setWrapText(true);
        return style;
    }
}
