package com.example;

import com.intellij.openapi.project.Project;
import com.intellij.psi.*;
import com.intellij.psi.search.GlobalSearchScope;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;

import static com.example.CommonAction.generateLogicDescriptions;
import static com.example.CommonAction.generateLogicDescriptionsBiz;

public class BizExcelExporter {

    private final Set<String> usedSheetNames = new HashSet<>();

    public void exportControllerExcel(File outputFile, List<PsiClass> controllers) throws Exception {
        Workbook workbook = new XSSFWorkbook();

        // 스타일 생성
        CellStyle grayHeaderStyle = createGrayHeaderStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);
        CellStyle categoryStyle = createCategoryStyle(workbook, true);
        CellStyle nonBottomCategoryStyle = createCategoryStyle(workbook, false);
        CellStyle firstColumnCategoryStyle = createFirstColumnCategoryStyle(workbook);
        CellStyle firstColumnDataStyle = createFirstColumnDataStyle(workbook);

        for (PsiClass clazz : controllers) {
            for (PsiMethod method : clazz.getMethods()) {
                if (!isApiMethod(method)) continue;

                // 시트 이름으로 API 이름 사용
                String sheetName = sanitizeSheetName(method.getName());
                Sheet sheet = workbook.createSheet(sheetName);
                int rowNum = 0;

                // 메인 헤더 생성 (A1:I2)
                createMainHeaders(sheet, grayHeaderStyle, dataStyle, rowNum, method);
                rowNum = 1;

                // API 기본 정보 섹션
                rowNum = createApiBasicInfoSection(sheet, method, clazz, rowNum, dataStyle, categoryStyle, firstColumnCategoryStyle, firstColumnDataStyle);

                // Java Class Layer 섹션
                rowNum = createJavaClassLayerSection(sheet, method, clazz, rowNum, dataStyle, categoryStyle, firstColumnCategoryStyle, firstColumnDataStyle);

                // Logic 설명 섹션
                rowNum = createLogicSection(sheet, method, rowNum, dataStyle, categoryStyle, firstColumnCategoryStyle);

                // 파라미터 섹션
                rowNum = createParameterSection(sheet, method, rowNum, dataStyle, categoryStyle, firstColumnCategoryStyle);

                // 요청예시 섹션
                rowNum = createRequestExampleSection(sheet, method, clazz, rowNum, dataStyle, categoryStyle, nonBottomCategoryStyle, firstColumnCategoryStyle);

                // Response JSON 섹션
                rowNum = createResponseSection(sheet, method, rowNum, dataStyle, nonBottomCategoryStyle, firstColumnCategoryStyle);


                // 컬럼 너비 조정 (A~I열)
                for (int i = 0; i < 9; i++) {
                    sheet.autoSizeColumn(i);
                    if (sheet.getColumnWidth(i) < 2500) {
                        sheet.setColumnWidth(i, 2500);
                    }
                    if (sheet.getColumnWidth(i) > 8000) {
                        sheet.setColumnWidth(i, 6000);
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
        // 첫 번째 헤더 행
        Row row0 = sheet.createRow(startRow);
        createCell(row0, 0, "목록", null);

        // 첫 번째 헤더 행
        Row row1 = sheet.createRow(startRow);
        createCell(row1, 0, "API Name", headerStyle);
        createCell(row1, 1, method.getName(), dataStyle);
        createCell(row1, 2, "", dataStyle);
        createCell(row1, 3, "", dataStyle);
        createCell(row1, 4, "Http Method", headerStyle);
        createCell(row1, 5, getHttpMethod(method), dataStyle);
        createCell(row1, 6, "", dataStyle);

        sheet.addMergedRegion(new CellRangeAddress(startRow, startRow, 1, 3)); // API Name
        sheet.addMergedRegion(new CellRangeAddress(startRow, startRow, 5, 6)); // Http Method

    }
    private void createEmptyCells(Row row, int fromCol, int toCol, CellStyle style) {
        for (int col = fromCol; col <= toCol; col++) {
            createCell(row, col, "", style);
        }
    }

    private int createApiBasicInfoSection(Sheet sheet, PsiMethod method, PsiClass clazz, int startRow, CellStyle dataStyle, CellStyle categoryStyle, CellStyle firstColumnCategoryStyle, CellStyle firstColumnDataStyle) {
        int currentRow = startRow;

        // URL
        Row row0 = sheet.createRow(currentRow++);
        createCell(row0, 0, "URL", firstColumnCategoryStyle);
        createCell(row0, 1, getUrlPath(method, clazz), dataStyle);
        createEmptyCells(row0, 2, 6, dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 1, 6));

        // API 설명 행
        Row row01 = sheet.createRow(currentRow++);
        createCell(row01, 0, "요구사항 ID", firstColumnCategoryStyle);
        createCell(row01, 1, "", dataStyle);
        createEmptyCells(row01, 2, 6, dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 1, 6));

        // API 설명 행
        Row row1 = sheet.createRow(currentRow++);
        createCell(row1, 0, "API 설명", firstColumnCategoryStyle);
        createCell(row1, 1, getDescription(method), dataStyle);
        createEmptyCells(row1, 2, 6, dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 1, 6));

        // 행 높이 자동 조정 흉내: 텍스트 줄 수 기준 수동 조정
        String descriptionText = getDescription(method);
        int lines = descriptionText.split("\n").length;
        float lineHeight = sheet.getDefaultRowHeightInPoints(); // 보통 15.0
        row1.setHeightInPoints(lines * lineHeight); // 줄 수 × 기본 높이

        // API 상세 설명 행
        Row row2 = sheet.createRow(currentRow++);
        createCell(row2, 0, "API 상세 설명", firstColumnCategoryStyle);
        createCell(row2, 1, getDetailDescription(method), dataStyle);
        createEmptyCells(row2, 2, 6, dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 1, 6));

        // 행 높이 자동 조정 흉내: 텍스트 줄 수 기준 수동 조정
        String descriptionText2 = getDescription(method);
        int lines2 = descriptionText2.split("\n").length;
        float lineHeight2 = sheet.getDefaultRowHeightInPoints(); // 보통 15.0
        row2.setHeightInPoints(lines2 * lineHeight2); // 줄 수 × 기본 높이

        return currentRow;
    }

    private int createJavaClassLayerSection(Sheet sheet, PsiMethod method, PsiClass clazz, int startRow, CellStyle dataStyle, CellStyle categoryStyle, CellStyle firstColumnCategoryStyle, CellStyle firstColumnDataStyle) {
        int currentRow = startRow;

        // Java Class Layer 헤더
        Row headerRow = sheet.createRow(currentRow++);
        createCell(headerRow, 0, "Java Class Layer", firstColumnCategoryStyle);
        createEmptyCells(headerRow, 1, 6, firstColumnCategoryStyle);
//        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6));
        CellRangeAddress region2 = new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6);
        sheet.addMergedRegion(region2);
        // 테두리 스타일 적용
        RegionUtil.setBorderRight(BorderStyle.THIN, region2, sheet);

        // Controller Class / Method
        Row row1 = sheet.createRow(currentRow++);
        createCell(row1, 0, "Controller Class / Method", firstColumnCategoryStyle);
        createCell(row1, 1, clazz.getName() + " / " + method.getName(), dataStyle);
        createEmptyCells(row1, 2, 6, dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 1, 6));

        // Service Class / Method
        Row row2 = sheet.createRow(currentRow++);
        createCell(row2, 0, "Service Class / Method", firstColumnCategoryStyle);
        createCell(row2, 1, getServiceInfo(method), dataStyle);
        createEmptyCells(row2, 2, 6, dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 1, 6));

        // Repository Class / Method
        Row row3 = sheet.createRow(currentRow++);
        createCell(row3, 0, "Repository Class", firstColumnCategoryStyle);
        createCell(row3, 1, getRepositoryInfo(method), dataStyle);
        createEmptyCells(row3, 2, 6, dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 1, 6));

        return currentRow;
    }
//
//    private int createLogicSection(Sheet sheet, PsiMethod method, int startRow, CellStyle dataStyle, CellStyle categoryStyle) {
//        int currentRow = startRow;
//
//        // Logic 설명 헤더
//        Row headerRow = sheet.createRow(currentRow++);
//        createCell(headerRow, 0, "Logic 설명", categoryStyle);
//        createCell(headerRow, 1, "로직단계유형", categoryStyle);
//        createCell(headerRow, 2, "메서드", categoryStyle);
//        createCell(headerRow, 3, "Core 클래스", categoryStyle);
//        createCell(headerRow, 4, "Core 메소드", categoryStyle);
//        for (int i = 5; i < 9; i++) {
//            createCell(headerRow, i, "", categoryStyle);
//        }
//
////        // Logic 데이터 행들
////        for (int i = 0; i < 5; i++) { // 5개 행 생성
////            Row dataRow = sheet.createRow(currentRow++);
////            createCell(dataRow, 0, "", dataStyle);
////            createCell(dataRow, 1, i == 0 ? "Controller, Service" : "", dataStyle);
////            createCell(dataRow, 2, i == 0 ? method.getName() + "()" : "", dataStyle);
////            createCell(dataRow, 3, i == 0 ? getCoreClass(method) : "", dataStyle);
////            createCell(dataRow, 4, i == 0 ? getCoreMethod(method) : "", dataStyle);
////            for (int j = 5; j < 9; j++) {
////                createCell(dataRow, j, "", dataStyle);
////            }
////        }
//
//        List<String> logicDescriptions = generateLogicDescriptions(method);
//
//        for (int i = 0; i < Math.max(logicDescriptions.size(), 3); i++) {
//            Row dataRow = sheet.createRow(currentRow++);
//            createCell(dataRow, 0, i < logicDescriptions.size() ? logicDescriptions.get(i) : "", dataStyle);
//            createCell(dataRow, 1, String.valueOf(i + 1), dataStyle); // TODO : Service 클래스라면 Service, Controller 클래스라면 Controller
//            createCell(dataRow, 2, "", dataStyle); // TODO : 메서드 클래스명
//            createCell(dataRow, 3, "", dataStyle); // TODO : 호출 클래스에 CORE라는 패키지명이 포함되면
//            createCell(dataRow, 34, "", dataStyle); // TODO : 호출 클래스의 메서드에 CORE라는 패키지명이 포함되면
//            createEmptyCells(dataRow, 5, 6, dataStyle);// TODO : 메서드 클래스에 CORE라는 패키지명이 포함되면
//            sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 2, 5));
//        }
//
//        return currentRow;
//    }
private String returnPropertMethod(PsiMethod method) {
    String methodName = method.getName().toLowerCase(); // 소문자 변환

    if (methodName.contains("realdelete")) {
        return "realdeleteEntities";
    } else if (methodName.contains("undelete")) {
        return "undeleteEntities";
    } else if (methodName.contains("delete")) {
        return "deleteEntities";
    } else if (methodName.contains("create")) {
        return "createEntities";
    } else if (methodName.contains("update")) {
        return "updateEntities";
    } else {
        return "getCustomQueryPredicates";
    }
}


private int createLogicSection(Sheet sheet, PsiMethod method, int startRow, CellStyle dataStyle, CellStyle categoryStyle, CellStyle firstColumnCategoryStyle) {
    int currentRow = startRow;

    // Logic 설명 헤더
    Row headerRow = sheet.createRow(currentRow++);
    createCell(headerRow, 0, "Logic 설명", categoryStyle);
    createEmptyCells(headerRow, 1, 6, firstColumnCategoryStyle);
//    sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6));
    CellRangeAddress region = new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6);
    sheet.addMergedRegion(region);
    // 테두리 스타일 적용
    RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);

    Row row1 = sheet.createRow(currentRow++);
    createCell(row1, 0, "로직단계유형", categoryStyle);
    createCell(row1, 1, "메서드", categoryStyle);
    createCell(row1, 2, "Core 클래스", categoryStyle);
    createCell(row1, 3, "Core 메소드", categoryStyle);
    for (int i = 4; i < 6; i++) {
        createCell(row1, i, "", categoryStyle);
    }
    createCell(row1, 6, "Description", categoryStyle);

    sheet.setColumnWidth(2, 20 * 256); // 열 인덱스 2 (즉, "Core 클래스")
    sheet.setColumnWidth(3, 30 * 256); // 열 인덱스 3 (즉, "Core 메소드")

    List<String> logicDescriptions = generateLogicDescriptionsBiz(method);
    String classType = getClassType(method); // Controller 또는 Service
    String methodClassName = method.getContainingClass() != null ? method.getContainingClass().getName() : "";

    for (int i = 0; i < Math.max(logicDescriptions.size(), 3); i++) {
        Row dataRow = sheet.createRow(currentRow++);
        createCell(dataRow, 0, i < logicDescriptions.size() ? logicDescriptions.get(i) : "", dataStyle);
        createCell(dataRow, 1, i==0 ? "Controller" : i==1 ? "Service" : "", dataStyle);
        createCell(dataRow, 3, i==0 ? method.getName() : i==1 ? returnPropertMethod(method) : "", dataStyle);
//        createCell(dataRow, 2, i==0 ? "Controller" : "Service", methodClassName + "." + method.getName() + "()", dataStyle);

        String coreClass = "";
        String coreMethod = "";

        // CORE 패키지명 여부 확인
        if (isInCorePackage(method)) {
            coreClass = methodClassName;
            coreMethod = method.getName() + "()";
        }

        createCell(dataRow, 2, i==1 ? "CrudService" : coreClass, dataStyle);
//        createCell(dataRow, 4, i==1? returnPropertMethod(method) : "", dataStyle);

        createEmptyCells(dataRow, 4, 6, dataStyle); // 나머지 빈 셀
////        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 2, 5));
//        CellRangeAddress region2 = new CellRangeAddress(currentRow - 1, currentRow - 1, 2, 5);
//        sheet.addMergedRegion(region2);
//        // 테두리 스타일 적용
//        RegionUtil.setBorderRight(BorderStyle.THIN, region2, sheet);
    }

    return currentRow;
}
    private String getClassType(PsiMethod method) {
        PsiClass containingClass = method.getContainingClass();
        if (containingClass != null) {
            String qualifiedName = containingClass.getQualifiedName();
            if (qualifiedName != null) {
                if (qualifiedName.toLowerCase().contains("controller")) {
                    return "Controller";
                } else if (qualifiedName.toLowerCase().contains("service")) {
                    return "Service";
                }
            }
        }
        return "기타";
    }

    private boolean isInCorePackage(PsiMethod method) {
        PsiClass containingClass = method.getContainingClass();
        if (containingClass != null) {
            String qualifiedName = containingClass.getQualifiedName();
            return qualifiedName != null && qualifiedName.toLowerCase().contains("core");
        }
        return false;
    }


    private int createParameterSection(Sheet sheet, PsiMethod method, int startRow, CellStyle dataStyle, CellStyle categoryStyle, CellStyle firstColumnCategoryStyle) {
        int currentRow = startRow;

        // 파라미터 헤더
        Row headerRow = sheet.createRow(currentRow++);
        createCell(headerRow, 0, "파라미터", categoryStyle);
        CellRangeAddress region = new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6);
        sheet.addMergedRegion(region);
        // 테두리 스타일 적용
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);

        Row row1 = sheet.createRow(currentRow++);
        createCell(row1, 0, "속성", categoryStyle);
        createCell(row1, 1, "타입", categoryStyle);
        createCell(row1, 2, "필수여부", categoryStyle);
        createCell(row1, 3, "설명", categoryStyle);
        for (int i = 4; i < 6; i++) {
            createCell(row1, i, "", categoryStyle);
        }
        createCell(row1, 6, "Description", categoryStyle);

        // 파라미터 데이터
        String[] paramInfo = getParameterInfoDetailed(method);
        String[] properties = paramInfo[0].split("\n");
        String[] types = paramInfo[1].split("\n");
        String[] required = paramInfo[2].split("\n");
        String[] descriptions = paramInfo[3].split("\n");
        String[] detailDescriptions = paramInfo[4].split("\n");

        int maxParams = Math.max(1, Math.max(properties.length, Math.max(types.length, Math.max(required.length, descriptions.length))));

        for (int i = 0; i < Math.max(maxParams, 3); i++) { // 최소 3개 행
            Row dataRow = sheet.createRow(currentRow++);
//            createCell(dataRow, 0, "", dataStyle);
            createCell(dataRow, 0, i < properties.length ? properties[i] : "", dataStyle);
            createCell(dataRow, 1, i < types.length ? types[i] : "", dataStyle);
            createCell(dataRow, 2, i < required.length ? required[i] : "", dataStyle);
            createCell(dataRow, 3, i < descriptions.length ? i==0 ?descriptions[i] : properties[i].equals("saveHist") ? "이력 저장 여부" : "" : "", dataStyle);
//            createCell(dataRow, 4, i < descriptions.length ? returnSemiDescriptions() : "", dataStyle);
            for (int j = 4; j < 7; j++) {
                createCell(dataRow, j, "", dataStyle);
            }
//            createCell(dataRow, 6, i < detailDescriptions.length ? detailDescriptions[i] : "", dataStyle);
        }

        return currentRow;
    }

    private int createRequestExampleSection(Sheet sheet, PsiMethod method, PsiClass clazz, int startRow, CellStyle dataStyle, CellStyle categoryStyle, CellStyle nonBottomCategoryStyle, CellStyle firstColumnCategoryStyle) {
        int currentRow = startRow;

        // 요청예시 헤더
        Row headerRow = sheet.createRow(currentRow++);
        createCell(headerRow, 0, "요청예시", categoryStyle);
        createEmptyCells(headerRow, 1, 6, categoryStyle);
//        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6));
        CellRangeAddress region = new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6);
        sheet.addMergedRegion(region);
        // 테두리 스타일 적용
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);

        // 요청예시 데이터 (여러 행으로 분할)
        String requestExample = getRequestExample(method, clazz);
        String[] exampleLines = requestExample.split("\n");

        for (int i = 0; i < exampleLines.length ; i++) { // 최소 3개 행
            Row dataRow = sheet.createRow(currentRow++);
            createCell(dataRow, 0, i < exampleLines.length ? exampleLines[i] : "", dataStyle);
            createEmptyCells(headerRow, 1, 6, dataStyle);
//            sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6));
            CellRangeAddress region2 = new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6);
            sheet.addMergedRegion(region2);
            // 테두리 스타일 적용
            RegionUtil.setBorderRight(BorderStyle.THIN, region2, sheet);
        }

        return currentRow;
    }

    private int createResponseSection(Sheet sheet, PsiMethod method, int startRow, CellStyle dataStyle, CellStyle categoryStyle, CellStyle firstColumnCategoryStyle) {
        int currentRow = startRow;

        // Response JSON 헤더
        Row headerRow = sheet.createRow(currentRow++);
        createCell(headerRow, 0, "Response JSON", categoryStyle);
        createEmptyCells(headerRow, 1, 6, categoryStyle);
//        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6));
        CellRangeAddress region = new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6);
        sheet.addMergedRegion(region);
        // 테두리 스타일 적용
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);

        Row row1 = sheet.createRow(currentRow++);
        createCell(row1, 0, "Element", categoryStyle);
        createCell(row1, 1, "Type", categoryStyle);
        createCell(row1, 2, "Content", categoryStyle);
        for (int i = 3; i < 6; i++) {
            createCell(row1, i, "", categoryStyle);
        }
        createCell(row1, 6, "Description", categoryStyle);

        // Response 데이터
        for (int i = 0; i < 3; i++) { // 3개 행
            Row dataRow = sheet.createRow(currentRow++);
//            createCell(dataRow, 0, "", dataStyle);
            createCell(dataRow, 0, i == 0 ? getResponseElement(method) : "", dataStyle);
            createCell(dataRow, 1, i == 0 ? getResponseType(method) : "", dataStyle);
            createCell(dataRow, 2, i == 0 ? getResponseContent(method) : "", dataStyle);
            for (int j = 2; j < 6; j++) {
                createCell(dataRow, j, "", dataStyle);
            }
            createCell(dataRow, 6, i == 0 ? getResponseDescription(method) : "", dataStyle);
        }

        // Response JSON Sample 헤더
        Row sampleHeaderRow = sheet.createRow(currentRow++);
        createCell(sampleHeaderRow, 0, "Response JSON Sample", categoryStyle);
        createEmptyCells(sampleHeaderRow, 1, 6, categoryStyle);
//        sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6));
        CellRangeAddress region2 = new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6);
        sheet.addMergedRegion(region2);
        // 테두리 스타일 적용
        RegionUtil.setBorderRight(BorderStyle.THIN, region2, sheet);

        // Response JSON Sample 데이터
        String jsonSample = getResponseJsonSample(method);
        String[] sampleLines = jsonSample.split("\n");

        for (int i = 0; i < Math.max(sampleLines.length, 1); i++) { // 최소 3개 행
            Row dataRow = sheet.createRow(currentRow++);
            createCell(dataRow, 0, i < sampleLines.length ? sampleLines[i] : "", dataStyle);
            createEmptyCells(dataRow, 1, 6, dataStyle);
//            sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6));
            CellRangeAddress region3 = new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 6);
            sheet.addMergedRegion(region3);
            // 테두리 스타일 적용
            RegionUtil.setBorderRight(BorderStyle.THIN, region3, sheet);
        }

        return currentRow;
    }

    // 나머지 메소드들은 기존과 동일하게 유지
    private void createCell(Row row, int column, String value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    private boolean isApiMethod(PsiMethod method) {
        String[] apiAnnotations = {
                "org.springframework.web.bind.annotation.RequestMapping",
                "org.springframework.web.bind.annotation.GetMapping",
                "org.springframework.web.bind.annotation.PostMapping",
                "org.springframework.web.bind.annotation.PutMapping",
                "org.springframework.web.bind.annotation.DeleteMapping",
                "RequestMapping",
                "GetMapping",
                "PostMapping",
                "PutMapping",
                "DeleteMapping"
        };

        return hasAnyAnnotation(method, apiAnnotations);
    }

    private boolean hasAnnotation(PsiMethod method, String annotationFqn) {
        PsiAnnotation annotation = method.getAnnotation(annotationFqn);
        return annotation != null;
    }

    private boolean hasAnyAnnotation(PsiMethod method, String... annotationFqns) {
        PsiModifierList modifierList = method.getModifierList();
        for (PsiAnnotation annotation : modifierList.getAnnotations()) {
            String qName = annotation.getQualifiedName();
            if (qName != null && Arrays.asList(annotationFqns).contains(qName)) {
                return true;
            }
        }
        return false;
    }

    private String getApiName(PsiMethod method) {
        PsiAnnotation operationAnnotation = method.getAnnotation("io.swagger.v3.oas.annotations.Operation");
        if (operationAnnotation != null) {
            String summary = getAnnotationValue(operationAnnotation, "summary");
            if (!summary.isEmpty()) {
                return summary;
            }
        }
        return method.getName();
    }

    private String getHttpMethod(PsiMethod method) {
        if (hasAnyAnnotation(method, "org.springframework.web.bind.annotation.GetMapping")) return "GET";
        if (hasAnyAnnotation(method, "org.springframework.web.bind.annotation.PostMapping")) return "POST";
        if (hasAnyAnnotation(method, "org.springframework.web.bind.annotation.PutMapping")) return "PUT";
        if (hasAnyAnnotation(method, "org.springframework.web.bind.annotation.DeleteMapping")) return "DELETE";

        PsiAnnotation requestMapping = method.getAnnotation("org.springframework.web.bind.annotation.RequestMapping");
        PsiAnnotation requestMapping2 = method.getAnnotation("RequestMapping");
        if (requestMapping != null ) {
            String methodValue = getAnnotationValue(requestMapping, "method");
            if (!methodValue.isEmpty()) {
                return methodValue.replace("RequestMethod.", "");
            }
        }
        if (requestMapping2 != null ) {
            String methodValue = getAnnotationValue(requestMapping2, "method");
            if (!methodValue.isEmpty()) {
                return methodValue.replace("RequestMethod.", "");
            }
        }
        return "GET"; // 기본값
    }

    private String getUrlPath(PsiMethod method, PsiClass clazz) {
        StringBuilder path = new StringBuilder();

        // 클래스 레벨 경로
        PsiAnnotation classRequestMapping = clazz.getAnnotation("org.springframework.web.bind.annotation.RequestMapping");
        if (classRequestMapping != null) {
            String classPath = getAnnotationValue(classRequestMapping, "value");
            if (!classPath.isEmpty()) {
                path.append(classPath);
            }
        }

        // 클래스 레벨 경로
        PsiAnnotation classRequestMapping2 = clazz.getAnnotation("RequestMapping");
        if (classRequestMapping2 != null) {
            String classPath = getAnnotationValue(classRequestMapping2, "value");
            if (!classPath.isEmpty()) {
                path.append(classPath);
            }
        }

        // 메소드 레벨 경로
        String methodPath = "";

        PsiAnnotation requestMapping = method.getAnnotation("org.springframework.web.bind.annotation.RequestMapping");
        PsiAnnotation requestMapping2 = method.getAnnotation("RequestMapping");
        if (requestMapping != null) {
            methodPath = getAnnotationValue(requestMapping, "value");
        } else if (requestMapping2 != null){
            methodPath = getAnnotationValue(requestMapping2, "value");
        }else {
            // 다른 매핑 어노테이션들 확인
            String[] mappingAnnotations = {
                    "org.springframework.web.bind.annotation.GetMapping",
                    "org.springframework.web.bind.annotation.PostMapping",
                    "org.springframework.web.bind.annotation.PutMapping",
                    "org.springframework.web.bind.annotation.DeleteMapping"
            };

            for (String annotationFqn : mappingAnnotations) {
                PsiAnnotation annotation = method.getAnnotation(annotationFqn);
                if (annotation != null) {
                    methodPath = getAnnotationValue(annotation, "value");
                    break;
                }
            }
        }

        path.append(methodPath);
        return path.toString();
    }

    private String getDescription(PsiMethod method) {
        PsiAnnotation operationAnnotation = method.getAnnotation("io.swagger.v3.oas.annotations.Operation");
        if (operationAnnotation != null) {
            String summary = getAnnotationValue(operationAnnotation, "summary");
            if (!summary.isEmpty()) {
                return summary;
            }
        }
        PsiAnnotation operationAnnotation2 = method.getAnnotation("Operation");
        if (operationAnnotation2 != null) {
            String summary = getAnnotationValue(operationAnnotation2, "summary");
            if (!summary.isEmpty()) {
                return summary;
            }
        }
        return method.getName() + " API";
    }

    private String getDetailDescription(PsiMethod method) {
        PsiAnnotation operationAnnotation = method.getAnnotation("io.swagger.v3.oas.annotations.Operation");
        if (operationAnnotation != null) {
            String description = getAnnotationValue(operationAnnotation, "description");
            if (!description.isEmpty()) {
                return removeHtmlTags(description);
            }
        }
        PsiAnnotation operationAnnotation2 = method.getAnnotation("Operation");
        if (operationAnnotation2 != null) {
            String description = getAnnotationValue(operationAnnotation2, "description");
            if (!description.isEmpty()) {
                return removeHtmlTags(description);
            }
        }
        return method.getName() + " 메소드의 상세 설명";
    }

    private String getServiceInfo(PsiMethod method) {
        String fallbackMethodName = method.getName();
        PsiClass containingClass = method.getContainingClass();

        if (containingClass == null || method.getBody() == null) {
            return "Service / " + fallbackMethodName;
        }

        Map<String, String> serviceFields = new HashMap<>();

        // 1. 클래스 내 @Autowired 또는 이름에 'Service'가 포함된 필드 수집
        for (PsiField field : containingClass.getAllFields()) {
            boolean isAutowired = field.hasAnnotation("Autowired");
            boolean isLikelyService = field.getName() != null && field.getName().toLowerCase().contains("service");

            if (isAutowired || isLikelyService) {
                String fieldName = field.getName(); // 예: cdsCodeClsService
                String fieldType = field.getType().getPresentableText(); // 예: CdsCodeClsService
                if (fieldName != null && fieldType != null) {
                    serviceFields.put(fieldName, fieldType);
                }
            }
        }

        // 2. 메서드 내에서 서비스 객체가 호출하는 메서드 추출
        AtomicReference<String> serviceInfo = new AtomicReference<>();

        method.accept(new JavaRecursiveElementVisitor() {
            @Override
            public void visitMethodCallExpression(PsiMethodCallExpression expression) {
                super.visitMethodCallExpression(expression);
                PsiExpression qualifier = expression.getMethodExpression().getQualifierExpression();
                if (qualifier instanceof PsiReferenceExpression) {
                    String qualifierName = ((PsiReferenceExpression) qualifier).getReferenceName();
                    if (qualifierName != null && serviceFields.containsKey(qualifierName)) {
                        String serviceClass = serviceFields.get(qualifierName);
                        String calledMethod = expression.getMethodExpression().getReferenceName();
                        serviceInfo.set(serviceClass + " / " + calledMethod);
                    }
                }
            }
        });

        return serviceInfo.get() != null ? serviceInfo.get() : "Service / " + fallbackMethodName;
    }

    private String getRepositoryInfo(PsiMethod controllerMethod) {
        String controllerMethodName = controllerMethod.getName();
        PsiClass controllerClass = controllerMethod.getContainingClass();
        if (controllerClass == null || controllerMethod.getBody() == null) {
            return "Repository / " + controllerMethodName;
        }

        Map<String, String> serviceFields = new HashMap<>();

        // 1. Controller의 Service 필드 수집
        for (PsiField field : controllerClass.getAllFields()) {
            boolean isAutowired = field.hasAnnotation("Autowired");
            boolean isLikelyService = field.getName() != null && field.getName().toLowerCase().contains("service");
            if (isAutowired || isLikelyService) {
                serviceFields.put(field.getName(), field.getType().getCanonicalText()); // cdsCodeClsService -> com.example.CdsCodeClsService
            }
        }
        List<String> repositoryInfoList = new ArrayList<>();

        try {
            controllerMethod.accept(new JavaRecursiveElementVisitor() {
                @Override
                public void visitMethodCallExpression(PsiMethodCallExpression callExpr) {
                    super.visitMethodCallExpression(callExpr);
                    PsiExpression qualifier = callExpr.getMethodExpression().getQualifierExpression();
                    if (qualifier instanceof PsiReferenceExpression) {
                        PsiElement resolvedQualifier = ((PsiReferenceExpression) qualifier).resolve();
                        if (resolvedQualifier instanceof PsiVariable) {
                            PsiType qualifierType = ((PsiVariable) resolvedQualifier).getType();
                            if (qualifierType instanceof PsiClassType) {
                                PsiClass qualifierClass = ((PsiClassType) qualifierType).resolve();
                                if (qualifierClass != null) {
                                    for (PsiField serviceField : qualifierClass.getAllFields()) {
                                        if (serviceField.hasAnnotation("Autowired") ||
                                                serviceField.getName().toLowerCase().contains("repository")) {
                                            String entry = serviceField.getName();
                                            repositoryInfoList.add(entry);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            });
        } catch (FoundRepositoryInfoException ignored) {
        }

        if (!repositoryInfoList.isEmpty()) {
            return String.join(", ", repositoryInfoList);
        } else {
            return "Repository / " + controllerMethodName;
        }
    }

//    // TODO 중요!!!!!!!!!!!!!!!!!!!!!!!!!!!
//    private String methodCall(PsiMethod controllerMethod) {
//        String controllerMethodName = controllerMethod.getName();
//        PsiClass controllerClass = controllerMethod.getContainingClass();
//
//        Map<String, String> serviceFields = new HashMap<>();
//
//        // 1. Controller의 Service 필드 수집
//        for (PsiField field : controllerClass.getAllFields()) {
//            boolean isAutowired = field.hasAnnotation("Autowired");
//            boolean isLikelyService = field.getName() != null && field.getName().toLowerCase().contains("service");
//            if (isAutowired || isLikelyService) {
//                serviceFields.put(field.getName(), field.getType().getCanonicalText()); // cdsCodeClsService -> com.example.CdsCodeClsService
//            }
//        }
//
//        // 2. Controller method 내에서 Service 메서드 호출 분석
//        final String[] repositoryInfo = {null};
//
//        try {
//            controllerMethod.accept(new JavaRecursiveElementVisitor() {
//                @Override
//                public void visitMethodCallExpression(PsiMethodCallExpression callExpr) {
//                    super.visitMethodCallExpression(callExpr);
//                    PsiExpression qualifier = callExpr.getMethodExpression().getQualifierExpression();
//                    if (!(qualifier instanceof PsiReferenceExpression)) return;
//
//                    String serviceFieldName = ((PsiReferenceExpression) qualifier).getReferenceName();
//                    if (!serviceFields.containsKey(serviceFieldName)) return;
//
//                    // 3. Resolve Service method
//                    PsiElement resolvedElement = callExpr.getMethodExpression().resolve();
//                    if (!(resolvedElement instanceof PsiMethod)) return;
//                    PsiMethod resolvedServiceMethod = (PsiMethod) resolvedElement;
//                    // TODO 이 서비스클래스의 패키지 경로에 core 라는 단어가 포함이 돼있으면
//                    PsiClass serviceClass = resolvedServiceMethod.getContainingClass();
//                    if (serviceClass == null) return;
//
//                    // 4. Service 클래스 내 Repository 필드 수집 (생성자 주입 or @Autowired)
//                    Map<String, String> repositoryFields = new HashMap<>();
//                    for (PsiField serviceField : serviceClass.getAllFields()) {
//                        if (serviceField.hasAnnotation("Autowired") || serviceField.getName().toLowerCase().contains("repository")) {
//                            repositoryFields.put(serviceField.getName(), serviceField.getType().getPresentableText());
//                        }
//                    }
//
//                    // 5. Service 메서드 내부에서 호출된 Repository 식별
//                    resolvedServiceMethod.accept(new JavaRecursiveElementVisitor() {
//                        @Override
//                        public void visitMethodCallExpression(PsiMethodCallExpression repoCall) {
//                            super.visitMethodCallExpression(repoCall);
//                            PsiExpression repoQualifier = repoCall.getMethodExpression().getQualifierExpression();
//                            if (repoQualifier instanceof PsiReferenceExpression) {
//                                String repoFieldName = ((PsiReferenceExpression) repoQualifier).getReferenceName();
//                                if (repositoryFields.containsKey(repoFieldName)) {
//                                    String repoClass = repositoryFields.get(repoFieldName);
//                                    String repoMethod = repoCall.getMethodExpression().getReferenceName();
//                                    repositoryInfo[0] = repoClass + " / " + repoMethod;
//                                    throw new FoundRepositoryInfoException(); // 빠른 중단
//                                }
//                            }
//                        }
//                    });
//                }
//            });
//        } catch (FoundRepositoryInfoException ignored) {}
//
//        return repositoryInfo[0] != null ? repositoryInfo[0] : "Repository / " + controllerMethodName;
//    }

    // 빠른 탐색 종료를 위한 예외
    private static class FoundRepositoryInfoException extends RuntimeException {}


    private String getCoreClass(PsiMethod method) {
        // TODO : 개선 - CONTROLLER 가 호출하는 sERVICE 클래스가 호출하는 메소드중 메소드 경로 중 `core` 이 포함된 클래스명을 반환
        return "CoreClass";
    }

    private String getCoreMethod(PsiMethod method) {
        // TODO : 개선 - CONTROLLER 가 호출하는 sERVICE 클래스가 호출하는 메소드중 메소드 경로 중 `core` 이 포함된 클래스가 호출하는 메소드를 반환
        return "coreMethod()";
    }

    private String[] getParameterInfoDetailed(PsiMethod method) {
        StringBuilder properties = new StringBuilder();
        StringBuilder types = new StringBuilder();
        StringBuilder required = new StringBuilder();
        StringBuilder descriptions = new StringBuilder();
        StringBuilder detailDescriptions = new StringBuilder();

        PsiParameter[] parameters = method.getParameterList().getParameters();

        for (int i = 0; i < parameters.length; i++) {
            PsiParameter param = parameters[i];
            if (i > 0) {
                properties.append("\n");
                types.append("\n");
                required.append("\n");
                descriptions.append("\n");
                detailDescriptions.append("\n");
            }

            String paramName = param.getName();
            String paramType = param.getType().getPresentableText();
            boolean isRequired = false;
            String description = "";

            // 어노테이션 정보 확인
            PsiAnnotation requestParam = param.getAnnotation("org.springframework.web.bind.annotation.RequestParam");
            if (requestParam != null) {
                String value = getAnnotationValue(requestParam, "value");
                if (!value.isEmpty()) paramName = value;

                String requiredValue = getAnnotationValue(requestParam, "required");
                isRequired = !"false".equals(requiredValue);
                description = "요청 파라미터";
            }

            PsiAnnotation pathVariable = param.getAnnotation("org.springframework.web.bind.annotation.PathVariable");
            if (pathVariable != null) {
                String value = getAnnotationValue(pathVariable, "value");
                if (!value.isEmpty()) paramName = value;
                isRequired = true;
                description = "경로 변수";
            }

            PsiAnnotation requestBody = param.getAnnotation("org.springframework.web.bind.annotation.RequestBody");
            if (requestBody != null) {
                isRequired = true;
                description = "요청 본문";
            }

            properties.append(paramName != null ? paramName : "");
            types.append(paramType);
            required.append(isRequired ? "Y" : "N");
            descriptions.append(description);
            detailDescriptions.append(paramType + " 타입의 " + description);
        }

        return new String[]{
                properties.toString(),
                types.toString(),
                required.toString(),
                descriptions.toString(),
                detailDescriptions.toString()
        };
    }
    private String getRequestExample(PsiMethod method, PsiClass clazz) {
        String httpMethod = getHttpMethod(method);
        String url = getUrlPath(method, clazz);
        StringBuilder queryParams = new StringBuilder();

        StringBuilder example = new StringBuilder();
        example.append(httpMethod).append(" http://localhost:8080").append(url);

        PsiParameter[] parameters = method.getParameterList().getParameters();

        // 쿼리 파라미터 처리
        for (PsiParameter param : parameters) {
            if (param.getAnnotation("org.springframework.web.bind.annotation.RequestParam") != null ||
                    param.getAnnotation("RequestParam") != null) {
                if (queryParams.length() == 0) {
                    queryParams.append("?");
                } else {
                    queryParams.append("&");
                }
                queryParams.append(param.getName()).append("=").append("true"); // 기본값을 true로 설정
            }
        }

        // 쿼리 파라미터 추가
        example.append(queryParams);

        // Content-Type 헤더만 출력 (본문 없음)
        if ("POST".equals(httpMethod) || "PUT".equals(httpMethod) || "PATCH".equals(httpMethod)) {
            example.append("\nContent-Type: application/json");
        }

        return example.toString();
    }

    private String getResponseType(PsiMethod method) {
        if(method.getName().equals("getCustomQueryPredicates")) {
            return "List";
        } else {
            return "String";
        }
    }
    private String getResponseElement(PsiMethod method) {
        if(method.getName().equals("getCustomQueryPredicates")) {
            return "ResponseEntity";
        } else {
            return "return String";
        }
    }
    private String getResponseContent(PsiMethod method) {
        if(method.getName().equals("getCustomQueryPredicates")) {
            return "객체 Object 배열";
        } else {
            return "Success";
        }
    }

    private String getResponseDescription(PsiMethod method) {
        PsiType returnType = method.getReturnType();
        if (returnType == null) {
            return "응답 본문 없음";
        }

        String returnTypeName = returnType.getPresentableText();
        if (returnTypeName.contains("ResponseEntity")) {
            return "HTTP 상태 코드와 응답 본문을 포함한 응답";
        } else if (returnTypeName.equals("String")) {
            return "문자열 메시지";
        } else if (returnTypeName.equals("void")) {
            return "응답 본문 없음";
        } else {
            return returnTypeName + " 타입의 객체 정보";
        }
    }

    private String getResponseJsonSample(PsiMethod method) {
        PsiType returnType = method.getReturnType();
        if (returnType == null) {
            return "HTTP 200 OK (본문 없음)";
        }

        String returnTypeName = returnType.getPresentableText();

        if (returnTypeName.contains("ResponseEntity")) {
//            return "{\n  \"status\": \"success\",\n  \"data\": {},\n  \"message\": \"처리 완료\"\n}";
            return "Success";
        } else if (returnTypeName.equals("String")) {
            return "\"처리가 완료되었습니다.\"";
        } else if (returnTypeName.equals("void")) {
            return "HTTP 200 OK (본문 없음)";
        } else {
            return "{\n  // " + returnTypeName + " 객체의 JSON 표현\n}";
        }
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

    // 1번째 열 (A열)용 카테고리 스타일 - 오른쪽 굵은 테두리
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

    // 1번째 열 (A열)용 데이터 스타일 - 오른쪽 굵은 테두리
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
    private List<String> generateLogicDescriptionsRecursive(PsiMethod method, Set<PsiMethod> visited) {
        List<String> descriptions = new ArrayList<>();

        if (method == null || visited.contains(method)) return descriptions;
        visited.add(method);

        method.accept(new JavaRecursiveElementVisitor() {
            @Override
            public void visitMethodCallExpression(PsiMethodCallExpression expression) {
                super.visitMethodCallExpression(expression);

                PsiMethod calledMethod = (PsiMethod) expression.getMethodExpression().resolve();
                if (calledMethod == null) return;

                PsiClass containingClass = calledMethod.getContainingClass();
                if (containingClass == null) return;

                String className = containingClass.getName();
                String methodName = calledMethod.getName();

                String line = "Call: " + className + "." + methodName + "()";
                descriptions.add(line);

                // core 포함 서비스 로직만 재귀 추적
                String qualifiedName = containingClass.getQualifiedName();
                if (qualifiedName != null && qualifiedName.contains("core")) {
                    descriptions.addAll(generateLogicDescriptionsRecursive(calledMethod, visited));
                }
            }
        });

        return descriptions;
    }


}
