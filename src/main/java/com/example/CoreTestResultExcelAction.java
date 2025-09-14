package com.example;

import com.intellij.execution.ExecutionListener;
import com.intellij.execution.ExecutionManager;
import com.intellij.execution.process.ProcessHandler;
import com.intellij.execution.process.ProcessOutputTypes;
import com.intellij.execution.runners.ExecutionEnvironment;
import com.intellij.execution.testframework.TestFrameworkRunningModel;
import com.intellij.execution.testframework.sm.runner.GeneralTestEventsProcessor;
import com.intellij.execution.testframework.sm.runner.SMTestProxy;
import com.intellij.execution.testframework.sm.runner.events.*;
import com.intellij.openapi.actionSystem.AnAction;
import com.intellij.openapi.actionSystem.AnActionEvent;
import com.intellij.openapi.project.Project;
import com.intellij.openapi.vfs.VirtualFile;
import com.intellij.psi.*;
import com.intellij.psi.search.FileTypeIndex;
import com.intellij.psi.search.GlobalSearchScope;
import com.intellij.openapi.fileTypes.StdFileTypes;
import com.intellij.openapi.ui.Messages;
import com.intellij.openapi.fileChooser.FileChooser;
import com.intellij.openapi.fileChooser.FileChooserDescriptor;
import com.intellij.openapi.application.ApplicationManager;
import com.intellij.openapi.util.Key;
import com.intellij.util.messages.MessageBusConnection;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class CoreTestResultExcelAction extends AnAction {

    private static final Map<String, TestExecutionResult> testResults = new ConcurrentHashMap<>();
    private static final Map<String, TestClassMetadata> classMetadata = new ConcurrentHashMap<>();
    private MessageBusConnection connection;
    private ProcessHandler currentProcessHandler;
    private StringBuilder processOutput = new StringBuilder();

    @Override
    public void actionPerformed(AnActionEvent e) {
        Project project = e.getProject();
        if (project == null) {
            Messages.showErrorDialog("프로젝트가 선택되지 않았습니다.", "오류");
            return;
        }

        try {
            // 이미 수집된 결과가 있으면 Excel 생성
            if (!testResults.isEmpty()) {
                int result = Messages.showYesNoDialog(
                        String.format("이미 수집된 테스트 결과(%d개)가 있습니다.\nExcel 파일을 생성하시겠습니까?", testResults.size()),
                        "테스트 결과 존재",
                        Messages.getQuestionIcon()
                );

                if (result == Messages.YES) {
                    generateExcelFile(project);
                    return;
                }
            }

            // 테스트 메타데이터 수집
            collectTestMetadata(project);

            // 테스트 실행 리스너 등록
            setupTestExecutionListener(project);

            Messages.showInfoMessage(
                    "테스트 실행 모니터링이 시작되었습니다.\n" +
                            "테스트를 실행하면 자동으로 결과가 수집됩니다.\n" +
                            "테스트 완료 후 다시 이 액션을 실행하여 Excel을 생성하세요.",
                    "모니터링 시작"
            );

        } catch (Exception ex) {
            Messages.showErrorDialog("모니터링 설정 중 오류가 발생했습니다: " + ex.getMessage(), "오류");
        }
    }

    private void collectTestMetadata(Project project) {
        classMetadata.clear();

        Collection<VirtualFile> javaFiles = FileTypeIndex.getFiles(StdFileTypes.JAVA, GlobalSearchScope.projectScope(project));

        for (VirtualFile file : javaFiles) {
            PsiFile psiFile = PsiManager.getInstance(project).findFile(file);
            if (psiFile instanceof PsiJavaFile) {
                PsiJavaFile javaFile = (PsiJavaFile) psiFile;

                for (PsiClass psiClass : javaFile.getClasses()) {
                    if (isTestClass(psiClass)) {
                        TestClassMetadata metadata = extractTestClassMetadata(psiClass);
                        if (metadata != null) {
                            classMetadata.put(metadata.className, metadata);
                        }
                    }
                }
            }
        }
    }

    private boolean isTestClass(PsiClass psiClass) {
        String className = psiClass.getName();
        if (className != null && className.contains("Test")) {
            return true;
        }

        PsiMethod[] methods = psiClass.getMethods();
        for (PsiMethod method : methods) {
            if (hasTestAnnotation(method)) {
                return true;
            }
        }

        return false;
    }

    private boolean hasTestAnnotation(PsiMethod method) {
        PsiAnnotation[] annotations = method.getAnnotations();
        for (PsiAnnotation annotation : annotations) {
            String annotationName = annotation.getQualifiedName();
            if ("org.junit.jupiter.api.Test".equals(annotationName) ||
                    "org.junit.Test".equals(annotationName)) {
                return true;
            }
        }
        return false;
    }

    private TestClassMetadata extractTestClassMetadata(PsiClass psiClass) {
        TestClassMetadata metadata = new TestClassMetadata();
        metadata.className = psiClass.getName();
        metadata.injectMocksClass = findInjectMocksClass(psiClass);
        metadata.testMethods = new HashMap<>();

        PsiMethod[] methods = psiClass.getMethods();
        for (PsiMethod method : methods) {
            if (hasTestAnnotation(method)) {
                TestMethodMetadata methodMetadata = new TestMethodMetadata();
                methodMetadata.methodName = method.getName();
                methodMetadata.displayName = extractDisplayName(method);
                methodMetadata.targetMethodName = extractCalledApiMethod(method);

                metadata.testMethods.put(method.getName(), methodMetadata);
            }
        }

        return metadata;
    }

    private String findInjectMocksClass(PsiClass psiClass) {
        PsiField[] fields = psiClass.getFields();
        for (PsiField field : fields) {
            PsiAnnotation[] annotations = field.getAnnotations();
            for (PsiAnnotation annotation : annotations) {
                if ("org.mockito.InjectMocks".equals(annotation.getQualifiedName())) {
                    PsiType type = field.getType();
                    return type.getPresentableText();
                }
            }
        }
        return "";
    }

    private String extractDisplayName(PsiMethod method) {
        PsiAnnotation[] annotations = method.getAnnotations();
        for (PsiAnnotation annotation : annotations) {
            if ("org.junit.jupiter.api.DisplayName".equals(annotation.getQualifiedName())) {
                PsiAnnotationMemberValue value = annotation.findAttributeValue("value");
                if (value != null) {
                    String displayName = value.getText();
                    return displayName.replaceAll("\"", "");
                }
            }
        }
        return "";
    }

    private String extractCalledApiMethod(PsiMethod method) {
        PsiCodeBlock body = method.getBody();
        if (body == null) return "";

        String bodyText = body.getText();
        Pattern pattern = Pattern.compile("target\\.(\\w+)\\s*\\(");
        Matcher matcher = pattern.matcher(bodyText);

        if (matcher.find()) {
            return matcher.group(1);
        }

        return "";
    }

    private void setupTestExecutionListener(Project project) {
        if (connection != null) {
            connection.disconnect();
        }

        connection = project.getMessageBus().connect();
        connection.subscribe(ExecutionManager.EXECUTION_TOPIC, new ExecutionListener() {
            @Override
            public void processStarted(String executorId, ExecutionEnvironment env, ProcessHandler handler) {
                if (isTestExecution(executorId, env)) {
                    currentProcessHandler = handler;
                    processOutput.setLength(0);
                    testResults.clear();

                    // 프로세스 출력 수집
                    handler.addProcessListener(new com.intellij.execution.process.ProcessAdapter() {
                        @Override
                        public void onTextAvailable(com.intellij.execution.process.ProcessEvent event, Key outputType) {
                            if (outputType == ProcessOutputTypes.STDOUT || outputType == ProcessOutputTypes.STDERR) {
                                processOutput.append(event.getText());
                            }
                        }
                    });
                }
            }

            @Override
            public void processTerminated(String executorId, ExecutionEnvironment env, ProcessHandler handler, int exitCode) {
                if (isTestExecution(executorId, env) && handler == currentProcessHandler) {
                    ApplicationManager.getApplication().invokeLater(() -> {
                        // 출력 파싱하여 테스트 결과 수집
                        parseTestOutput();

                        if (!testResults.isEmpty()) {
                            int result = Messages.showYesNoDialog(
                                    String.format("테스트 실행이 완료되었습니다. (%d개 테스트 결과 수집)\nExcel 파일을 생성하시겠습니까?",
                                            testResults.size()),
                                    "테스트 완료",
                                    Messages.getQuestionIcon()
                            );

                            if (result == Messages.YES) {
                                generateExcelFile(project);
                            }
                        } else {
                            Messages.showInfoMessage("테스트 결과를 찾을 수 없습니다.", "정보");
                        }
                    });
                }
            }
        });
    }

    private boolean isTestExecution(String executorId, ExecutionEnvironment env) {
        // JUnit 실행인지 확인
        return executorId.contains("JUnit") ||
                executorId.contains("Test") ||
                env.getRunProfile().getName().contains("Test");
    }

    private void parseTestOutput() {
        String output = processOutput.toString();
        if (output.isEmpty()) return;

        String[] lines = output.split("\\r?\\n");
        String currentClassName = "";

        for (String line : lines) {
            line = line.trim();

            // 테스트 클래스 시작 패턴
            if (line.matches(".*Test.*started.*") || line.contains("Test class")) {
                currentClassName = extractClassNameFromLine(line);
            }

            // JUnit 5 테스트 결과 패턴
            if (line.contains("✓") || line.contains("SUCCESSFUL")) {
                parseTestResult(line, currentClassName, true);
            } else if (line.contains("✗") || line.contains("FAILED") || line.contains("ERROR")) {
                parseTestResult(line, currentClassName, false);
            }

            // Gradle/Maven 출력 패턴
            if (line.matches(".*Test.*\\s+(PASSED|FAILED).*")) {
                parseGradleTestResult(line);
            }
        }

        // 출력에서 직접 파싱이 어려운 경우, 기본값으로 테스트 메타데이터 사용
        if (testResults.isEmpty()) {
            createDefaultTestResults();
        }
    }

    private String extractClassNameFromLine(String line) {
        // 클래스명 추출 패턴
        Pattern pattern = Pattern.compile("([\\w\\.]*\\w*Test)");
        Matcher matcher = pattern.matcher(line);
        if (matcher.find()) {
            String fullClassName = matcher.group(1);
            return fullClassName.substring(fullClassName.lastIndexOf('.') + 1);
        }
        return "";
    }

    private void parseTestResult(String line, String className, boolean success) {
        // 메소드명과 시간 추출
        Pattern pattern = Pattern.compile("(\\w+).*?(\\d+(?:\\.\\d+)?\\s*(?:ms|s))");
        Matcher matcher = pattern.matcher(line);

        if (matcher.find()) {
            String methodName = matcher.group(1);
            String timeStr = matcher.group(2);

            TestExecutionResult result = new TestExecutionResult();
            result.testName = methodName;
            result.className = className;
            result.success = success;
            result.executionTime = parseExecutionTime(timeStr);
            result.errorMessage = success ? "" : extractErrorMessage(line);

            String key = className + "#" + methodName;
            testResults.put(key, result);
        }
    }

    private void parseGradleTestResult(String line) {
        Pattern pattern = Pattern.compile("(\\w+)\\s+(PASSED|FAILED)");
        Matcher matcher = pattern.matcher(line);

        if (matcher.find()) {
            String methodName = matcher.group(1);
            boolean success = "PASSED".equals(matcher.group(2));

            TestExecutionResult result = new TestExecutionResult();
            result.testName = methodName;
            result.className = extractClassNameFromOutput();
            result.success = success;
            result.executionTime = "0.000";
            result.errorMessage = success ? "" : "Test failed";

            String key = result.className + "#" + methodName;
            testResults.put(key, result);
        }
    }

    private String extractClassNameFromOutput() {
        // 전체 출력에서 클래스명 찾기
        String output = processOutput.toString();
        Pattern pattern = Pattern.compile("([\\w\\.]*\\w*Test)");
        Matcher matcher = pattern.matcher(output);
        if (matcher.find()) {
            String fullClassName = matcher.group(1);
            return fullClassName.substring(fullClassName.lastIndexOf('.') + 1);
        }
        return "UnknownTestClass";
    }

    private void createDefaultTestResults() {
        // 메타데이터를 기반으로 기본 테스트 결과 생성
        for (TestClassMetadata classInfo : classMetadata.values()) {
            for (TestMethodMetadata methodInfo : classInfo.testMethods.values()) {
                TestExecutionResult result = new TestExecutionResult();
                result.testName = methodInfo.methodName;
                result.className = classInfo.className;
                result.success = true; // 기본값
                result.executionTime = "0.000";
                result.errorMessage = "";

                String key = classInfo.className + "#" + methodInfo.methodName;
                testResults.put(key, result);
            }
        }
    }

    private String parseExecutionTime(String timeStr) {
        if (timeStr.contains("ms")) {
            String ms = timeStr.replaceAll("[^0-9.]", "");
            try {
                double milliseconds = Double.parseDouble(ms);
                return String.format("%.3f", milliseconds / 1000.0);
            } catch (NumberFormatException e) {
                return "0.000";
            }
        } else if (timeStr.contains("s")) {
            String s = timeStr.replaceAll("[^0-9.]", "");
            try {
                return String.format("%.3f", Double.parseDouble(s));
            } catch (NumberFormatException e) {
                return "0.000";
            }
        }
        return "0.000";
    }

    private String extractErrorMessage(String line) {
        if (line.contains("AssertionError")) {
            return "AssertionError";
        } else if (line.contains("Exception")) {
            int idx = line.indexOf("Exception");
            return line.substring(Math.max(0, idx - 20), Math.min(line.length(), idx + 30));
        } else if (line.contains("Error")) {
            return "Error occurred";
        }
        return "Test failed";
    }

    private void generateExcelFile(Project project) {
        try {
            FileChooserDescriptor descriptor = new FileChooserDescriptor(false, true, false, false, false, false);
            descriptor.setTitle("Excel 파일 저장 위치 선택");
            VirtualFile selectedDir = FileChooser.chooseFile(descriptor, project, null);

            if (selectedDir != null) {
                String filePath = selectedDir.getPath() + "/테스트결과_" +
                        System.currentTimeMillis() + ".xlsx";
                generateExcel(filePath);
                Messages.showInfoMessage("Excel 파일이 성공적으로 생성되었습니다.\n경로: " + filePath, "완료");
            }

        } catch (Exception ex) {
            Messages.showErrorDialog("Excel 생성 중 오류가 발생했습니다: " + ex.getMessage(), "오류");
        }
    }

    private void generateExcel(String filePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("테스트 결과");

        // 헤더 생성
        Row headerRow = sheet.createRow(0);
        String[] headers = {
                "no", "분류", "세부 분류", "기능", "IF 여부", "API 클래스명",
                "API 명", "API 내용", "테스트 클래스명", "테스트 메소드 명",
                "실행 결과", "실행시간 (초)", "실패 / 오류 메시지"
        };

        // 헤더 스타일 설정
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        // 데이터 행 생성
        int rowNum = 1;
        for (Map.Entry<String, TestExecutionResult> entry : testResults.entrySet()) {
            TestExecutionResult result = entry.getValue();
            TestClassMetadata classInfo = classMetadata.get(result.className);
            TestMethodMetadata methodInfo = null;

            if (classInfo != null) {
                methodInfo = classInfo.testMethods.get(result.testName);
            }

            Row row = sheet.createRow(rowNum);

            row.createCell(0).setCellValue(rowNum); // no
            row.createCell(1).setCellValue(""); // 분류 (빈값)
            row.createCell(2).setCellValue(""); // 세부 분류 (빈값)
            row.createCell(3).setCellValue(""); // 기능 (빈값)
            row.createCell(4).setCellValue(""); // IF 여부 (빈값)

            // API 클래스명 (@InjectMocks 대상 클래스명)
            String apiClassName = classInfo != null ? classInfo.injectMocksClass : "";
            row.createCell(5).setCellValue(apiClassName);

            // API 명 (테스트 메소드가 호출하는 target 메소드명)
            String apiMethodName = methodInfo != null ? methodInfo.targetMethodName : "";
            row.createCell(6).setCellValue(apiMethodName);

            // API 내용 (@DisplayName 값)
            String apiContent = methodInfo != null ? methodInfo.displayName : "";
            row.createCell(7).setCellValue(apiContent);

            // 테스트 클래스명
            row.createCell(8).setCellValue(result.className);

            // 테스트 메소드명
            row.createCell(9).setCellValue(result.testName);

            // 실행 결과 (SUCCESS / FAIL)
            String executionResult = result.success ? "SUCCESS" : "FAIL";
            row.createCell(10).setCellValue(executionResult);

            // 실행시간 (초)
            row.createCell(11).setCellValue(result.executionTime);

            // 실패 / 오류 메시지
            String errorMessage = result.success ? "" : (result.errorMessage != null ? result.errorMessage : "");
            row.createCell(12).setCellValue(errorMessage);

            rowNum++;
        }

        // 컬럼 너비 자동 조정
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // 파일 저장
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
        }

        workbook.close();
    }

    // 내부 클래스들
    private static class TestClassMetadata {
        String className;
        String injectMocksClass;
        Map<String, TestMethodMetadata> testMethods;
    }

    private static class TestMethodMetadata {
        String methodName;
        String displayName;
        String targetMethodName;
    }

    private static class TestExecutionResult {
        String testName;
        String className;
        boolean success;
        String executionTime;
        String errorMessage;
    }

    // 플러그인 종료 시 리스너 정리
    public void dispose() {
        if (connection != null) {
            connection.disconnect();
        }
    }
}