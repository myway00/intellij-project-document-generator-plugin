package com.example;

import com.intellij.psi.*;

import java.util.ArrayList;
import java.util.List;

public class CommonAction {
    public static List<String> generateLogicDescriptionsBiz(PsiMethod method) {
        List<String> descriptions = new ArrayList<>();
        String methodName = method.getName().toLowerCase(); // 대소문자 구분 없이 비교

        if (methodName.contains("create")) {
            descriptions.add("1. 컨트롤러 진입");
            descriptions.add("2. create 처리 수행");
        } else if (methodName.contains("update")) {
            descriptions.add("1. 컨트롤러 진입");
            descriptions.add("2. update 처리 수행");
        } else if (methodName.contains("realdelete")) {
            descriptions.add("1. 컨트롤러 진입");
            descriptions.add("2. realdelete 처리 수행");
        } else if (methodName.contains("undelete")) {
            descriptions.add("1. 컨트롤러 진입");
            descriptions.add("2. undelete 처리 수행");
        } else if (methodName.contains("delete")) {
            descriptions.add("1. 컨트롤러 진입");
            descriptions.add("2. delete 처리 수행");
        } else {
            descriptions.add("1. 조회 컨트롤러 진입");
            descriptions.add("2. 조회 처리 수행");
        }

        return descriptions;
    }


        /**
         * CORE 모듈 메소드의 로직을 한 줄씩 분석하여 설명 생성
         * @param method 분석할 메소드
         * @param isController 컨트롤러 메소드인지 여부 (true면 서비스 메소드도 분석)
         */
    public static List<String> generateLogicDescriptions(PsiMethod method, boolean isController) {
        List<String> descriptions = new ArrayList<>();

        if (method.getBody() == null) {
            descriptions.add("메소드 구현부가 없습니다.");
            return descriptions;
        }

        PsiCodeBlock body = method.getBody();
        PsiStatement[] statements = body.getStatements();

        int stepNumber = 1;

        for (PsiStatement statement : statements) {
            String description = analyzeStatement(statement, stepNumber, isController);
            if (!description.isEmpty()) {
                descriptions.add(description);
                stepNumber++;
            }
        }

        if (descriptions.isEmpty()) {
            descriptions.add("메소드 로직이 비어있습니다.");
        }

        return descriptions;
    }
    /**
     * 개별 Statement를 분석하여 로직 설명 생성
     */
    private static String analyzeStatement(PsiStatement statement, int stepNumber, boolean isController) {
        if (statement == null) return "";

        String statementText = statement.getText().trim();

        // 주석 제거
        statementText = removeComments(statementText);
        if (statementText.isEmpty()) return "";

        // Statement 타입별 분석
        if (statement instanceof PsiDeclarationStatement) {
            return analyzeDeclarationStatement((PsiDeclarationStatement) statement);
        } else if (statement instanceof PsiExpressionStatement) {
            return analyzeExpressionStatement((PsiExpressionStatement) statement, isController);
        } else if (statement instanceof PsiIfStatement) {
            return analyzeIfStatement((PsiIfStatement) statement);
        } else if (statement instanceof PsiForStatement) {
            return analyzeForStatement((PsiForStatement) statement);
        } else if (statement instanceof PsiWhileStatement) {
            return analyzeWhileStatement((PsiWhileStatement) statement);
        } else if (statement instanceof PsiReturnStatement) {
            return analyzeReturnStatement((PsiReturnStatement) statement);
        } else if (statement instanceof PsiTryStatement) {
            return analyzeTryStatement((PsiTryStatement) statement);
        } else if (statement instanceof PsiThrowStatement) {
            return analyzeThrowStatement((PsiThrowStatement) statement);
        } else {
            // 기타 Statement들
            return "기타 로직 처리: " + truncateText(statementText, 50);
        }
    }

    /**
     * 변수 선언문 분석
     */
    private static String analyzeDeclarationStatement(PsiDeclarationStatement statement) {
        PsiElement[] elements = statement.getDeclaredElements();
        if (elements.length > 0 && elements[0] instanceof PsiVariable) {
            PsiVariable variable = (PsiVariable) elements[0];
            String varName = variable.getName();
            String varType = variable.getType().getPresentableText();

            if (variable.getInitializer() != null) {
                String initText = variable.getInitializer().getText();
                return String.format("%s 타입의 변수 '%s'를 선언하고 초기화", varType, varName);
            } else {
                return String.format("%s 타입의 변수 '%s'를 선언", varType, varName);
            }
        }
        return "변수 선언 처리";
    }

    /**
     * 표현식문 분석 (서비스 메소드 분석 기능 추가)
     */
    private static String analyzeExpressionStatement(PsiExpressionStatement statement, boolean isController) {
        PsiExpression expression = statement.getExpression();

        if (expression instanceof PsiMethodCallExpression) {
            PsiMethodCallExpression methodCall = (PsiMethodCallExpression) expression;
            String methodName = methodCall.getMethodExpression().getReferenceName();

            if (methodName != null) {
                StringBuilder result = new StringBuilder();
                
                // 기본 메소드 분석
                if (methodName.startsWith("set")) {
                    result.append("속성 값 설정: ").append(methodName);
                } else if (methodName.startsWith("get")) {
                    result.append("속성 값 조회: ").append(methodName);
                } else if (methodName.contains("save") || methodName.contains("insert") || methodName.contains("create")) {
                    result.append("데이터 저장/삽입 처리: ").append(methodName);
                } else if (methodName.contains("update") || methodName.contains("modify")) {
                    result.append("데이터 수정 처리: ").append(methodName);
                } else if (methodName.contains("delete") || methodName.contains("remove")) {
                    result.append("데이터 삭제 처리: ").append(methodName);
                } else if (methodName.contains("find") || methodName.contains("select") || methodName.contains("search")) {
                    result.append("데이터 조회 처리: ").append(methodName);
                } else if (methodName.contains("validate") || methodName.contains("check")) {
                    result.append("데이터 검증 처리: ").append(methodName);
                } else {
                    result.append("메소드 호출: ").append(methodName);
                }

                // 컨트롤러일 경우 서비스 메소드 분석 추가
                if (isController && isServiceMethodCall(methodCall)) {
                    String serviceAnalysis = analyzeServiceMethod(methodCall);
                    if (!serviceAnalysis.isEmpty()) {
                        result.append("\n    └─ 서비스 로직 분석:\n").append(serviceAnalysis);
                    }
                }

                return result.toString();
            }
        } else if (expression instanceof PsiAssignmentExpression) {
            PsiAssignmentExpression assignment = (PsiAssignmentExpression) expression;
            
            // 할당문에서 서비스 메소드 호출 확인
            if (isController && assignment.getRExpression() instanceof PsiMethodCallExpression) {
                PsiMethodCallExpression methodCall = (PsiMethodCallExpression) assignment.getRExpression();
                StringBuilder result = new StringBuilder("변수 할당 처리");
                
                if (isServiceMethodCall(methodCall)) {
                    String serviceAnalysis = analyzeServiceMethod(methodCall);
                    if (!serviceAnalysis.isEmpty()) {
                        result.append("\n    └─ 서비스 로직 분석:\n").append(serviceAnalysis);
                    }
                }
                return result.toString();
            }
            return "변수 할당 처리";
        }

        return "표현식 처리";
    }

    /**
     * 서비스 메소드 호출인지 확인
     */
    private static boolean isServiceMethodCall(PsiMethodCallExpression methodCall) {
        PsiExpression qualifierExpression = methodCall.getMethodExpression().getQualifierExpression();
        if (qualifierExpression instanceof PsiReferenceExpression) {
            PsiReferenceExpression ref = (PsiReferenceExpression) qualifierExpression;
            PsiElement resolved = ref.resolve();
            
            if (resolved instanceof PsiField) {
                PsiField field = (PsiField) resolved;
                String fieldTypeName = field.getType().getCanonicalText();
                
                // 서비스 클래스 패턴 확인 (Service로 끝나는 클래스명)
                return fieldTypeName.toLowerCase().contains("service") || 
                       field.hasAnnotation("org.springframework.beans.factory.annotation.Autowired");
            }
        }
        return false;
    }

    /**
     * 서비스 메소드 분석
     */
    private static String analyzeServiceMethod(PsiMethodCallExpression methodCall) {
        try {
            PsiMethod serviceMethod = (PsiMethod) methodCall.resolveMethod();
            if (serviceMethod != null && serviceMethod.getBody() != null) {
                List<String> serviceDescriptions = generateLogicDescriptions(serviceMethod, false);
                
                StringBuilder result = new StringBuilder();
                for (int i = 0; i < serviceDescriptions.size(); i++) {
                    result.append("      ").append(i + 1).append(". ").append(serviceDescriptions.get(i));
                    if (i < serviceDescriptions.size() - 1) {
                        result.append("\n");
                    }
                }
                return result.toString();
            }
        } catch (Exception e) {
            return "      서비스 메소드 분석 중 오류 발생";
        }
        return "";
    }

    /**
     * If문 분석
     */
    private static String analyzeIfStatement(PsiIfStatement statement) {
        PsiExpression condition = statement.getCondition();
        if (condition != null) {
            String conditionText = condition.getText();
            return "조건 분기 처리: " + truncateText(conditionText, 30);
        }
        return "조건 분기 처리";
    }

    /**
     * For문 분석
     */
    private static String analyzeForStatement(PsiForStatement statement) {
        return "반복 처리 (for 루프)";
    }

    /**
     * While문 분석
     */
    private static String analyzeWhileStatement(PsiWhileStatement statement) {
        return "반복 처리 (while 루프)";
    }

    /**
     * Return문 분석
     */
    private static String analyzeReturnStatement(PsiReturnStatement statement) {
        PsiExpression returnValue = statement.getReturnValue();
        if (returnValue != null) {
            String returnText = returnValue.getText();
            if (returnText.contains("new") && returnText.contains("ResponseEntity")) {
                return "ResponseEntity 객체를 생성하여 응답 반환";
            } else if (returnText.contains("ResponseEntity")) {
                return "ResponseEntity 응답 반환";
            } else {
                return "결과 값 반환: " + truncateText(returnText, 30);
            }
        }
        return "결과 반환";
    }

    /**
     * Try-Catch문 분석
     */
    private static String analyzeTryStatement(PsiTryStatement statement) {
        return "예외 처리 블록 시작";
    }

    /**
     * Throw문 분석
     */
    private static String analyzeThrowStatement(PsiThrowStatement statement) {
        return "예외 발생 처리";
    }

    /**
     * 주석 제거 유틸리티
     */
    private static String removeComments(String text) {
        // 한 줄 주석 제거
        text = text.replaceAll("//.*", "");
        // 여러 줄 주석 제거
        text = text.replaceAll("/\\*.*?\\*/", "");
        return text.trim();
    }

    /**
     * 텍스트 길이 제한 유틸리티
     */
    private static String truncateText(String text, int maxLength) {
        if (text == null) return "";
        if (text.length() <= maxLength) return text;
        return text.substring(0, maxLength) + "...";
    }

}
