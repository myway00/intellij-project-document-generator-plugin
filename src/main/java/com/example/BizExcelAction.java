package com.example;

import com.intellij.openapi.actionSystem.AnAction;
import com.intellij.openapi.actionSystem.AnActionEvent;
import com.intellij.openapi.actionSystem.CommonDataKeys;
import com.intellij.openapi.project.Project;
import com.intellij.openapi.ui.Messages;
import com.intellij.notification.Notification;
import com.intellij.notification.NotificationType;
import com.intellij.notification.Notifications;
import com.intellij.openapi.vfs.VirtualFile;
import com.intellij.psi.PsiClass;
import com.intellij.psi.PsiFile;
import com.intellij.psi.PsiJavaFile;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class BizExcelAction extends AnAction {
    @Override
    public void actionPerformed(AnActionEvent e) {
        Project project = e.getProject();
        if (project == null) {
            Messages.showErrorDialog("프로젝트를 찾을 수 없습니다.", "오류");
            return;
        }

        try {
            VirtualFile[] files = e.getData(CommonDataKeys.VIRTUAL_FILE_ARRAY);
            if (files == null || files.length == 0) {
                Messages.showErrorDialog("선택된 파일이 없습니다.", "오류");
                return;
            }

            List<PsiClass> controllers = new ArrayList<>();

            for (VirtualFile file : files) {
                if (file.getName().endsWith(".java")) {
                    PsiFile psiFile = com.intellij.psi.PsiManager.getInstance(project).findFile(file);
                    if (psiFile instanceof PsiJavaFile) {
                        PsiJavaFile javaFile = (PsiJavaFile) psiFile;
                        PsiClass[] psiClasses = javaFile.getClasses();

                        for (PsiClass psiClass : psiClasses) {
                            String qualifiedName = psiClass.getQualifiedName();
                            if (qualifiedName != null) {
                                // controllers 리스트에 PsiClass 또는 클래스 이름 저장
                                controllers.add(psiClass);
                            }
                        }
                    }
                }
            }

            if (controllers.isEmpty()) {
                Messages.showErrorDialog("Java 클래스를 찾을 수 없습니다.", "오류");
                return;
            }

            showNotification("API 가이드 엑셀 생성을 시작합니다...", NotificationType.INFORMATION);

            BizExcelExporter exporter = new BizExcelExporter();

            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
            String fileName = "API_가이드_" + timestamp + ".xlsx";
            File output = new File(System.getProperty("user.home"), fileName);

            exporter.exportControllerExcel(output, controllers);

            int choice = Messages.showYesNoDialog(
                    "API 가이드 엑셀 파일이 성공적으로 생성되었습니다!\n\n파일 경로: " + output.getAbsolutePath() + "\n\n파일을 열어보시겠습니까?",
                    "API 가이드 생성 완료",
                    "파일 열기", "확인",
                    Messages.getQuestionIcon()
            );

            if (choice == Messages.YES && java.awt.Desktop.isDesktopSupported()) {
                java.awt.Desktop.getDesktop().open(output);
            }

            showNotification("API 가이드 엑셀 파일 생성이 완료되었습니다: " + fileName, NotificationType.INFORMATION);

        } catch (Exception ex) {
            ex.printStackTrace();
            Messages.showErrorDialog("API 가이드 엑셀 생성 중 오류가 발생했습니다:\n" + ex.getMessage(), "오류");
            showNotification("API 가이드 생성 실패: " + ex.getMessage(), NotificationType.ERROR);
        }
    }

    private void showNotification(String content, NotificationType type) {
        Notification notification = new Notification(
            "ControllerGuide",
            "Controller Guide Plugin",
            content,
            type
        );
        Notifications.Bus.notify(notification);
    }
}
