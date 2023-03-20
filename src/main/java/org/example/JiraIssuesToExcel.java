package org.example;

import com.atlassian.jira.rest.client.api.JiraRestClient;
import com.atlassian.jira.rest.client.api.domain.Issue;
import com.atlassian.jira.rest.client.internal.async.AsynchronousJiraRestClientFactory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.net.URI;

public class JiraIssuesToExcel {

    public static void main(String[] args) throws Exception {

        // Jira API 접근 정보 설정
        String jiraUrl = "https://your-jira-instance.com";
        String jiraUsername = "your-jira-username";
        String jiraPassword = "your-jira-password";

        // Excel 파일 생성
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Issues");

        // 이슈 검색 쿼리 설정
        String jqlQuery = "assignee in (a, b, c)";

        // Jira API 호출
        AsynchronousJiraRestClientFactory factory = new AsynchronousJiraRestClientFactory();
        URI jiraServerUri = new URI(jiraUrl);
        JiraRestClient restClient = factory.createWithBasicHttpAuthentication(jiraServerUri, jiraUsername, jiraPassword);
        Iterable<Issue> issues = restClient.getSearchClient().searchJql(jqlQuery).claim().getIssues();

        // 이슈 정보 엑셀에 추가
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Key");
        headerRow.createCell(1).setCellValue("Summary");
        headerRow.createCell(2).setCellValue("Start Date");
        headerRow.createCell(3).setCellValue("Due Date");
        headerRow.createCell(4).setCellValue("Reporter");

        int rowIdx = 1;
        for (Issue issue : issues) {
            Row dataRow = sheet.createRow(rowIdx++);
            dataRow.createCell(0).setCellValue(issue.getKey());
            dataRow.createCell(1).setCellValue(issue.getSummary());
            //dataRow.createCell(2).setCellValue(issue.getStartDate() != null ? issue.getStartDate().toString() : "");
            dataRow.createCell(3).setCellValue(issue.getDueDate() != null ? issue.getDueDate().toString() : "");
            dataRow.createCell(4).setCellValue(issue.getReporter().getDisplayName());
        }

        // Excel 파일 저장
        FileOutputStream outputStream = new FileOutputStream("issues.xlsx");
        workbook.write(outputStream);
        workbook.close();
    }
}
