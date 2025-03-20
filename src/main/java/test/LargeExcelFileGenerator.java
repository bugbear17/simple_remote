package test;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class LargeExcelFileGenerator {
    public static void main(String[] args) {
        // 파일 이름과 크기 설정
        String fileName = "large_excel_file.xlsx";
        int rows = 300000; // 행 개수
        int cols = 100;     // 열 개수

        // SXSSFWorkbook: 메모리 효율적인 스트리밍 방식 사용
        try (SXSSFWorkbook workbook = new SXSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(fileName)) {

            // 시트 생성
            Sheet sheet = workbook.createSheet("Large Data");

            // 열 제목 (헤더) 생성
            Row headerRow = sheet.createRow(0);
            for (int col = 0; col < cols; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue("Column " + (col + 1));
            }

            // 데이터 행 추가
            for (int row = 1; row <= rows; row++) {
                Row dataRow = sheet.createRow(row);
                for (int col = 0; col < cols; col++) {
                    Cell cell = dataRow.createCell(col);
                    cell.setCellValue("Row" + row + "-Col" + (col + 1));
                }

                // 진행 상황 출력
                if (row % 1000 == 0) {
                    System.out.println(row + " rows written...");
                }
            }

            // 파일 쓰기
            workbook.write(fos);
            System.out.println("Excel file '" + fileName + "' created successfully!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}