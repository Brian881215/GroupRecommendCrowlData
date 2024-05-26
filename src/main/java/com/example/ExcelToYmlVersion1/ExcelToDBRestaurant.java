package com.example.ExcelToYmlVersion1;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;

public class ExcelToDBRestaurant {
    public static void main(String[] args){
        System.out.println("hello");
        String jdbcURL = "jdbc:mysql://localhost:3307/GroupRecommend?useSSL=false&allowPublicKeyRetrieval=true";
        String username = "Brian881215";
        String password = "Brian881215";

        String excelFilePath = "/Users/huangchengen/eclipse-workspace_ee/GroupRecommend/restaurants_data.xlsx";

        try (Connection connection = DriverManager.getConnection(jdbcURL, username, password)) {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            Sheet firstSheet = workbook.getSheetAt(0);
            Row row;

            String sql = "INSERT INTO recommend_restaurant (name, averageCost, district, address, openHours, rating, reviewCount, link, tags,classification) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
            PreparedStatement statement = connection.prepareStatement(sql);

            int rowCount = firstSheet.getLastRowNum();

            for (int i = 1; i <= rowCount; i++) {
                row = firstSheet.getRow(i);

                String name = row.getCell(0).getStringCellValue();
                double averageCost = row.getCell(1).getNumericCellValue();
                String district = row.getCell(2).getStringCellValue();
                String address = row.getCell(3).getStringCellValue();
                String openHours = row.getCell(4).getStringCellValue();
                double rating = row.getCell(5).getNumericCellValue();
                int reviewCount = (int) row.getCell(6).getNumericCellValue();
                String link = row.getCell(7).getStringCellValue();
                String tags = row.getCell(8).getStringCellValue();
                String classification = row.getCell(9).getStringCellValue();
                
                statement.setString(1, name);
                statement.setDouble(2, averageCost);
                statement.setString(3, district);
                statement.setString(4, address);
                statement.setString(5, openHours);
                statement.setDouble(6, rating);
                statement.setInt(7, reviewCount);
                statement.setString(8, link);
                statement.setString(9, tags);
                statement.setString(10, classification);
                statement.addBatch();
            }

            statement.executeBatch(); // Execute all the insert statements at once
            workbook.close();
            inputStream.close();

            System.out.println("Data is inserted successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
