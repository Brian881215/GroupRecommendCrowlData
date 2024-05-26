package com.example.ExcelToYmlVersion1;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.yaml.snakeyaml.DumperOptions;
import org.yaml.snakeyaml.Yaml;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.*;
@SpringBootApplication
public class ExcelToYmlVersion1Application {

	public static void main(String[] args)throws IOException  {
		SpringApplication.run(ExcelToYmlVersion1Application.class, args);
		
	    // Load Excel file
        FileInputStream excelFile = new FileInputStream("/Users/huangchengen/eclipse-workspace_ee/ExcelToYmlVersion1/datahubSapTesting.xlsx");
        Workbook workbook = new XSSFWorkbook(excelFile);//用XSSFWorkbook創建一個工作簿對象
        Sheet datatypeSheet = workbook.getSheetAt(0);//工作簿中取得第一個工作表

        // Prepare the data structure for YAML
        Map<String, Object> yamlData = new LinkedHashMap<>();//用linkedHashMap來儲存要轉換為yaml的數據
        //自行設定yaml文件的基本結構
        yamlData.put("version", 1);
        yamlData.put("source", "DataHub");
        yamlData.put("owners", Collections.singletonMap("users", Collections.singletonList("mjames")));
        yamlData.put("url", "https://github.com/datahub-project/datahub/");
        List<Map<String, Object>> nodes = new ArrayList<>();

        // Iterate over rows
        for (Row currentRow : datatypeSheet) {
        	//跳過第一行標題行
        	if (currentRow.getRowNum() == 0) {
                continue;
            }

            // Read cells (assuming they are in order: GlossaryTerm, GlossaryNode, Description, Owners)
            String glossaryTerm = getCellValue(currentRow.getCell(0));
            String glossaryNode = getCellValue(currentRow.getCell(1));
            String description = getCellValue(currentRow.getCell(2));
            String owners = getCellValue(currentRow.getCell(3));

            // Process the GlossaryNode to create nested structure
            String[] nodeLevels = glossaryNode.split("/");
            Map<String, Object> currentNode = findOrCreateNode(nodes, nodeLevels, 0);

            // Add the term to the last node
            if (!currentNode.containsKey("terms")) {
                currentNode.put("terms", new ArrayList<>());
            }
            List<Map<String, Object>> terms = (List<Map<String, Object>>) currentNode.get("terms");
            Map<String, Object> term = new LinkedHashMap<>();
            term.put("name", glossaryTerm);
            term.put("description", description);
            term.put("owners", Collections.singletonMap("users", Arrays.asList(owners.split(","))));
            terms.add(term);
        }

        yamlData.put("nodes", nodes);

        // Write YAML file
        DumperOptions options = new DumperOptions();
        options.setDefaultFlowStyle(DumperOptions.FlowStyle.BLOCK);
        Yaml yaml = new Yaml(options);
        FileWriter writer = new FileWriter("glossaryOutput.yaml");
        yaml.dump(yamlData, writer);

        // Close resources
        workbook.close();
        writer.close();
		System.out.println("helloworld");
	}
	private static String getCellValue(Cell cell) {//如果單元格為空，則返回空字符串
	    return cell == null ? "" : cell.toString();
	}
	//遞迴查找或創建節點
	private static Map<String, Object> findOrCreateNode(List<Map<String, Object>> nodes, String[] levels, int index) {
        if (index >= levels.length) {//條件符合時表示已經遍歷過底層了
            return null;
        }

        String level = levels[index];
        for (Map<String, Object> node : nodes) {
            if (node.get("name").equals(level)) {
                // Found the node, go deeper
                if (index < levels.length - 1) {
                    if (!node.containsKey("nodes")) {
                        node.put("nodes", new ArrayList<>());
                    }
                    return findOrCreateNode((List<Map<String, Object>>) node.get("nodes"), levels, index + 1);
                }
                return node;
            }
        }

        // Node not found, create it
        Map<String, Object> newNode = new LinkedHashMap<>();
        newNode.put("name", level);
        newNode.put("description", level + " related terms.");
        nodes.add(newNode);

        // Go deeper if not at the last level
        if (index < levels.length - 1) {
            newNode.put("nodes", new ArrayList<>());
            return findOrCreateNode((List<Map<String, Object>>) newNode.get("nodes"), levels, index + 1);
        }

        return newNode;
    }

}
