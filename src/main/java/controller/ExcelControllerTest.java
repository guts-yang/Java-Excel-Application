package controller;

import model.ExcelDataModel;
import model.ExcelDataAccess;
import view.ExcelView;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel控制器测试类
 * 用于测试MVC架构的交互功能
 */
public class ExcelControllerTest {
    
    public static void main(String[] args) {
        System.out.println("开始测试MVC架构交互...");
        
        // 测试1：模型层功能测试
        testModelLayer();
        
        // 测试2：数据访问层测试
        testDataAccessLayer();
        
        // 测试3：MVC集成测试
        // 注意：完整的集成测试在GUI运行时进行
        System.out.println("\nMVC架构测试完成，请运行MainApp进行完整功能测试");
    }
    
    /**
     * 测试Model层功能
     */
    private static void testModelLayer() {
        System.out.println("\n测试1：模型层功能测试");
        
        // 创建模型实例
        ExcelDataModel model = new ExcelDataModel();
        
        // 添加测试数据
        List<String> headers = new ArrayList<>();
        headers.add("姓名");
        headers.add("年龄");
        headers.add("邮箱");
        
        List<List<String>> data = new ArrayList<>();
        List<String> row1 = new ArrayList<>();
        row1.add("张三");
        row1.add("25");
        row1.add("zhangsan@example.com");
        data.add(row1);
        
        List<String> row2 = new ArrayList<>();
        row2.add("李四");
        row2.add("30");
        row2.add("lisi@example.com");
        data.add(row2);
        
        // 设置数据
        model.setData(data, headers);
        model.setFileName("测试文件.xlsx");
        
        // 验证数据设置是否成功
        System.out.println("文件名: " + model.getFileName());
        System.out.println("表头: " + model.getHeaders());
        System.out.println("数据行数: " + model.getRowCount());
        System.out.println("列数: " + model.getColumnCount());
        
        // 测试添加行
        List<String> newRow = new ArrayList<>();
        newRow.add("王五");
        newRow.add("28");
        newRow.add("wangwu@example.com");
        model.addRow(newRow);
        System.out.println("添加行后数据行数: " + model.getRowCount());
        
        // 测试更新行
        List<String> updatedRow = new ArrayList<>();
        updatedRow.add("王五更新");
        updatedRow.add("29");
        updatedRow.add("wangwu_update@example.com");
        model.updateRow(2, updatedRow);
        System.out.println("更新后第三行数据: " + model.getData().get(2));
        
        // 测试删除行
        model.deleteRow(1);
        System.out.println("删除行后数据行数: " + model.getRowCount());
        
        System.out.println("模型层功能测试通过！");
    }
    
    /**
     * 测试数据访问层功能
     */
    private static void testDataAccessLayer() {
        System.out.println("\n测试2：数据访问层测试");
        
        try {
            ExcelDataAccess dataAccess = new ExcelDataAccess();
            
            // 准备测试数据
            List<String> headers = new ArrayList<>();
            headers.add("测试列1");
            headers.add("测试列2");
            
            List<List<String>> testData = new ArrayList<>();
            List<String> row1 = new ArrayList<>();
            row1.add("测试数据1");
            row1.add("测试数据2");
            testData.add(row1);
            
            // 测试写入功能
            String testFilePath = "D:\\coding\\java\\class_1\\m11d24\\project1\\test_data.xlsx";
            dataAccess.writeExcel(testFilePath, headers, testData);
            System.out.println("Excel写入测试成功，文件保存至: " + testFilePath);
            
            // 测试读取功能
            Object[] result = dataAccess.readExcel(testFilePath);
            List<String> readHeaders = (List<String>) result[0];
            List<List<String>> readData = (List<List<String>>) result[1];
            
            System.out.println("读取的表头: " + readHeaders);
            System.out.println("读取的数据: " + readData);
            
            System.out.println("数据访问层测试通过！");
        } catch (IOException e) {
            System.out.println("数据访问层测试失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}