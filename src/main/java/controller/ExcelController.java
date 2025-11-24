package controller;

import model.ExcelDataAccess;
import model.ExcelDataModel;
import view.ExcelView;

import javax.swing.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Observable;
import java.util.Observer;

/**
 * Excel控制器类 - MVC架构中的Controller层
 * 负责处理用户请求，协调Model和View之间的交互
 */
public class ExcelController implements Observer {
    private ExcelDataModel model;
    private ExcelView view;
    private ExcelDataAccess dataAccess;

    /**
     * 构造函数，初始化Controller并建立Model和View的连接
     */
    public ExcelController(ExcelDataModel model, ExcelView view) {
        this.model = model;
        this.view = view;
        this.dataAccess = new ExcelDataAccess();

        // 将Controller注册为Model的观察者，以便接收数据变化通知
        this.model.addObserver(this);

        // 将Controller传递给View，以便View可以调用Controller的方法
        this.view.setController(this);

        // 初始化时更新View
        updateView();
    }

    /**
     * 处理Excel文件导入
     */
    public void importExcel() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("选择Excel文件");
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
                "Excel文件 (*.xlsx)", "xlsx"));

        int userSelection = fileChooser.showOpenDialog(view);
        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File fileToOpen = fileChooser.getSelectedFile();
            try {
                // 从文件读取数据
                Object[] result = dataAccess.readExcel(fileToOpen.getAbsolutePath());
                List<String> headers = (List<String>) result[0];
                List<List<String>> data = (List<List<String>>) result[1];

                // 更新Model
                model.setFileName(fileToOpen.getName());
                model.setData(data, headers);

                // 显示成功消息
                JOptionPane.showMessageDialog(view, "Excel文件导入成功！", "成功", JOptionPane.INFORMATION_MESSAGE);
            } catch (IOException e) {
                // 显示错误消息
                JOptionPane.showMessageDialog(view, "导入失败：" + e.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                e.printStackTrace();
            }
        }
    }

    /**
     * 处理Excel文件导出
     */
    public void exportExcel() {
        if (model.getHeaders().isEmpty()) {
            JOptionPane.showMessageDialog(view, "没有数据可导出", "警告", JOptionPane.WARNING_MESSAGE);
            return;
        }

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("保存Excel文件");
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
                "Excel文件 (*.xlsx)", "xlsx"));

        // 设置默认文件名
        fileChooser.setSelectedFile(new File(model.getFileName()));

        int userSelection = fileChooser.showSaveDialog(view);
        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File fileToSave = fileChooser.getSelectedFile();
            
            // 确保文件扩展名为.xlsx
            String filePath = fileToSave.getAbsolutePath();
            if (!filePath.toLowerCase().endsWith(".xlsx")) {
                filePath += ".xlsx";
                fileToSave = new File(filePath);
            }

            try {
                // 写入文件
                dataAccess.writeExcel(filePath, model.getHeaders(), model.getData());

                // 更新Model中的文件名
                model.setFileName(fileToSave.getName());

                // 显示成功消息
                JOptionPane.showMessageDialog(view, "Excel文件导出成功！", "成功", JOptionPane.INFORMATION_MESSAGE);
            } catch (IOException e) {
                // 显示错误消息
                JOptionPane.showMessageDialog(view, "导出失败：" + e.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                e.printStackTrace();
            }
        }
    }

    /**
     * 添加新行数据
     */
    public void addRow(List<String> rowData) {
        model.addRow(rowData);
    }

    /**
     * 更新指定行数据
     */
    public void updateRow(int rowIndex, List<String> newData) {
        model.updateRow(rowIndex, newData);
    }

    /**
     * 删除指定行数据
     */
    public void deleteRow(int rowIndex) {
        model.deleteRow(rowIndex);
    }

    /**
     * 清空所有数据
     */
    public void clearAllData() {
        int confirm = JOptionPane.showConfirmDialog(view, "确定要清空所有数据吗？", "确认", JOptionPane.YES_NO_OPTION);
        if (confirm == JOptionPane.YES_OPTION) {
            model.clearData();
            model.setFileName("未命名.xlsx");
        }
    }

    /**
     * 创建新的空白表格
     */
    public void createNewTable() {
        // 默认创建一个包含姓名、年龄、邮箱三列的空表格
        List<String> defaultHeaders = new ArrayList<>();
        defaultHeaders.add("姓名");
        defaultHeaders.add("年龄");
        defaultHeaders.add("邮箱");
        
        List<List<String>> emptyData = new ArrayList<>();
        
        model.setData(emptyData, defaultHeaders);
        model.setFileName("未命名.xlsx");
    }

    /**
     * 从Model获取数据并更新View
     */
    private void updateView() {
        view.updateTable(model.getHeaders(), model.getData());
        view.updateTitle(model.getFileName());
    }

    /**
     * 当Model数据变化时，自动调用此方法更新View
     */
    @Override
    public void update(Observable o, Object arg) {
        updateView();
    }
}