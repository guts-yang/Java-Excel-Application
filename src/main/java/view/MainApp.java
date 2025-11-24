package view;

import controller.ExcelController;
import model.ExcelDataModel;

import javax.swing.*;
import java.awt.*;

/**
 * 主应用程序入口类
 * MVC架构中的View层组件，负责启动应用程序
 */
public class MainApp {
    public static void main(String[] args) {
        // 在事件调度线程中启动GUI
        SwingUtilities.invokeLater(() -> {
            // 初始化Model层
            ExcelDataModel model = new ExcelDataModel();
            
            // 初始化View层
            ExcelView view = new ExcelView();
            
            // 初始化Controller层，连接Model和View
            ExcelController controller = new ExcelController(model, view);
            
            // 设置View的可见性
            view.setVisible(true);
        });
    }
}