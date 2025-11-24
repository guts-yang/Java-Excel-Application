import javax.swing.*;
import javax.swing.table.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.ArrayList;
import java.util.List;

/**
 * 简单的Excel应用测试类
 * 不依赖外部库，用于验证项目结构和GUI功能
 */
public class TestApp {
    private JFrame frame;
    private JTable table;
    private DefaultTableModel tableModel;
    private JTextField statusBar;
    
    public static void main(String[] args) {
        // 在事件调度线程中启动GUI
        SwingUtilities.invokeLater(() -> {
            TestApp app = new TestApp();
            app.createAndShowGUI();
        });
    }
    
    private void createAndShowGUI() {
        // 创建主窗口
        frame = new JFrame("Excel处理应用 - 测试版");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);
        
        // 创建菜单栏
        JMenuBar menuBar = new JMenuBar();
        JMenu fileMenu = new JMenu("文件");
        JMenu editMenu = new JMenu("编辑");
        
        JMenuItem newItem = new JMenuItem("新建");
        newItem.addActionListener(e -> createNewFile());
        fileMenu.add(newItem);
        
        JMenuItem exitItem = new JMenuItem("退出");
        exitItem.addActionListener(e -> System.exit(0));
        fileMenu.add(exitItem);
        
        JMenuItem addRowItem = new JMenuItem("添加行");
        addRowItem.addActionListener(e -> addRow());
        editMenu.add(addRowItem);
        
        menuBar.add(fileMenu);
        menuBar.add(editMenu);
        frame.setJMenuBar(menuBar);
        
        // 创建表格
        String[] columnNames = {"列1", "列2", "列3", "列4"};
        tableModel = new MyTableModel(columnNames, 0);
        table = new JTable(tableModel);
        
        // 添加示例数据
        tableModel.addRow(new Object[]{"示例1", "数据1", "值1", "内容1"});
        tableModel.addRow(new Object[]{"示例2", "数据2", "值2", "内容2"});
        
        JScrollPane scrollPane = new JScrollPane(table);
        frame.add(scrollPane, BorderLayout.CENTER);
        
        // 创建状态栏
        statusBar = new JTextField("就绪 - 测试模式");
        statusBar.setEditable(false);
        frame.add(statusBar, BorderLayout.SOUTH);
        
        // 创建工具栏
        JToolBar toolBar = new JToolBar();
        JButton newButton = new JButton("新建");
        newButton.addActionListener(e -> createNewFile());
        toolBar.add(newButton);
        
        JButton addButton = new JButton("添加行");
        addButton.addActionListener(e -> addRow());
        toolBar.add(addButton);
        
        frame.add(toolBar, BorderLayout.NORTH);
        
        // 显示窗口
        frame.setVisible(true);
        System.out.println("测试应用已启动！");
    }
    
    private void createNewFile() {
        int confirm = JOptionPane.showConfirmDialog(frame, "确定要创建新文件吗？");
        if (confirm == JOptionPane.YES_OPTION) {
            tableModel.setRowCount(0);
            statusBar.setText("已创建新文件");
        }
    }
    
    private void addRow() {
        tableModel.addRow(new Object[]{"", "", "", ""});
        statusBar.setText("已添加新行");
    }
    
    // 使用Swing自带的表格模型类
    class MyTableModel extends DefaultTableModel {
        public MyTableModel(String[] columnNames, int rowCount) {
            super(columnNames, rowCount);
        }
        
        @Override
        public boolean isCellEditable(int row, int column) {
            return true; // 所有单元格都可编辑
        }
    }
}