package view;

import controller.ExcelController;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel视图类 - MVC架构中的View层
 * 负责用户界面展示和用户输入接收
 */
public class ExcelView extends JFrame {
    private JTable table;
    private DefaultTableModel tableModel;
    private ExcelController controller;
    private JToolBar toolBar;

    /**
     * 构造函数，初始化用户界面
     */
    public ExcelView() {
        super("Excel处理应用程序");
        initUI();
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(800, 600);
        setLocationRelativeTo(null); // 窗口居中显示
    }

    /**
     * 设置控制器
     */
    public void setController(ExcelController controller) {
        this.controller = controller;
    }

    /**
     * 初始化用户界面
     */
    private void initUI() {
        // 创建菜单栏
        JMenuBar menuBar = new JMenuBar();
        setJMenuBar(menuBar);

        // 文件菜单
        JMenu fileMenu = new JMenu("文件");
        menuBar.add(fileMenu);

        JMenuItem newItem = new JMenuItem("新建");
        newItem.addActionListener(e -> handleNewFile());
        fileMenu.add(newItem);

        JMenuItem importItem = new JMenuItem("导入");
        importItem.addActionListener(e -> handleImport());
        fileMenu.add(importItem);

        JMenuItem exportItem = new JMenuItem("导出");
        exportItem.addActionListener(e -> handleExport());
        fileMenu.add(exportItem);

        fileMenu.addSeparator();

        JMenuItem exitItem = new JMenuItem("退出");
        exitItem.addActionListener(e -> System.exit(0));
        fileMenu.add(exitItem);

        // 编辑菜单
        JMenu editMenu = new JMenu("编辑");
        menuBar.add(editMenu);

        JMenuItem addRowItem = new JMenuItem("添加行");
        addRowItem.addActionListener(e -> handleAddRow());
        editMenu.add(addRowItem);

        JMenuItem deleteRowItem = new JMenuItem("删除选中行");
        deleteRowItem.addActionListener(e -> handleDeleteRow());
        editMenu.add(deleteRowItem);

        JMenuItem clearItem = new JMenuItem("清空所有");
        clearItem.addActionListener(e -> handleClear());
        editMenu.add(clearItem);

        // 创建工具栏
        toolBar = new JToolBar();
        toolBar.setFloatable(false);
        add(toolBar, BorderLayout.NORTH);

        JButton newButton = new JButton("新建");
        newButton.addActionListener(e -> handleNewFile());
        toolBar.add(newButton);

        JButton importButton = new JButton("导入");
        importButton.addActionListener(e -> handleImport());
        toolBar.add(importButton);

        JButton exportButton = new JButton("导出");
        exportButton.addActionListener(e -> handleExport());
        toolBar.add(exportButton);

        toolBar.addSeparator();

        JButton addRowButton = new JButton("添加行");
        addRowButton.addActionListener(e -> handleAddRow());
        toolBar.add(addRowButton);

        JButton deleteRowButton = new JButton("删除行");
        deleteRowButton.addActionListener(e -> handleDeleteRow());
        toolBar.add(deleteRowButton);

        // 创建表格
        tableModel = new DefaultTableModel() {
            @Override
            public boolean isCellEditable(int row, int column) {
                return true; // 允许单元格编辑
            }
        };

        table = new JTable(tableModel);
        table.getSelectionModel().addListSelectionListener(e -> {
            // 当用户编辑单元格时，更新数据
            table.getModel().addTableModelListener(evt -> {
                int row = evt.getFirstRow();
                int column = evt.getColumn();
                if (row >= 0 && column >= 0) {
                    updateRowData(row);
                }
            });
        });

        // 添加右键菜单
        JPopupMenu popupMenu = new JPopupMenu();
        JMenuItem addRowPopup = new JMenuItem("添加行");
        addRowPopup.addActionListener(e -> handleAddRow());
        popupMenu.add(addRowPopup);

        JMenuItem deleteRowPopup = new JMenuItem("删除行");
        deleteRowPopup.addActionListener(e -> handleDeleteRow());
        popupMenu.add(deleteRowPopup);

        table.setComponentPopupMenu(popupMenu);

        // 添加表格到滚动面板
        JScrollPane scrollPane = new JScrollPane(table);
        add(scrollPane, BorderLayout.CENTER);

        // 创建状态栏
        JLabel statusLabel = new JLabel("就绪");
        add(statusLabel, BorderLayout.SOUTH);

        // 创建默认空白表格
        updateTable(new ArrayList<>(), new ArrayList<>());
    }

    /**
     * 更新表格数据
     */
    public void updateTable(List<String> headers, List<List<String>> data) {
        // 清空现有表格模型
        tableModel.setRowCount(0);
        tableModel.setColumnCount(0);

        // 设置表头
        if (!headers.isEmpty()) {
            tableModel.setColumnIdentifiers(headers.toArray());
        }

        // 添加数据行
        for (List<String> row : data) {
            tableModel.addRow(row.toArray());
        }
    }

    /**
     * 更新窗口标题
     */
    public void updateTitle(String fileName) {
        setTitle("Excel处理应用程序 - " + fileName);
    }

    /**
     * 处理新建文件
     */
    private void handleNewFile() {
        if (controller != null) {
            controller.createNewTable();
        }
    }

    /**
     * 处理导入Excel
     */
    private void handleImport() {
        if (controller != null) {
            controller.importExcel();
        }
    }

    /**
     * 处理导出Excel
     */
    private void handleExport() {
        if (controller != null) {
            controller.exportExcel();
        }
    }

    /**
     * 处理添加行
     */
    private void handleAddRow() {
        if (controller != null && tableModel.getColumnCount() > 0) {
            List<String> newRow = new ArrayList<>();
            for (int i = 0; i < tableModel.getColumnCount(); i++) {
                newRow.add(""); // 添加空单元格
            }
            controller.addRow(newRow);
            // 选中新添加的行
            int lastRow = tableModel.getRowCount() - 1;
            table.setRowSelectionInterval(lastRow, lastRow);
            table.editCellAt(lastRow, 0);
        } else {
            JOptionPane.showMessageDialog(this, "请先创建或导入表格", "提示", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    /**
     * 处理删除行
     */
    private void handleDeleteRow() {
        if (controller != null) {
            int selectedRow = table.getSelectedRow();
            if (selectedRow >= 0) {
                controller.deleteRow(selectedRow);
            } else {
                JOptionPane.showMessageDialog(this, "请先选择要删除的行", "提示", JOptionPane.INFORMATION_MESSAGE);
            }
        }
    }

    /**
     * 处理清空所有数据
     */
    private void handleClear() {
        if (controller != null) {
            controller.clearAllData();
        }
    }

    /**
     * 更新行数据到模型
     */
    private void updateRowData(int rowIndex) {
        if (controller != null && rowIndex >= 0 && rowIndex < tableModel.getRowCount()) {
            List<String> rowData = new ArrayList<>();
            for (int i = 0; i < tableModel.getColumnCount(); i++) {
                Object value = tableModel.getValueAt(rowIndex, i);
                rowData.add(value != null ? value.toString() : "");
            }
            controller.updateRow(rowIndex, rowData);
        }
    }
}