package model;

import java.util.ArrayList;
import java.util.List;
import java.util.Observable;

/**
 * Excel数据模型类 - MVC架构中的Model层
 * 管理Excel数据，提供数据操作方法，并通知观察者（View）数据变化
 */
public class ExcelDataModel extends Observable {
    private List<List<String>> data; // 存储Excel数据的二维列表
    private List<String> headers;    // 存储表头信息
    private String currentFileName;  // 当前文件名

    public ExcelDataModel() {
        this.data = new ArrayList<>();
        this.headers = new ArrayList<>();
        this.currentFileName = "未命名.xlsx";
    }

    /**
     * 设置数据并通知观察者
     */
    public void setData(List<List<String>> newData, List<String> newHeaders) {
        this.data = new ArrayList<>(newData);
        this.headers = new ArrayList<>(newHeaders);
        notifyObservers();
    }

    /**
     * 添加新行数据
     */
    public void addRow(List<String> rowData) {
        this.data.add(new ArrayList<>(rowData));
        notifyObservers();
    }

    /**
     * 更新指定行数据
     */
    public void updateRow(int rowIndex, List<String> newData) {
        if (rowIndex >= 0 && rowIndex < data.size()) {
            this.data.set(rowIndex, new ArrayList<>(newData));
            notifyObservers();
        }
    }

    /**
     * 删除指定行数据
     */
    public void deleteRow(int rowIndex) {
        if (rowIndex >= 0 && rowIndex < data.size()) {
            this.data.remove(rowIndex);
            notifyObservers();
        }
    }

    /**
     * 获取所有数据
     */
    public List<List<String>> getData() {
        return new ArrayList<>(data); // 返回副本以保护数据
    }

    /**
     * 获取表头信息
     */
    public List<String> getHeaders() {
        return new ArrayList<>(headers); // 返回副本以保护数据
    }

    /**
     * 设置文件名
     */
    public void setFileName(String fileName) {
        this.currentFileName = fileName;
        notifyObservers();
    }

    /**
     * 获取文件名
     */
    public String getFileName() {
        return currentFileName;
    }

    /**
     * 获取数据行数
     */
    public int getRowCount() {
        return data.size();
    }

    /**
     * 获取列数（基于表头）
     */
    public int getColumnCount() {
        return headers.size();
    }

    /**
     * 清空数据
     */
    public void clearData() {
        this.data.clear();
        this.headers.clear();
        notifyObservers();
    }

    /**
     * 通知观察者数据变化
     */
    private void notifyObservers() {
        setChanged();
        super.notifyObservers();
    }
}