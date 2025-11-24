# Java Excel处理应用程序（MVC架构）

## 项目概述

本项目是一个基于Java的Excel处理应用程序，通过实现简单的Excel导入、导出功能来演示和学习MVC（Model-View-Controller）架构模式。项目明确划分了Model（数据模型层）、View（视图层）和Controller（控制层），展示了各层之间的职责划分与交互流程。

## MVC架构实现

### 1. Model层（数据模型层）

Model层负责数据管理和业务逻辑，包括：

- **ExcelDataModel.java**：核心数据模型类，管理Excel数据，提供数据操作方法，并通知观察者（View）数据变化
- **ExcelDataAccess.java**：数据访问类，负责Excel文件的读取和写入操作，实现数据持久化

Model层的主要职责：
- 维护应用程序的数据状态
- 提供数据访问和修改的方法
- 通过观察者模式通知View层数据变化
- 处理数据的持久化（Excel文件读写）

### 2. View层（视图层）

View层负责用户界面展示和用户输入接收，包括：

- **ExcelView.java**：主界面类，包含表格显示、菜单、工具栏等组件
- **MainApp.java**：应用程序入口，初始化MVC各层并启动应用

View层的主要职责：
- 展示数据给用户
- 接收用户输入
- 调用Controller层处理用户请求
- 响应Model层的数据变化更新界面

### 3. Controller层（控制层）

Controller层负责处理用户请求，协调Model和View之间的交互：

- **ExcelController.java**：控制器类，处理用户操作，调用Model层方法，并更新View层
- **ExcelControllerTest.java**：控制器测试类，用于测试MVC架构交互功能

Controller层的主要职责：
- 接收View层传递的用户请求
- 调用Model层方法进行数据操作
- 将数据变化通知View层更新
- 处理业务逻辑和流程控制

## 项目结构

```
src/main/java/
├── model/               # Model层
│   ├── ExcelDataModel.java      # 数据模型类
│   └── ExcelDataAccess.java     # 数据访问类
├── view/                # View层
│   ├── ExcelView.java           # 视图类
│   └── MainApp.java             # 应用入口
└── controller/          # Controller层
    ├── ExcelController.java     # 控制器类
    └── ExcelControllerTest.java # 测试类
```

## 功能特性

1. **Excel导入**：支持从XLSX文件导入数据到应用程序
2. **Excel导出**：支持将应用程序中的数据导出为XLSX文件
3. **数据编辑**：支持在表格中直接编辑数据
4. **行操作**：支持添加、删除、更新行数据
5. **新建表格**：支持创建新的空白表格
6. **清空数据**：支持清空所有数据

## 技术栈

- **Java 8**：开发语言
- **Apache POI 5.2.3**：Excel文件处理库
- **Swing**：图形用户界面框架
- **Maven**：项目管理和依赖管理

## 使用方法

### 1. 环境准备

- JDK 1.8或更高版本
- Maven 3.0或更高版本

### 2. 编译和运行

#### 使用Maven命令

```bash
# 进入项目目录
cd d:\coding\java\class_1\m11d24\project1

# 下载依赖
mvn clean install

# 运行应用程序
mvn exec:java -Dexec.mainClass="view.MainApp"
```

#### 使用批处理脚本（Windows）

```bash
# 在项目目录中运行
compile_and_run.bat
```

### 3. 功能使用

- **新建表格**：点击"新建"按钮或菜单，创建一个包含默认列的空白表格
- **导入Excel**：点击"导入"按钮，选择XLSX文件导入数据
- **导出Excel**：点击"导出"按钮，选择保存位置导出数据
- **添加行**：点击"添加行"按钮，在表格末尾添加新行
- **删除行**：选中行后点击"删除行"按钮，删除选中的行
- **编辑数据**：直接在表格单元格中编辑数据

## MVC架构优势演示

本项目通过以下方式展示MVC架构的优势：

1. **关注点分离**：Model专注于数据，View专注于界面，Controller专注于控制逻辑
2. **低耦合**：各层之间通过明确的接口交互，降低依赖
3. **高内聚**：每层内部职责单一明确
4. **易维护性**：修改某一层不会影响其他层
5. **可扩展性**：可以独立扩展各层功能

## 扩展建议

1. 添加数据验证功能
2. 支持更多Excel格式（如XLS）
3. 添加单元格格式设置功能
4. 实现撤销/重做功能
5. 添加数据筛选和排序功能

## 作者

MVC架构学习项目

## 日期

2024年