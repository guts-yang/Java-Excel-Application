import java.io.*;
import java.nio.file.*;
import java.util.*;

public class run_app {
    public static void main(String[] args) {
        try {
            // 设置项目根目录
            Path projectDir = Paths.get(".").toAbsolutePath().normalize();
            System.out.println("项目目录: " + projectDir);
            
            // 创建输出目录
            Path outputDir = projectDir.resolve("bin");
            if (!Files.exists(outputDir)) {
                Files.createDirectories(outputDir);
                System.out.println("创建输出目录: " + outputDir);
            }
            
            // 创建lib目录（如果需要存放依赖）
            Path libDir = projectDir.resolve("lib");
            if (!Files.exists(libDir)) {
                Files.createDirectories(libDir);
                System.out.println("创建lib目录: " + libDir);
            }
            
            // 复制Maven依赖到lib目录（如果有）
            copyMavenDependencies(libDir);
            
            // 编译项目
            compileProject(projectDir, outputDir);
            
            // 运行应用程序
            runApplication(outputDir);
            
        } catch (Exception e) {
            System.err.println("执行过程中出错:");
            e.printStackTrace();
        }
    }
    
    private static void copyMavenDependencies(Path libDir) throws IOException {
        // 从pom.xml中读取依赖并提示用户下载
        System.out.println("请确保已下载以下依赖到lib目录:");
        System.out.println("- poi-5.2.3.jar");
        System.out.println("- poi-ooxml-5.2.3.jar");
        System.out.println("- poi-ooxml-lite-5.2.3.jar");
        System.out.println("- xmlbeans-5.2.0.jar");
        System.out.println("- commons-compress-1.23.0.jar");
        System.out.println("- commons-collections4-4.4.jar");
    }
    
    private static void compileProject(Path projectDir, Path outputDir) throws IOException, InterruptedException {
        // 查找所有Java源文件
        List<String> javaFiles = new ArrayList<>();
        Path srcDir = projectDir.resolve("src").resolve("main").resolve("java");
        System.out.println("源代码目录: " + srcDir);
        
        Files.walk(srcDir)
             .filter(Files::isRegularFile)
             .filter(p -> p.toString().endsWith(".java"))
             .forEach(p -> {
                 String path = p.toString();
                 System.out.println("找到源文件: " + path);
                 javaFiles.add(path);
             });
        
        if (javaFiles.isEmpty()) {
            System.err.println("未找到Java源文件！");
            System.exit(1);
        }
        
        // 构建编译命令
        List<String> command = new ArrayList<>();
        command.add("javac");
        command.add("-d");
        command.add(outputDir.toString());
        command.add("-cp");
        command.add(outputDir + ";lib/*");
        command.addAll(javaFiles);
        
        System.out.println("执行编译命令: javac -d " + outputDir + " -cp " + (outputDir + ";lib/*") + " [源文件列表]");
        
        // 使用ProcessBuilder替代Runtime.exec
        ProcessBuilder pb = new ProcessBuilder(command);
        pb.redirectErrorStream(true);
        Process process = pb.start();
        
        printProcessOutput(process);
        
        int exitCode = process.waitFor();
        if (exitCode == 0) {
            System.out.println("编译成功!");
        } else {
            System.out.println("编译失败，退出码: " + exitCode);
            System.exit(1);
        }
    }
    
    private static void runApplication(Path outputDir) throws IOException {
        // 使用ProcessBuilder启动应用
        List<String> command = new ArrayList<>();
        command.add("java");
        command.add("-cp");
        command.add(outputDir + ";lib/*");
        command.add("view.MainApp");
        
        System.out.println("执行运行命令: java -cp " + (outputDir + ";lib/*") + " view.MainApp");
        
        ProcessBuilder pb = new ProcessBuilder(command);
        pb.redirectErrorStream(true);
        Process process = pb.start();
        
        // 启动新线程处理输出，避免阻塞
        new Thread(() -> {
            try {
                printProcessOutput(process);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }).start();
        
        System.out.println("应用程序已启动...");
    }
    
    private static void printProcessOutput(Process process) throws IOException {
        // 打印标准输出
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()))) {
            String line;
            while ((line = reader.readLine()) != null) {
                System.out.println(line);
            }
        }
        
        // 打印错误输出
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getErrorStream()))) {
            String line;
            while ((line = reader.readLine()) != null) {
                System.err.println(line);
            }
        }
    }
}