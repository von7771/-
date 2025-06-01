package com.papercheck.controller;

import com.papercheck.model.CheckResult;
import com.papercheck.service.PaperFormatChecker;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.stage.FileChooser;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.List;
import java.util.Optional;

/**
 * 主界面控制器
 */
public class MainController {
    private static final Logger logger = LoggerFactory.getLogger(MainController.class);
    
    @FXML
    private Button selectFileButton;
    
    @FXML
    private Label filePathLabel;
    
    @FXML
    private Button checkButton;
    
    @FXML
    private TableView<CheckResult> resultsTable;
    
    @FXML
    private TableColumn<CheckResult, String> issueTypeColumn;
    
    @FXML
    private TableColumn<CheckResult, String> locationColumn;
    
    @FXML
    private TableColumn<CheckResult, String> descriptionColumn;
    
    @FXML
    private TableColumn<CheckResult, String> suggestionColumn;
    
    @FXML
    private ListView<String> rulesListView;
    
    @FXML
    private Label statusLabel;
    
    @FXML
    private Button exportButton;
    
    @FXML
    private Button previewButton;
    
    @FXML
    private Button fixButton;
    
    private File selectedFile;
    private PaperFormatChecker checker;
    private ObservableList<CheckResult> checkResults = FXCollections.observableArrayList();
    
    @FXML
    public void initialize() {
        // 初始化表格列
        issueTypeColumn.setCellValueFactory(new PropertyValueFactory<>("issueType"));
        locationColumn.setCellValueFactory(new PropertyValueFactory<>("location"));
        descriptionColumn.setCellValueFactory(new PropertyValueFactory<>("description"));
        suggestionColumn.setCellValueFactory(new PropertyValueFactory<>("suggestion"));
        
        resultsTable.setItems(checkResults);
        
        // 初始化检查规则列表
        ObservableList<String> rules = FXCollections.observableArrayList(
            "1. 标题格式: 论文标题应使用黑体、三号字",
            "2. 正文格式: 正文应使用宋体、小四号字",
            "3. 段落格式: 段落首行缩进2字符",
            "4. 页边距: 上下2.5cm，左右3.0cm",
            "5. 行间距: 1.5倍行距",
            "6. 页码: 页码应位于页面底部居中",
            "7. 图表标题: 图表标题应居中显示",
            "8. 参考文献格式: 应符合GB/T 7714-2015标准"
        );
        rulesListView.setItems(rules);
        
        // 初始化检查器
        checker = new PaperFormatChecker();
        
        // 初始化按钮状态
        previewButton.setDisable(true);
        fixButton.setDisable(true);
    }
    
    @FXML
    public void handleSelectFile() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("选择Word文档");
        fileChooser.getExtensionFilters().addAll(
            new FileChooser.ExtensionFilter("Word文档", "*.docx", "*.doc")
        );
        
        selectedFile = fileChooser.showOpenDialog(selectFileButton.getScene().getWindow());
        
        if (selectedFile != null) {
            filePathLabel.setText(selectedFile.getPath());
            checkButton.setDisable(false);
            statusLabel.setText("已选择文件: " + selectedFile.getName());
            logger.info("已选择文件: {}", selectedFile.getPath());
        }
    }
    
    @FXML
    public void handleCheck() {
        if (selectedFile == null) {
            showAlert(Alert.AlertType.ERROR, "错误", "请先选择一个Word文档");
            return;
        }
        
        statusLabel.setText("正在检查...");
        checkButton.setDisable(true);
        
        // 清空之前的结果
        checkResults.clear();
        
        try {
            // 执行检查
            List<CheckResult> results = checker.checkDocument(selectedFile);
            
            // 显示结果
            checkResults.addAll(results);
            
            statusLabel.setText("检查完成，发现 " + results.size() + " 个问题");
            exportButton.setDisable(results.isEmpty());
            
            // 如果有问题，启用预览和修复按钮
            previewButton.setDisable(results.isEmpty());
            fixButton.setDisable(results.isEmpty());
            
            if (results.isEmpty()) {
                showAlert(Alert.AlertType.INFORMATION, "检查结果", "恭喜！未发现格式问题。");
            }
            
            logger.info("文档检查完成，发现 {} 个问题", results.size());
        } catch (Exception e) {
            logger.error("检查文档时出错", e);
            statusLabel.setText("检查失败");
            showAlert(Alert.AlertType.ERROR, "错误", "检查文档时出错: " + e.getMessage());
        } finally {
            checkButton.setDisable(false);
        }
    }
    
    @FXML
    public void handlePreview() {
        if (selectedFile == null || checkResults.isEmpty()) {
            showAlert(Alert.AlertType.ERROR, "错误", "请先检查文档并确保有需要修复的问题");
            return;
        }
        
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("保存预览文档");
        fileChooser.getExtensionFilters().add(
            new FileChooser.ExtensionFilter("Word文档", "*.docx")
        );
        
        // 设置默认文件名
        String originalName = selectedFile.getName();
        String previewName = originalName.substring(0, originalName.lastIndexOf('.')) + "_预览.docx";
        fileChooser.setInitialFileName(previewName);
        
        File previewFile = fileChooser.showSaveDialog(previewButton.getScene().getWindow());
        
        if (previewFile != null) {
            try {
                statusLabel.setText("正在创建预览...");
                previewButton.setDisable(true);
                
                // 创建预览文档
                List<CheckResult> fixedResults = checker.createFixedDocumentPreview(selectedFile, previewFile);
                
                statusLabel.setText("预览文档已保存至: " + previewFile.getPath());
                logger.info("预览文档已保存至: {}", previewFile.getPath());
                
                // 显示预览结果
                Alert alert = new Alert(Alert.AlertType.INFORMATION);
                alert.setTitle("预览文档已创建");
                alert.setHeaderText("已修复 " + fixedResults.size() + " 个问题");
                alert.setContentText("预览文档已保存至:\n" + previewFile.getPath() + 
                                    "\n\n您可以打开此文档查看修复效果，然后决定是否应用这些修改。");
                alert.showAndWait();
            } catch (Exception e) {
                logger.error("创建预览文档时出错", e);
                statusLabel.setText("创建预览失败");
                showAlert(Alert.AlertType.ERROR, "错误", "创建预览文档时出错: " + e.getMessage());
            } finally {
                previewButton.setDisable(false);
            }
        }
    }
    
    @FXML
    public void handleFix() {
        if (selectedFile == null || checkResults.isEmpty()) {
            showAlert(Alert.AlertType.ERROR, "错误", "请先检查文档并确保有需要修复的问题");
            return;
        }
        
        // 确认是否修复
        Alert confirmAlert = new Alert(Alert.AlertType.CONFIRMATION);
        confirmAlert.setTitle("确认修复");
        confirmAlert.setHeaderText("您确定要修复文档格式问题吗？");
        confirmAlert.setContentText("此操作将修改原始文档的格式。建议在修复前备份原始文档。");
        
        Optional<ButtonType> result = confirmAlert.showAndWait();
        if (result.isPresent() && result.get() == ButtonType.OK) {
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("保存修复后的文档");
            fileChooser.getExtensionFilters().add(
                new FileChooser.ExtensionFilter("Word文档", "*.docx")
            );
            
            // 设置默认文件名
            String originalName = selectedFile.getName();
            String fixedName = originalName.substring(0, originalName.lastIndexOf('.')) + "_已修复.docx";
            fileChooser.setInitialFileName(fixedName);
            
            File fixedFile = fileChooser.showSaveDialog(fixButton.getScene().getWindow());
            
            if (fixedFile != null) {
                try {
                    statusLabel.setText("正在修复...");
                    fixButton.setDisable(true);
                    
                    // 修复文档
                    List<CheckResult> fixedResults = checker.fixDocument(selectedFile, fixedFile);
                    
                    statusLabel.setText("文档已修复并保存至: " + fixedFile.getPath());
                    logger.info("修复后的文档已保存至: {}", fixedFile.getPath());
                    
                    // 显示修复结果
                    showAlert(Alert.AlertType.INFORMATION, "修复完成", 
                              "已修复 " + fixedResults.size() + " 个问题\n" +
                              "修复后的文档已保存至:\n" + fixedFile.getPath());
                    
                    // 更新检查结果（清空，因为问题已修复）
                    checkResults.clear();
                    previewButton.setDisable(true);
                    fixButton.setDisable(true);
                    exportButton.setDisable(true);
                } catch (Exception e) {
                    logger.error("修复文档时出错", e);
                    statusLabel.setText("修复失败");
                    showAlert(Alert.AlertType.ERROR, "错误", "修复文档时出错: " + e.getMessage());
                } finally {
                    fixButton.setDisable(false);
                }
            }
        }
    }
    
    @FXML
    public void handleExport() {
        if (checkResults.isEmpty()) {
            showAlert(Alert.AlertType.INFORMATION, "导出报告", "没有检查结果可导出");
            return;
        }
        
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("导出检查报告");
        fileChooser.getExtensionFilters().add(
            new FileChooser.ExtensionFilter("HTML文件", "*.html")
        );
        fileChooser.setInitialFileName("论文格式检查报告.html");
        
        File file = fileChooser.showSaveDialog(exportButton.getScene().getWindow());
        
        if (file != null) {
            try {
                checker.exportReportToHtml(checkResults, file);
                statusLabel.setText("报告已导出至: " + file.getPath());
                logger.info("报告已导出至: {}", file.getPath());
                
                showAlert(Alert.AlertType.INFORMATION, "导出成功", "检查报告已成功导出至: " + file.getPath());
            } catch (Exception e) {
                logger.error("导出报告时出错", e);
                statusLabel.setText("导出失败");
                showAlert(Alert.AlertType.ERROR, "错误", "导出报告时出错: " + e.getMessage());
            }
        }
    }
    
    private void showAlert(Alert.AlertType alertType, String title, String content) {
        Alert alert = new Alert(alertType);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(content);
        alert.showAndWait();
    }
} 