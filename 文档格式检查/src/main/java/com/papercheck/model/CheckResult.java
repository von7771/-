package com.papercheck.model;

/**
 * 论文格式检查结果
 */
public class CheckResult {
    private String issueType;    // 问题类型
    private String location;     // 问题位置
    private String description;  // 问题描述
    private String suggestion;   // 修改建议

    public CheckResult() {
    }

    public CheckResult(String issueType, String location, String description, String suggestion) {
        this.issueType = issueType;
        this.location = location;
        this.description = description;
        this.suggestion = suggestion;
    }

    public String getIssueType() {
        return issueType;
    }

    public void setIssueType(String issueType) {
        this.issueType = issueType;
    }

    public String getLocation() {
        return location;
    }

    public void setLocation(String location) {
        this.location = location;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public String getSuggestion() {
        return suggestion;
    }

    public void setSuggestion(String suggestion) {
        this.suggestion = suggestion;
    }

    @Override
    public String toString() {
        return "CheckResult{" +
                "issueType='" + issueType + '\'' +
                ", location='" + location + '\'' +
                ", description='" + description + '\'' +
                ", suggestion='" + suggestion + '\'' +
                '}';
    }
} 