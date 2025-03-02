tell application "System Events"
    tell current application
        activate
        display dialog "Word2PDF需要访问系统事件的权限来执行文档转换。\n\n1. 请在接下来的系统对话框中点击'允许'来授予权限。\n2. 如果没有看到系统对话框，请前往：\n   系统设置 > 隐私与安全性 > 辅助功能，\n   手动添加并允许Word2PDF访问权限。" buttons {"继续"} default button "继续"
    end tell
    try
        get name of application processes
    on error
        display dialog "未能获取系统事件权限。\n\n请按以下步骤操作：\n1. 打开系统设置\n2. 点击隐私与安全性\n3. 点击左侧的辅助功能\n4. 点击右侧的锁图标并验证身份\n5. 在右侧找到Word2PDF\n6. 确保Word2PDF旁边的开关已打开\n7. 重新启动应用程序" buttons {"好"} default button "好"
    end try
end tell