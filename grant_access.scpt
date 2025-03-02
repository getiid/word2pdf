tell application "System Events"
    tell current application
        activate
        display dialog "Word2PDF需要访问系统事件和文件系统的权限来执行文档转换。\n\n1. 请在接下来的系统对话框中点击'允许'来授予权限。\n2. 如果没有看到系统对话框，请前往：\n   系统设置 > 隐私与安全性 > 辅助功能、文件与文件夹、完全磁盘访问权限，\n   手动添加并允许Word2PDF访问权限。" buttons {"继续"} default button "继续"
    end tell
    try
        set UI elements enabled to true
        set systemEventsEnabled to true
    on error
        display dialog "请在系统偏好设置中授予应用程序完全磁盘访问权限。" buttons {"好的"} default button 1
    end try
    try
        get name of application processes
        tell application "Finder"
            try
                make new folder at desktop with properties {name:"Word2PDF_Test"}
                delete folder "Word2PDF_Test" of desktop
                set testFile to (path to desktop as text) & "Word2PDF_Test.txt" as alias
                try
                    open for access testFile with write permission
                    close access testFile
                    delete testFile
                on error
                    display dialog "未能获取完整的文件系统权限。\n\n请按以下步骤操作：\n1. 打开系统设置\n2. 点击隐私与安全性\n3. 点击左侧的完全磁盘访问权限\n4. 点击右侧的锁图标并验证身份\n5. 在右侧找到Word2PDF\n6. 确保Word2PDF旁边的开关已打开\n7. 重新启动应用程序" buttons {"好"} default button "好"
                end try
            on error
                display dialog "未能获取文件系统权限。\n\n请按以下步骤操作：\n1. 打开系统设置\n2. 点击隐私与安全性\n3. 点击左侧的文件与文件夹\n4. 点击右侧的锁图标并验证身份\n5. 在右侧找到Word2PDF\n6. 确保Word2PDF旁边的开关已打开\n7. 重新启动应用程序" buttons {"好"} default button "好"
            end try
        end tell
    on error
        display dialog "未能获取系统事件权限。\n\n请按以下步骤操作：\n1. 打开系统设置\n2. 点击隐私与安全性\n3. 点击左侧的辅助功能\n4. 点击右侧的锁图标并验证身份\n5. 在右侧找到Word2PDF\n6. 确保Word2PDF旁边的开关已打开\n7. 重新启动应用程序" buttons {"好"} default button "好"
    end try
end tell