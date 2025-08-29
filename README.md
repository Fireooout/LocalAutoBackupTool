# 本地文件自动备份助手 (Local File Auto Backup Assistant) v1.2.1

### 1. 软件简介
本地文件自动备份助手是一款**轻量级桌面应用**，旨在帮助用户轻松实现重要文件和文件夹，游戏存档等的自动备份与管理。通过简单的设置，您可以指定需要备份的文件/目录、备份存储位置和自动备份间隔，软件将在后台自动执行备份操作，确保您的数据安全。


### 2. 主要功能
- **多源备份支持**：同时备份多个文件和多个目录，满足复杂备份需求
- **自定义备份间隔**：根据需求设置自动备份的时间间隔，灵活适配不同场景
- **智能备份清理**：限制最大备份数量，超过上限时自动删除最旧备份，节省存储空间
- **快捷键手动备份**：支持自定义全局快捷键，一键触发手动备份，操作更高效
- **完整备份管理**：可视化查看所有备份记录，支持删除无用备份、还原所需备份
- **托盘后台运行**：最小化到系统托盘，不占用桌面空间，后台持续守护数据安全
- **双语界面切换**：一键切换中文/英文界面，适配不同用户使用习惯

### 3. 更新日志
#### 版本 1.1.0
- 修复打包过程中单文件版本被清理的问题
- 更新托盘图标为folder-sync.png和folder-sync-start.png
- 修复托盘菜单启用自动备份时创建多个图标的问题
- 优化打包配置，确保图标文件正确打包
- 其他性能和稳定性改进

### 4. 使用说明
#### 3.1 基本设置
##### 3.1.1 添加源路径
- 点击「添加文件」按钮：选择单个或多个需要备份的文件
- 点击「添加目录」按钮：选择需要备份的文件夹（含子目录所有内容）
- 移除路径：在源路径列表中选中不需要的路径，点击「移除」按钮删除
- 查看位置：选中路径后点击「打开位置」，可直接跳转至该路径的实际存储位置

##### 3.1.2 设置备份目录
- 直接输入：在「备份目录」输入框中手动输入目标路径
- 可视化选择：点击「浏览」按钮，通过文件管理器选择备份存储目录
- 验证目录：点击「打开位置」，确认当前设置的备份目录是否正确

##### 3.1.3 配置备份参数
- **自动保存间隔**：输入数字（单位：秒），设置自动备份的时间周期（如300代表5分钟）
- **最大备份数量**：输入数字，设置保留的备份份数上限（超过将自动清理旧备份）
- **快捷键设置**：
  1. 在输入框中按格式输入快捷键（如`ctrl+F1`）
  2. 勾选「启用快捷键」选项，激活全局快捷键功能

##### 3.1.4 高级设置（1.2.1优化）
在「高级设置」标签页中，可配置更多备份选项：
- **备份前缀设置**：自定义备份文件的前缀命名规则
- **后缀类型选择**：可选择时间戳、序号或自定义后缀
  - 时间戳：使用日期时间作为后缀（默认方式）
  - 序号：使用递增序号作为后缀
  - 自定义：支持用户自定义后缀格式，可包含日期时间格式
- **重名文件处理**：设置遇到重名备份文件时的处理方式
  - 覆盖：直接覆盖同名备份文件
  - 重命名：自动为重名文件添加序号后缀
  - 跳过：保留原有备份，跳过当前备份
- **备份计数器管理**：查看和重置备份序号计数器

#### 3.2 开始备份
- **自动备份**：点击「开始自动备份」按钮，按钮将切换为「停止自动备份」，软件进入后台自动备份模式
- **手动备份**：无论是否开启自动备份，点击「手动备份」按钮可立即执行一次完整备份

#### 3.3 备份管理
在「备份管理」标签页中，可对所有备份记录进行操作：
| 功能按钮       | 作用说明                                                                 |
|----------------|--------------------------------------------------------------------------|
| 刷新列表       | 更新备份记录，确保显示最新的备份数据                                     |
| 删除选中       | 删除选中的备份记录（不可逆，建议确认后操作）                             |
| 还原选中       | 将选中的备份还原到原始文件/目录位置（会覆盖现有文件，需谨慎操作）         |
| 打开位置       | 跳转至选中备份的实际存储目录，查看备份文件详情                           |

#### 3.4 其他功能
- **界面语言切换**：点击窗口右上角的「中文」或「English」按钮，实时切换界面语言
- **最小化到托盘**：点击窗口关闭按钮时，程序不会退出，而是最小化到系统托盘继续运行
- **托盘菜单操作**：
  1. 点击托盘图标，显示操作菜单
  2. 选择「显示窗口」：恢复程序主窗口
  3. 选择「启用自动备份」：切换自动备份状态
  4. 选择「退出」：完全关闭程序

### 5. 注意事项
1. **路径限制**：源路径（需备份的文件/目录）与备份路径不能相同，也不能存在包含关系（如源路径为`D:\Data`，备份路径不能为`D:\Data\Backup`）
2. **存储检查**：确保备份目录所在磁盘有足够的存储空间，避免因空间不足导致备份失败
3. **数据安全**：执行「还原选中」操作前，建议确认原始文件状态，避免误操作覆盖重要数据
4. **快捷键冲突**：若设置的快捷键无法使用，可能与其他程序快捷键冲突，建议更换组合（如`ctrl+alt+S`）
5. **配置文件**：程序会自动创建配置文件"backup_config.ini"保存设置，请勿随意删除或修改
6. **日志文件**：程序运行时会生成日志文件"backup_tool.log"记录操作和错误信息，便于问题排查

### 6. 反馈与支持
如有任何问题或建议，请在项目页面提交 issue，开发者将尽快回复处理。


## 二、English Version

### 1. Software Introduction
Local File Auto Backup Assistant is a **lightweight desktop application** designed to help users easily implement automatic backup and management of important files and folders. With simple settings, you can specify the files/folders to be backed up, backup storage location, and automatic backup interval. The software will automatically perform backup operations in the background to ensure your data security.

### 2. Key Features
- **Multi-source Backup Support**: Backup multiple files and directories simultaneously to meet complex backup needs
- **Custom Backup Interval**: Set the automatic backup interval according to requirements, flexibly adapting to different scenarios
- **Smart Backup Cleanup**: Limit the maximum number of backups; automatically delete the oldest backups when exceeding the limit to save storage space
- **Hotkey-triggered Manual Backup**: Support custom global hotkeys to trigger manual backup with one click for more efficient operation
- **Complete Backup Management**: Visually view all backup records, support deleting useless backups and restoring required backups
- **Tray Background Operation**: Minimize to the system tray without occupying desktop space, and continuously protect data security in the background
- **Bilingual Interface Switch**: One-click switch between Chinese/English interface to adapt to different user habits

### 3. Update Log
#### Version 1.1.0
- Fixed issue where single-file version was being cleaned up during packaging
- Updated tray icons to folder-sync.png and folder-sync-start.png
- Fixed issue where multiple tray icons were created when enabling auto-backup from tray menu
- Optimized packaging configuration to ensure icon files are correctly packaged
- Other performance and stability improvements

### 4. User Guide
#### 3.1 Basic Settings
##### 3.1.1 Add Source Paths
- Click the **Add File** button: Select single or multiple files to be backed up
- Click the **Add Folder** button: Select the folder to be backed up (including all contents in subdirectories)
- Remove Path: Select the unwanted path in the source path list and click the **Remove** button to delete it
- View Location: Select a path and click **Open Location** to jump directly to the actual storage location of the path

##### 3.1.2 Set Backup Directory
- Direct Input: Manually enter the target path in the **Backup Directory** input box
- Visual Selection: Click the **Browse** button to select the backup storage directory through the file manager
- Verify Directory: Click **Open Location** to confirm whether the currently set backup directory is correct

##### 3.1.3 Configure Backup Parameters
- **Auto-save Interval**: Enter a number (unit: seconds) to set the time cycle of automatic backup (e.g., 300 means 5 minutes)
- **Max Backup Count**: Enter a number to set the maximum number of retained backups (old backups will be automatically cleaned up when exceeded)
- **Hotkey Settings**:
  1. Enter the hotkey in the input box according to the format (e.g., `ctrl+F1`)
  2. Check the **Enable Hotkey** option to activate the global hotkey function

##### 3.1.4 Advanced Settings (Enhanced in v1.2.1)
In the "Advanced Settings" tab, you can configure more backup options:
- **Backup Prefix Settings**: Customize the prefix naming rule for backup files
- **Suffix Type Selection**: Choose from timestamp, sequential number, or custom suffix
  - Timestamp: Use date and time as suffix (default method)
  - Sequential Number: Use incremental sequential number as suffix
  - Custom: Support user-defined suffix format, which can include date and time format
- **Duplicate File Handling**: Set the handling method when encountering duplicate backup files
  - Overwrite: Directly overwrite the backup file with the same name
  - Rename: Automatically add a sequential number suffix to duplicate files
  - Skip: Keep the existing backup and skip the current backup
- **Backup Counter Management**: View and reset the backup sequential number counter

#### 3.2 Start Backup
- **Auto Backup**: Click the **Start Auto Backup** button, which will switch to **Stop Auto Backup**, and the software enters the background automatic backup mode
- **Manual Backup**: Regardless of whether auto backup is enabled, click the **Manual Backup** button to perform a complete backup immediately

#### 3.3 Backup Management
In the **Backup Management** tab, you can operate on all backup records:
| Function Button   | Description                                                              |
|-------------------|--------------------------------------------------------------------------|
| Refresh List      | Update backup records to ensure the latest backup data is displayed       |
| Delete Selected   | Delete the selected backup record (irreversible, please confirm before operation) |
| Restore Selected  | Restore the selected backup to the original file/directory location (will overwrite existing files, use with caution) |
| Open Location     | Jump to the actual storage directory of the selected backup to view backup file details |

#### 3.4 Other Functions
- **Interface Language Switch**: Click the **中文** or **English** button in the upper right corner of the window to switch the interface language in real time
- **Minimize to Tray**: When clicking the window close button, the program will not exit but minimize to the system tray and continue running
- **Tray Menu Operation**:
  1. Click the tray icon to display the operation menu
  2. Select **Show Window**: Restore the main program window
  3. Select **Enable Auto Backup**: Switch the auto backup status
  4. Select **Exit**: Close the program completely

### 7. Notes
1. **Path Restrictions**: The source path (files/directories to be backed up) and backup path cannot be the same or have an inclusion relationship (e.g., if the source path is `D:\Data`, the backup path cannot be `D:\Data\Backup`)
2. **Storage Check**: Ensure that the disk where the backup directory is located has sufficient storage space to avoid backup failure due to insufficient space
3. **Data Security**: Before performing the **Restore Selected** operation, it is recommended to confirm the status of the original file to avoid overwriting important data by mistake
4. **Hotkey Conflict**: If the set hotkey cannot be used, it may conflict with other program hotkeys. It is recommended to change the combination (e.g., `ctrl+alt+S`)
5. **Configuration File**: The program will automatically create a configuration file "backup_config.ini" to save settings, please do not delete or modify
6. **Log File**: The program will generate a log file "backup_tool.log" during operation to record operations and error information, which is convenient for troubleshooting

- Interface Simplification: Removed theme-related settings
- Comprehensive Performance Optimization: Improved backup algorithms, optimized operation flow, enhanced error handling and thread management

#### Version 1.2.0
- Interface beautification: Adopted Windows 11 style design, optimized overall visual experience
- Added duplicate backup file handling function, providing three handling methods
- Added custom backup file prefix function
- Extended backup file suffix customization methods, added sequential number suffix and custom suffix options
- Comprehensive optimization of performance, user experience and operation logic
- Enhanced error handling and thread management
- Improved configuration management function, supporting more configuration items
- Added custom theme color function

### 6. Feedback & Support
For any questions or suggestions, please submit an issue on the project page, and the developer will reply and handle it as soon as possible。
