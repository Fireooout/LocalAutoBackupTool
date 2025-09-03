# Local Auto Backup Tool  本地自动备份小工具

## 软件简介

本地自动备份小工具(LABT) 是一款功能强大的本地自动备份工具，旨在帮助用户轻松管理和保护重要数据。该软件提供了直观的图形用户界面，支持多种备份方式和灵活的配置选项，满足不同用户的备份需求。

### 主要功能

- **自动备份**：支持按设定的时间间隔自动备份文件和文件夹
- **手动备份**：提供一键手动备份功能，可随时进行即时备份
- **多源路径支持**：可同时添加多个文件和文件夹作为备份源
- **灵活的备份命名**：支持时间戳、序号和自定义后缀三种命名方式
- **备份管理**：提供备份列表查看、删除、还原和重命名功能
- **高级设置**：支持设置备份时间间隔、最大备份数量、快捷键等
- **自动编号模式**：支持自动和手动两种编号模式，确保备份文件有序管理
- **重复文件处理**：提供覆盖、重命名和跳过三种处理方式
- **完善的多语言支持**：支持中英文界面切换，所有UI元素完全依赖语言配置
- **系统托盘支持**：最小化到系统托盘，不占用任务栏空间
- **跳过隐藏文件**：可选择跳过隐藏文件和文件夹，提高备份灵活性
- **设置导入导出**：支持备份和恢复所有配置参数，便于配置迁移
- **增强托盘交互**：支持左键单击和双击显示窗口，提供更便捷的操作
- **优化的界面文本更新机制**：确保所有界面文本随语言切换而正确更新

### 系统要求

- Windows 操作系统
- Python 3.6 或更高版本
- 所需Python库：tkinter, os, sys, shutil, threading, datetime, logging, re, keyboard, pystray, PIL

### 使用方法

1. 运行程序后，在"备份设置"标签页中添加需要备份的文件或文件夹
2. 设置备份目标目录
3. 在"高级设置"标签页中配置备份选项，如时间间隔、最大备份数量等
4. 点击"开始备份"按钮启动自动备份，或使用快捷键进行手动备份
5. 在"备份管理"标签页中可以查看、删除、还原或重命名已有备份

## 更新日志

### 版本 1.5.0

#### 新增功能
- 完善了中英文翻译系统，修复了语言切换按钮和序号模式设置的翻译问题
- 增强了语言配置，添加了'chinese'和'english'翻译项，支持语言切换按钮文本的翻译
- 改进了界面文本更新机制，确保所有UI元素完全依赖语言配置
- 优化了多语言切换体验，提高了界面文本的一致性和完整性

#### 功能改进
- 修复了语言切换按钮的硬编码文本问题，改为使用语言配置获取翻译文本
- 改进了序号模式设置文本的条件判断逻辑，直接使用语言配置获取翻译文本
- 优化了界面元素的文本设置方法，提高了多语言支持的稳定性
- 增强了语言配置的完整性，确保所有界面文本都有对应的中英文翻译

#### 问题修复
- 修复了语言切换按钮创建代码中的语法错误
- 修复了序号模式标签、自动编号/手动编号单选按钮的条件判断文本设置问题
- 解决了部分UI元素文本不随语言切换而变化的问题
- 修复了中英文翻译不完整的问题

### 版本 1.4.0

#### 新增功能
- 增强了系统托盘图标功能，支持pystray和win32gui两种实现方式
- 添加了跳过隐藏文件/文件夹选项，提高备份灵活性
- 新增设置导入导出功能，支持备份和恢复所有配置参数
- 改进了托盘图标交互，支持左键单击显示窗口和双击显示窗口
- 优化了托盘图标菜单，提供更便捷的操作选项
- 添加了配置合并功能，导入设置时保留未导入的配置项
- 增强了窗口恢复机制，使用Windows API确保窗口正确显示
- 改进了托盘图标错误处理，提供更稳定的运行体验

#### 功能改进
- 优化了托盘图标创建和更新机制，提高图标显示稳定性
- 改进了备份配置文件格式，支持更多配置参数
- 优化了设置界面布局，提供更直观的用户体验
- 增强了错误处理和日志记录，便于问题诊断
- 改进了多语言支持，完善了界面文本翻译
- 优化了备份过程中的文件属性处理
- 改进了备份编号管理，提供更灵活的编号选项

#### 问题修复
- 修复了单击托盘图标无法显示窗口的问题
- 修复了某些情况下托盘图标无法正确创建的问题
- 修复了高DPI设置下界面元素显示异常的问题
- 修复了备份过程中文件属性处理的问题
- 修复了配置文件读取和写入的兼容性问题
- 修复了某些情况下窗口无法正确恢复的问题

### 版本 1.2.0

#### 新增功能
- 添加了备份序号模式设置，支持自动编号和手动编号两种模式
- 添加了设置起始序号功能，允许用户自定义备份编号的起始值
- 改进了备份编号重新分配机制，确保编号统一性
- 添加了备份快捷键和还原快捷键的独立设置
- 优化了备份列表显示，按时间排序显示最新备份
- 改进了备份还原功能，支持通过快捷键快速还原
- 添加了多语言支持，可在中英文界面之间切换
- 优化了DPI缩放支持，改善高分辨率显示器下的显示效果

#### 功能改进
- 优化了备份后缀生成逻辑，支持时间戳、序号和自定义后缀三种方式
- 改进了备份清理机制，按每项备份的最大数量清理旧备份
- 优化了备份文件/文件夹的识别和记录方式
- 改进了备份重命名功能，提供更友好的用户界面
- 优化了备份过程中的错误处理和日志记录
- 改进了备份文件夹结构，每个备份创建独立文件夹

#### 问题修复
- 修复了备份编号不连续的问题
- 修复了某些情况下备份失败后程序无响应的问题
- 修复了备份列表刷新不及时的问题
- 修复了高DPI设置下界面显示异常的问题
- 修复了备份还原过程中路径处理的问题

---

# Local Auto Backup Tool

## Software Introduction

Local Auto Backup Tool is a powerful local automatic backup tool designed to help users easily manage and protect important data. The software provides an intuitive graphical user interface, supports multiple backup methods and flexible configuration options to meet the backup needs of different users.

### Main Features

- **Automatic Backup**: Supports automatic backup of files and folders at set time intervals
- **Manual Backup**: Provides one-click manual backup function for instant backup at any time
- **Multi-source Path Support**: Can add multiple files and folders as backup sources simultaneously
- **Flexible Backup Naming**: Supports three naming methods: timestamp, number, and custom suffix
- **Backup Management**: Provides backup list viewing, deletion, restoration, and renaming functions
- **Advanced Settings**: Supports setting backup time intervals, maximum backup count, hotkeys, etc.
- **Auto Numbering Mode**: Supports both automatic and manual numbering modes to ensure orderly management of backup files
- **Duplicate File Handling**: Provides three handling methods: overwrite, rename, and skip
- **Enhanced Multi-language Support**: Supports switching between Chinese and English interfaces, all UI elements completely rely on language configuration
- **System Tray Support**: Minimizes to system tray without taking up taskbar space
- **Skip Hidden Files**: Option to skip hidden files and folders, increasing backup flexibility
- **Settings Import/Export**: Supports backup and restoration of all configuration parameters, facilitating configuration migration
- **Enhanced Tray Interaction**: Supports left-click and double-click to show window, providing more convenient operation
- **Optimized Interface Text Update Mechanism**: Ensures all interface texts update correctly with language switching

### System Requirements

- Windows Operating System
- Python 3.6 or higher
- Required Python libraries: tkinter, os, sys, shutil, threading, datetime, logging, re, keyboard, pystray, PIL

### Usage Instructions

1. After running the program, add files or folders to be backed up in the "Backup Settings" tab
2. Set the backup destination directory
3. Configure backup options in the "Advanced Settings" tab, such as time interval, maximum backup count, etc.
4. Click the "Start Backup" button to start automatic backup, or use hotkeys for manual backup
5. In the "Backup Management" tab, you can view, delete, restore, or rename existing backups

## Change Log

### Version 1.5.0

#### New Features
- Enhanced Chinese-English translation system, fixing translation issues for language switching buttons and numbering mode settings
- Improved language configuration, adding 'chinese' and 'english' translation items to support translation of language switching button text
- Enhanced interface text update mechanism, ensuring all UI elements completely rely on language configuration
- Optimized multi-language switching experience, improving consistency and completeness of interface text

#### Feature Improvements
- Fixed hardcoded text issues in language switching buttons, now using language configuration to get translated text
- Improved conditional judgment logic for numbering mode setting text, directly using language configuration to get translated text
- Optimized text setting methods for interface elements, improving stability of multi-language support
- Enhanced completeness of language configuration, ensuring all interface texts have corresponding Chinese and English translations

#### Bug Fixes
- Fixed syntax errors in language switching button creation code
- Fixed conditional judgment text setting issues for numbering mode labels, auto-numbering/manual numbering radio buttons
- Resolved issues where some UI element texts did not change with language switching
- Fixed incomplete Chinese-English translation issues

### Version 1.4.0

#### New Features
- Enhanced system tray icon functionality, supporting both pystray and win32gui implementations
- Added option to skip hidden files/folders, increasing backup flexibility
- Added settings import/export functionality, supporting backup and restoration of all configuration parameters
- Improved tray icon interaction, supporting left-click to show window and double-click to show window
- Optimized tray icon menu, providing more convenient operation options
- Added configuration merging functionality, preserving unimported configuration items when importing settings
- Enhanced window restoration mechanism, using Windows API to ensure proper window display
- Improved tray icon error handling, providing a more stable operating experience

#### Feature Improvements
- Optimized tray icon creation and update mechanism, improving icon display stability
- Improved backup configuration file format, supporting more configuration parameters
- Optimized settings interface layout, providing a more intuitive user experience
- Enhanced error handling and logging, facilitating problem diagnosis
- Improved multi-language support, completing interface text translation
- Optimized file attribute handling during backup process
- Improved backup number management, providing more flexible numbering options

#### Bug Fixes
- Fixed the issue where clicking the tray icon could not display the window
- Fixed the issue where the tray icon could not be created correctly in some cases
- Fixed the issue of abnormal interface element display under high DPI settings
- Fixed file attribute handling issues during backup process
- Fixed compatibility issues with configuration file reading and writing
- Fixed the issue where the window could not be restored correctly in some cases

### Version 1.2.0

#### New Features
- Added backup numbering mode settings, supporting both automatic and manual numbering modes
- Added set start number function, allowing users to customize the starting value of backup numbers
- Improved backup number reassignment mechanism to ensure numbering consistency
- Added independent settings for backup hotkey and restore hotkey
- Optimized backup list display, showing latest backups sorted by time
- Improved backup restoration function, supporting quick restoration via hotkeys
- Added multi-language support, allowing switching between Chinese and English interfaces
- Optimized DPI scaling support, improving display on high-resolution monitors

#### Feature Improvements
- Optimized backup suffix generation logic, supporting timestamp, number, and custom suffix methods
- Improved backup cleanup mechanism, cleaning old backups based on the maximum number per item
- Optimized identification and recording methods for backup files/folders
- Improved backup renaming function, providing a more user-friendly interface
- Optimized error handling and logging during the backup process
- Improved backup folder structure, creating independent folders for each backup

#### Bug Fixes
- Fixed the issue of non-consecutive backup numbers
- Fixed the issue of program unresponsiveness after backup failure in some cases
- Fixed the issue of untimely backup list refresh
- Fixed the issue of abnormal interface display under high DPI settings
- Fixed the issue of path handling during the backup restoration process

