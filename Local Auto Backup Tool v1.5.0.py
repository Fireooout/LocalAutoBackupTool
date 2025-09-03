import os
import sys
import shutil
import datetime
import logging
import threading
import configparser
import keyboard
import re
import stat
from pathlib import Path
from tkinter import ttk, filedialog, messagebox, Menu
import tkinter as tk
import win32gui
import win32con
import win32api
import win32file
import win32con


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


FONT_CONFIG = {

    'base_size': 12,

    'title_size': 14,

    'scale_factor': 1.3          

}


LANGUAGES = {

    'zh': {

        'app_title': '本地自动备份工具',

        'tab_backup': '自动备份',

        'tab_manage': '备份管理',

        'tab_settings': '高级设置',

        'frame_settings': '备份设置',

        'source_paths': '源路径:',

        'dest_dir': '备份目录:',

        'interval': '自动备份时间间隔(秒):',

        'max_backups': '每项最大备份数:',

        'hotkey': '备份快捷键:',

        'enable_hotkey': '启用备份快捷键',

        'browse': '浏览',

        'add_file': '添加文件',

        'add_dir': '添加目录',

        'remove': '移除',

        'open_location': '打开位置',

        'start_backup': '开始自动备份',

        'stop_backup': '停止自动备份',

        'manual_backup': '手动备份',

        'backup_time': '备份时间',

        'backup_name': '备份名称',

        'backup_path': '备份路径',

        'source_path': '源路径',

        'keep_backup': '永久保留',

        'refresh_list': '刷新列表',

        'delete_selected': '删除选中',

        'restore_selected': '还原选中',

        'open_backup': '打开备份位置',

        'rename_backup': '重命名备份',

        'appearance': '外观设置',

        'backup_options': '备份选项',

        'backup_prefix': '备份前缀:',

        'prefix_disabled': '禁用',

        'prefix_number': '序号',

        'prefix_custom': '自定义',

        'custom_prefix': '自定义前缀:',

        'suffix_type': '后缀类型:',

        'suffix_timestamp': '时间戳',

        'suffix_number': '序号',

        'suffix_custom': '自定义',

        'custom_suffix': '自定义后缀:',

        'handle_duplicates': '重名文件处理:',

        'duplicates_overwrite': '覆盖',

        'duplicates_rename': '重命名',

        'duplicates_skip': '跳过',

        'backup_count': '当前备份序号: {count}',

        'set_start_number': '设置起始序号',

        'save_settings': '保存并应用设置',

        'status_ready': '就绪',

        'backup_started': '自动备份已启动，间隔 {interval} 秒',

        'backup_stopped': '自动备份已停止',

        'backup_success': '备份成功完成 {timestamp}',

        'manual_backup_running': '正在执行手动备份...',

        'manual_backup_done': '手动备份完成',

        'hotkey_enabled': '备份快捷键已启用: {hotkey}',

        'hotkey_disabled': '备份快捷键已禁用',

        'restore_hotkey': '还原快捷键:',

        'enable_restore_hotkey': '启用还原快捷键',

        'restore_hotkey_enabled': '还原快捷键已启用: {hotkey}',

        'restore_hotkey_disabled': '还原快捷键已禁用',

        'error_no_source': '请添加源路径',

        'error_no_dest': '请选择备份目录',

        'error_invalid_number': '请输入有效的数字',

        'error_backup': '备份过程中出错: {error}',

        'error_hotkey_reg': '备份快捷键注册失败: {error}',

        'error_invalid_hotkey': '无效的备份快捷键格式',

        'error_path_contains': '源路径和备份目录不能相互包含',

        'select_backup': '请选择一个备份',

        'select_one_backup': '请选择一个备份进行操作',

        'error_backup_content': '备份内容无效或不存在',

        'confirm_restore': '确定要还原 {name} 吗？',

        'restore_success': '备份还原成功',

        'error_restore': '还原备份时出错: {error}',

        'error_delete': '删除备份时出错: {error}',

        'deleted_backup': '已删除备份: {name}',

        'error_no_source_path': '找不到源路径 {name}',

        'restored_backup': '已还原备份: {name}',

        'new_backup_name': '新备份名称:',

        'rename_success': '备份重命名成功',

        'error_rename': '重命名备份时出错: {error}',

        'error_prefix_empty': '自定义前缀不能为空',

        'error_suffix_empty': '自定义后缀不能为空',

        'source_path_removed': '以下源路径不存在，已从配置中移除:\n{paths}',

        'confirm_overwrite': '文件已存在，是否覆盖？',

        'confirm_set_start_number': '请输入新的起始序号:',

        'number_mode': '序号模式：',

        'auto_number': '自动编号',

        'manual_number': '手动编号',

        'tray_show': '显示窗口',

        'tray_minimize': '最小化到托盘',

        'tray_enable_auto': '启用自动备份',

        'tray_disable_auto': '禁用自动备份',

        'tray_exit': '退出',

        'ok_button': '确定',

        'cancel_button': '取消',

        'start_number_updated': '起始序号已更新',

        'error_set_start_number': '设置起始序号时出错: {error}',

        'skip_hidden': '备份时跳过隐藏文件/文件夹',

        'save_settings': '保存并应用设置',

        'restore_defaults': '恢复默认设置',

        'export_settings': '导出设置',

        'import_settings': '导入设置',

        'error': '错误',

        'file_not_found': '文件不存在',

        'invalid_config_format': '无效的配置文件格式',

        'config_merge_failed': '配置合并失败',

        'info': '提示',

        'settings_imported_success': '设置已成功导入并应用',

        'import_settings_failed': '导入设置失败',

        'settings_exported_success': '设置已成功导出到',

        'export_settings_failed': '导出设置失败',

        'confirm': '确认',

        'confirm_restore_defaults': '确定要恢复所有设置为默认值吗？此操作不可撤销。',

        'settings_restored_success': '已恢复所有设置为默认值。',

        'restore_defaults_failed': '恢复默认设置失败',

        'skipping_hidden_file': '跳过隐藏文件/文件夹: {path}',

        'error_checking_file_attributes': '检查文件属性时出错: {error}',

        'error_copying_file_attributes': '复制文件属性时出错: {error}',

        'renaming_backup': '重命名备份: {old_name} -> {new_name}',

        'rename_backup_failed': '重命名备份失败: {error}',

        'error_save_settings': '保存设置时出错',

        'error_save_advanced_settings': '保存高级设置时出错: {error}',

        'error_export_settings': '导出设置时出错: {error}',

        'error_import_settings': '导入设置时出错: {error}',

        'error_register_restore_hotkey': '注册还原快捷键失败: {error}',

        'error_during_backup': '备份过程中出错: {error}',

        'error_restore_defaults': '恢复默认设置时出错: {error}',

        'error_enable_hotkeys': '启用快捷键时出错: {error}',

        'error_reassign_backup_numbers': '重新分配备份编号时出错: {error}',

        'error_cleanup_old_backups': '清理旧备份时出错: {error}',

        'error_cleanup_old_backup_item': '按项清理旧备份时出错: {error}',

        'error_update_backup_list': '更新备份列表时出错: {error}',

        'error_get_auto_number_suffix': '获取自动编号后缀时出错: {error}',

        'error_update_source_path_file': '更新.source_path文件时出错: {error}',

        'error_exit_app': '退出应用程序时出错: {error}',

        'error_load_config': '加载配置文件失败: {error}',

        'error_save_config': '保存配置文件失败: {error}',

        'error_no_backup_dir': '备份目录不存在: {path}',

        'error_create_backup_dir': '创建备份目录失败: {error}',

        'error_copy_file': '复制文件失败: {error}',

        'error_copy_dir': '复制目录失败: {error}',

        'error_read_file': '读取文件失败: {error}',

        'error_write_file': '写入文件失败: {error}',

        'error_create_dir': '创建目录失败: {error}',

        'error_delete_file': '删除文件失败: {error}',

        'error_delete_dir': '删除目录失败: {error}',

        'error_rename_file': '重命名文件失败: {error}',

        'error_rename_dir': '重命名目录失败: {error}',

        'error_move_file': '移动文件失败: {error}',

        'error_move_dir': '移动目录失败: {error}',

        'error_list_dir': '列出目录内容失败: {error}',

        'error_get_file_info': '获取文件信息失败: {error}',

        'error_set_file_info': '设置文件信息失败: {error}',

        'error_get_file_attr': '获取文件属性失败: {error}',

        'error_set_file_attr': '设置文件属性失败: {error}',

        'error_get_file_time': '获取文件时间失败: {error}',

        'error_set_file_time': '设置文件时间失败: {error}',

        'error_get_file_size': '获取文件大小失败: {error}',

        'error_get_file_perm': '获取文件权限失败: {error}',

        'error_set_file_perm': '设置文件权限失败: {error}',

        'error_get_file_owner': '获取文件所有者失败: {error}',

        'error_set_file_owner': '设置文件所有者失败: {error}',

        'error_get_file_group': '获取文件组失败: {error}',

        'error_set_file_group': '设置文件组失败: {error}',

        'error_get_file_acl': '获取文件ACL失败: {error}',

        'error_set_file_acl': '设置文件ACL失败: {error}',

        'error_get_file_attr_ext': '获取文件扩展属性失败: {error}',

        'error_set_file_attr_ext': '设置文件扩展属性失败: {error}',

        'error_get_file_stream': '获取文件流失败: {error}',

        'error_set_file_stream': '设置文件流失败: {error}',

        'error_get_file_hardlink': '获取文件硬链接失败: {error}',

        'error_set_file_hardlink': '设置文件硬链接失败: {error}',

        'error_get_file_symlink': '获取文件符号链接失败: {error}',

        'error_set_file_symlink': '设置文件符号链接失败: {error}',

        'error_get_file_junction': '获取文件连接点失败: {error}',

        'error_set_file_junction': '设置文件连接点失败: {error}',

        'error_get_file_mount': '获取文件挂载点失败: {error}',

        'error_set_file_mount': '设置文件挂载点失败: {error}',

        'error_get_file_reparse': '获取文件重解析点失败: {error}',

        'error_set_file_reparse': '设置文件重解析点失败: {error}',

        'error_get_file_sparse': '获取文件稀疏属性失败: {error}',

        'error_set_file_sparse': '设置文件稀疏属性失败: {error}',

        'error_get_file_compressed': '获取文件压缩属性失败: {error}',

        'error_set_file_compressed': '设置文件压缩属性失败: {error}',

        'error_get_file_encrypted': '获取文件加密属性失败: {error}',

        'error_set_file_encrypted': '设置文件加密属性失败: {error}',

        'error_get_file_offline': '获取文件脱机属性失败: {error}',

        'error_set_file_offline': '设置文件脱机属性失败: {error}',

        'error_get_file_temporary': '获取文件临时属性失败: {error}',

        'error_set_file_temporary': '设置文件临时属性失败: {error}',

        'error_get_file_archive': '获取文件存档属性失败: {error}',

        'error_set_file_archive': '设置文件存档属性失败: {error}',

        'error_get_file_system': '获取文件系统属性失败: {error}',

        'error_set_file_system': '设置文件系统属性失败: {error}',

        'error_get_file_hidden': '获取文件隐藏属性失败: {error}',

        'error_set_file_hidden': '设置文件隐藏属性失败: {error}',

        'error_get_file_readonly': '获取文件只读属性失败: {error}',

        'error_set_file_readonly': '设置文件只读属性失败: {error}',

        'error_get_file_directory': '获取文件目录属性失败: {error}',

        'error_set_file_directory': '设置文件目录属性失败: {error}',

        'error_get_file_normal': '获取文件普通属性失败: {error}',

        'error_set_file_normal': '设置文件普通属性失败: {error}',

        'error_get_file_device': '获取文件设备属性失败: {error}',

        'error_set_file_device': '设置文件设备属性失败: {error}',

        'error_get_file_reserved': '获取文件保留属性失败: {error}',

        'error_set_file_reserved': '设置文件保留属性失败: {error}',

        'error_get_file_efile': '获取文件EFILE属性失败: {error}',

        'error_set_file_efile': '设置文件EFILE属性失败: {error}',

        'error_get_file_open': '获取文件打开属性失败: {error}',

        'error_set_file_open': '设置文件打开属性失败: {error}',

        'error_get_file_content_indexed': '获取文件内容索引属性失败: {error}',

        'error_set_file_content_indexed': '设置文件内容索引属性失败: {error}',

        'error_get_file_integrity_stream': '获取文件完整性流属性失败: {error}',

        'error_set_file_integrity_stream': '设置文件完整性流属性失败: {error}',

        'error_get_file_no_scrub_data': '获取文件无擦除数据属性失败: {error}',

        'error_set_file_no_scrub_data': '设置文件无擦除数据属性失败: {error}',

        'error_get_file_pinned': '获取文件固定属性失败: {error}',

        'error_set_file_pinned': '设置文件固定属性失败: {error}',

        'error_get_file_unpinned': '获取文件未固定属性失败: {error}',

        'error_set_file_unpinned': '设置文件未固定属性失败: {error}',

        'error_get_file_recall_on_open': '获取文件打开时召回属性失败: {error}',

        'error_set_file_recall_on_open': '设置文件打开时召回属性失败: {error}',

        'error_get_file_recall_on_data_access': '获取文件数据访问时召回属性失败: {error}',

        'error_set_file_recall_on_data_access': '设置文件数据访问时召回属性失败: {error}',

        'error_get_file_sticky': '获取文件粘性属性失败: {error}',

        'error_set_file_sticky': '设置文件粘性属性失败: {error}',

        'error_get_file_reparse_point': '获取文件重解析点属性失败: {error}',

        'error_set_file_reparse_point': '设置文件重解析点属性失败: {error}',

        'error_get_file_sparse_file': '获取文件稀疏文件属性失败: {error}',

        'error_set_file_sparse_file': '设置文件稀疏文件属性失败: {error}',

        'error_get_file_reparse_point': '获取文件重解析点属性失败: {error}',

        'error_set_file_reparse_point': '设置文件重解析点属性失败: {error}',

        'error_get_file_compressed': '获取文件压缩属性失败: {error}',

        'error_set_file_compressed': '设置文件压缩属性失败: {error}',

        'error_get_file_offline': '获取文件脱机属性失败: {error}',

        'error_set_file_offline': '设置文件脱机属性失败: {error}',

        'error_get_file_not_content_indexed': '获取文件非内容索引属性失败: {error}',

        'error_set_file_not_content_indexed': '设置文件非内容索引属性失败: {error}',

        'error_get_file_encrypted': '获取文件加密属性失败: {error}',

        'error_set_file_encrypted': '设置文件加密属性失败: {error}',

        'chinese': '中文',

        'english': '英文'

    },

    'en': {

        'app_title': 'Local Auto Backup Assistant',

        'tab_backup': 'Auto Backup',

        'tab_manage': 'Backup Management',

        'tab_settings': 'Advanced Settings',

        'frame_settings': 'Backup Settings',

        'source_paths': 'Source Paths:',

        'dest_dir': 'Backup Directory:',

        'interval': 'AutoBackup Interval (seconds):',

        'max_backups': 'Single Item Max Backups:',

        'hotkey': 'Backup Hotkey:',

        'enable_hotkey': 'Enable Backup Hotkey',

        'browse': 'Browse',

        'add_file': 'Add File',

        'add_dir': 'Add Directory',

        'remove': 'Remove',

        'open_location': 'Open Location',

        'start_backup': 'Start Auto Backup',

        'stop_backup': 'Stop Auto Backup',

        'manual_backup': 'Manual Backup',

        'backup_time': 'Backup Time',

        'backup_name': 'Backup Name',

        'backup_path': 'Backup Path',

        'source_path': 'Source Path',

        'keep_backup': 'Permanent Keep',

        'refresh_list': 'Refresh List',

        'delete_selected': 'Delete Selected',

        'restore_selected': 'Restore Selected',

        'open_backup': 'Open Backup Location',

        'rename_backup': 'Rename Backup',

        'appearance': 'Appearance',

        'backup_options': 'Backup Options',

        'backup_prefix': 'Backup Prefix:',

        'prefix_disabled': 'Disabled',

        'prefix_number': 'Number',

        'prefix_custom': 'Custom',

        'custom_prefix': 'Custom Prefix:',

        'suffix_type': 'Suffix Type:',

        'suffix_timestamp': 'Timestamp',

        'suffix_number': 'Number',

        'suffix_custom': 'Custom',

        'custom_suffix': 'Custom Suffix:',

        'handle_duplicates': 'Handle Duplicates:',

        'duplicates_overwrite': 'Overwrite',

        'duplicates_rename': 'Rename',

        'duplicates_skip': 'Skip',

        'backup_count': 'Current Backup Number: {count}',

        'set_start_number': 'Set Start Number',

        'save_settings': 'Save and Apply Settings',

        'status_ready': 'Ready',

        'backup_started': 'Auto backup started, interval {interval} seconds',

        'backup_stopped': 'Auto backup stopped',

        'backup_success': 'Backup successfully completed {timestamp}',

        'manual_backup_running': 'Performing manual backup...',

        'manual_backup_done': 'Manual backup completed',

        'hotkey_enabled': 'Backup hotkey enabled: {hotkey}',

        'hotkey_disabled': 'Backup hotkey disabled',

        'restore_hotkey': 'Restore Hotkey:',

        'enable_restore_hotkey': 'Enable Restore Hotkey',

        'restore_hotkey_enabled': 'Restore hotkey enabled: {hotkey}',

        'restore_hotkey_disabled': 'Restore hotkey disabled',

        'error_no_source': 'Please add source paths',

        'error_no_dest': 'Please select backup directory',

        'error_invalid_number': 'Please enter a valid number',

        'error_backup': 'Error during backup: {error}',

        'error_hotkey_reg': 'Failed to register Backup Hotkey: {error}',

        'error_invalid_hotkey': 'Invalid Backup Hotkey format',

        'error_path_contains': 'Source paths and backup directory cannot contain each other',

        'select_backup': 'Please select a backup',

        'select_one_backup': 'Please select one backup for operation',

        'error_backup_content': 'Invalid or non-existent backup content',

        'confirm_restore': 'Are you sure you want to restore {name}?',

        'restore_success': 'Backup restored successfully',

        'error_restore': 'Error restoring backup: {error}',

        'error_delete': 'Error deleting backup: {error}',

        'deleted_backup': 'Deleted backup: {name}',

        'error_no_source_path': 'Source path not found {name}',

        'restored_backup': 'Restored backup: {name}',

        'new_backup_name': 'New Backup Name:',

        'rename_success': 'Backup renamed successfully',

        'error_rename': 'Error renaming backup: {error}',

        'error_prefix_empty': 'Custom prefix cannot be empty',

        'error_suffix_empty': 'Custom suffix cannot be empty',

        'source_path_removed': 'The following source paths do not exist and have been removed from configuration:\n{paths}',

        'confirm_overwrite': 'File already exists, overwrite?',

        'confirm_set_start_number': 'Please enter new start number:',

        'number_mode': 'Number Mode:',

        'auto_number': 'Auto Numbering',

        'manual_number': 'Manual Numbering',

        'tray_show': 'Show Window',

        'tray_minimize': 'Minimize to Tray',

        'tray_enable_auto': 'Enable Auto Backup',

        'tray_disable_auto': 'Disable Auto Backup',

        'tray_exit': 'Exit',

        'ok_button': 'OK',

        'cancel_button': 'Cancel',

        'start_number_updated': 'Start number updated',

        'error_set_start_number': 'Error setting start number: {error}',

        'skip_hidden': 'Skip hidden files/folders when backing up',

        'save_settings': 'Save and Apply Settings',

        'restore_defaults': 'Restore Defaults',

        'export_settings': 'Export Settings',

        'import_settings': 'Import Settings',

        'error': 'Error',

        'file_not_found': 'File not found',

        'invalid_config_format': 'Invalid configuration file format',

        'config_merge_failed': 'Configuration merge failed',

        'info': 'Info',

        'settings_imported_success': 'Settings successfully imported and applied',

        'import_settings_failed': 'Failed to import settings',

        'settings_exported_success': 'Settings successfully exported to',

        'export_settings_failed': 'Failed to export settings',

        'confirm': 'Confirm',

        'confirm_restore_defaults': 'Are you sure you want to restore all settings to defaults? This action cannot be undone.',

        'settings_restored_success': 'All settings have been restored to defaults.',

        'restore_defaults_failed': 'Failed to restore defaults',

        'skipping_hidden_file': 'Skipping hidden file/folder: {path}',

        'error_checking_file_attributes': 'Error checking file attributes: {error}',

        'error_copying_file_attributes': 'Error copying file attributes: {error}',

        'renaming_backup': 'Renaming backup: {old_name} -> {new_name}',

        'rename_backup_failed': 'Failed to rename backup: {error}',

        'error_save_settings': 'Error saving settings',

        'error_save_advanced_settings': 'Error saving advanced settings: {error}',

        'error_export_settings': 'Error exporting settings: {error}',

        'error_import_settings': 'Error importing settings: {error}',

        'error_register_restore_hotkey': 'Failed to register restore hotkey: {error}',

        'error_during_backup': 'Error during backup: {error}',

        'error_restore_defaults': 'Error restoring defaults: {error}',

        'error_enable_hotkeys': 'Error enabling hotkeys: {error}',

        'error_reassign_backup_numbers': 'Error reassigning backup numbers: {error}',

        'error_cleanup_old_backups': 'Error cleaning up old backups: {error}',

        'error_cleanup_old_backup_item': 'Error cleaning up old backup item: {error}',

        'error_update_backup_list': 'Error updating backup list: {error}',

        'error_get_auto_number_suffix': 'Error getting auto number suffix: {error}',

        'error_update_source_path_file': 'Error updating .source_path file: {error}',

        'error_exit_app': 'Error exiting application: {error}',

        'error_load_config': 'Error loading configuration file: {error}',

        'error_save_config': 'Error saving configuration file: {error}',

        'error_no_backup_dir': 'Backup directory does not exist: {path}',

        'error_create_backup_dir': 'Failed to create backup directory: {error}',

        'error_copy_file': 'Failed to copy file: {error}',

        'error_copy_dir': 'Failed to copy directory: {error}',

        'error_read_file': 'Failed to read file: {error}',

        'error_write_file': 'Failed to write file: {error}',

        'error_create_dir': 'Failed to create directory: {error}',

        'error_delete_file': 'Failed to delete file: {error}',

        'error_delete_dir': 'Failed to delete directory: {error}',

        'error_rename_file': 'Failed to rename file: {error}',

        'error_rename_dir': 'Failed to rename directory: {error}',

        'error_move_file': 'Failed to move file: {error}',

        'error_move_dir': 'Failed to move directory: {error}',

        'error_list_dir': 'Failed to list directory contents: {error}',

        'error_get_file_info': 'Failed to get file information: {error}',

        'error_set_file_info': 'Failed to set file information: {error}',

        'error_get_file_attr': 'Failed to get file attributes: {error}',

        'error_set_file_attr': 'Failed to set file attributes: {error}',

        'error_get_file_time': 'Failed to get file time: {error}',

        'error_set_file_time': 'Failed to set file time: {error}',

        'error_get_file_size': 'Failed to get file size: {error}',

        'error_get_file_perm': 'Failed to get file permissions: {error}',

        'error_set_file_perm': 'Failed to set file permissions: {error}',

        'error_get_file_owner': 'Failed to get file owner: {error}',

        'error_set_file_owner': 'Failed to set file owner: {error}',

        'error_get_file_group': 'Failed to get file group: {error}',

        'error_set_file_group': 'Failed to set file group: {error}',

        'error_get_file_acl': 'Failed to get file ACL: {error}',

        'error_set_file_acl': 'Failed to set file ACL: {error}',

        'error_get_file_attr_ext': 'Failed to get file extended attributes: {error}',

        'error_set_file_attr_ext': 'Failed to set file extended attributes: {error}',

        'error_get_file_stream': 'Failed to get file stream: {error}',

        'error_set_file_stream': 'Failed to set file stream: {error}',

        'error_get_file_hardlink': 'Failed to get file hard link: {error}',

        'error_set_file_hardlink': 'Failed to set file hard link: {error}',

        'error_get_file_symlink': 'Failed to get file symbolic link: {error}',

        'error_set_file_symlink': 'Failed to set file symbolic link: {error}',

        'error_get_file_junction': 'Failed to get file junction: {error}',

        'error_set_file_junction': 'Failed to set file junction: {error}',

        'error_get_file_mount': 'Failed to get file mount point: {error}',

        'error_set_file_mount': 'Failed to set file mount point: {error}',

        'error_get_file_reparse': 'Failed to get file reparse point: {error}',

        'error_set_file_reparse': 'Failed to set file reparse point: {error}',

        'error_get_file_sparse': 'Failed to get file sparse attribute: {error}',

        'error_set_file_sparse': 'Failed to set file sparse attribute: {error}',

        'error_get_file_compressed': 'Failed to get file compressed attribute: {error}',

        'error_set_file_compressed': 'Failed to set file compressed attribute: {error}',

        'error_get_file_encrypted': 'Failed to get file encrypted attribute: {error}',

        'error_set_file_encrypted': 'Failed to set file encrypted attribute: {error}',

        'error_get_file_offline': 'Failed to get file offline attribute: {error}',

        'error_set_file_offline': 'Failed to set file offline attribute: {error}',

        'error_get_file_temporary': 'Failed to get file temporary attribute: {error}',

        'error_set_file_temporary': 'Failed to set file temporary attribute: {error}',

        'error_get_file_archive': 'Failed to get file archive attribute: {error}',

        'error_set_file_archive': 'Failed to set file archive attribute: {error}',

        'error_get_file_system': 'Failed to get file system attribute: {error}',

        'error_set_file_system': 'Failed to set file system attribute: {error}',

        'error_get_file_hidden': 'Failed to get file hidden attribute: {error}',

        'error_set_file_hidden': 'Failed to set file hidden attribute: {error}',

        'error_get_file_readonly': 'Failed to get file read-only attribute: {error}',

        'error_set_file_readonly': 'Failed to set file read-only attribute: {error}',

        'error_get_file_directory': 'Failed to get file directory attribute: {error}',

        'error_set_file_directory': 'Failed to set file directory attribute: {error}',

        'error_get_file_normal': 'Failed to get file normal attribute: {error}',

        'error_set_file_normal': 'Failed to set file normal attribute: {error}',

        'error_get_file_device': 'Failed to get file device attribute: {error}',

        'error_set_file_device': 'Failed to set file device attribute: {error}',

        'error_get_file_reserved': 'Failed to get file reserved attribute: {error}',

        'error_set_file_reserved': 'Failed to set file reserved attribute: {error}',

        'error_get_file_efile': 'Failed to get file EFILE attribute: {error}',

        'error_set_file_efile': 'Failed to set file EFILE attribute: {error}',

        'error_get_file_open': 'Failed to get file open attribute: {error}',

        'error_set_file_open': 'Failed to set file open attribute: {error}',

        'error_get_file_content_indexed': 'Failed to get file content indexed attribute: {error}',

        'error_set_file_content_indexed': 'Failed to set file content indexed attribute: {error}',

        'error_get_file_integrity_stream': 'Failed to get file integrity stream attribute: {error}',

        'error_set_file_integrity_stream': 'Failed to set file integrity stream attribute: {error}',

        'error_get_file_no_scrub_data': 'Failed to get file no scrub data attribute: {error}',

        'error_set_file_no_scrub_data': 'Failed to set file no scrub data attribute: {error}',

        'error_get_file_pinned': 'Failed to get file pinned attribute: {error}',

        'error_set_file_pinned': 'Failed to set file pinned attribute: {error}',

        'error_get_file_unpinned': 'Failed to get file unpinned attribute: {error}',

        'error_set_file_unpinned': 'Failed to set file unpinned attribute: {error}',

        'error_get_file_recall_on_open': 'Failed to get file recall on open attribute: {error}',

        'error_set_file_recall_on_open': 'Failed to set file recall on open attribute: {error}',

        'error_get_file_recall_on_data_access': 'Failed to get file recall on data access attribute: {error}',

        'error_set_file_recall_on_data_access': 'Failed to set file recall on data access attribute: {error}',

        'error_get_file_sticky': 'Failed to get file sticky attribute: {error}',

        'error_set_file_sticky': 'Failed to set file sticky attribute: {error}',

        'error_get_file_reparse_point': 'Failed to get file reparse point attribute: {error}',

        'error_set_file_reparse_point': 'Failed to set file reparse point attribute: {error}',

        'error_get_file_sparse_file': 'Failed to get file sparse file attribute: {error}',

        'error_set_file_sparse_file': 'Failed to set file sparse file attribute: {error}',

        'error_get_file_compressed': 'Failed to get file compressed attribute: {error}',

        'error_set_file_compressed': 'Failed to set file compressed attribute: {error}',

        'error_get_file_offline': 'Failed to get file offline attribute: {error}',

        'error_set_file_offline': 'Failed to set file offline attribute: {error}',

        'error_get_file_not_content_indexed': 'Failed to get file not content indexed attribute: {error}',

        'error_set_file_not_content_indexed': 'Failed to set file not content indexed attribute: {error}',

        'error_get_file_encrypted': 'failed to get file encrypted attribute: {error}',

        'error_set_file_encrypted': 'failed to set file encrypted attribute: {error}',

        'chinese': 'chinese',

        'english': 'english'

    }

}


def get_resource_path(relative_path):
    

    try:
                                            

        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    

    return os.path.join(base_path, relative_path)


class AutoBackupTool:
    def __init__(self, root):
        self.root = root
        self.root.title(LANGUAGES['zh']['app_title'])
        self.root.geometry("900x750")
        

        self.root.resizable(True, True)            
        

        self.set_dpi_awareness()
        

        self.root.title(LANGUAGES['zh']['app_title'])
        

        self.apply_font_config()
        

        self.config_file = "backup_config.ini"
        

        self.source_paths = []
        self.dest_dir = ""
        self.interval = 300
        self.max_backups = 5
        self.hotkey = "ctrl+F1"
        self.hotkey_var = tk.BooleanVar(value=False)
        self.hotkey_registered = False
        self.running = False
        self.stop_event = threading.Event()
        self.backup_thread = None
        

        self.current_lang = 'zh'
        self.lang_strings = LANGUAGES[self.current_lang]
        self.suffix_type = "number"                             
        self.custom_suffix = ""
        self.duplicate_handling = "rename"                           
        self.backup_counter = 1
        self.start_number = 1
        self.auto_number_mode = True               
        

        self.restore_hotkey = "ctrl+F2"
        self.restore_hotkey_var = tk.BooleanVar(value=False)
        self.restore_hotkey_registered = False
        

        self.skip_hidden = False
        self.skip_hidden_var = tk.BooleanVar(value=False)
        

        self.normal_icon = get_resource_path("folder-sync.ico")
        self.active_icon = get_resource_path("folder-sync-start.ico")              
        

        self.load_config()
        

        self.removed_paths = []
        for source_path in self.source_paths[:]:            
            if not os.path.exists(source_path):
                self.source_paths.remove(source_path)
                self.removed_paths.append(source_path)
                logging.warning(f"源路径不存在，已移除: {source_path}")
        

        if self.removed_paths:
            self.save_config()            
        

        self.create_widgets()
        

        self.set_icon()
        

        self.root.after(100, self.create_tray_icon)
        

        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        

        self.update_ui_texts()
        

        self.root.after(300, self.update_backup_list)
        

        if self.removed_paths:
            removed_paths_str = "\n".join(self.removed_paths)
            self.root.after(500, lambda: messagebox.showwarning(

                "Warning", 

                self.lang_strings['source_path_removed'].format(paths=removed_paths_str)

            ))
    

    def set_dpi_awareness(self):
        

        try:
            import ctypes
                     

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
                       

            self.root.tk.call('tk', 'scaling', FONT_CONFIG['scale_factor'])
            

            title_font_size = int(FONT_CONFIG['title_size'] * FONT_CONFIG['scale_factor'])
            title_font = ('Microsoft YaHei', title_font_size, 'bold')
            self.root.title_font = title_font
        except Exception as e:
            logging.error(f"设置DPI感知失败: {e}")
    

    def apply_font_config(self):
        

        try:
                              

            base_font_size = int(FONT_CONFIG['base_size'] * FONT_CONFIG['scale_factor'])
            title_font_size = int(FONT_CONFIG['title_size'] * FONT_CONFIG['scale_factor'])
            

            default_font = ('Microsoft YaHei', base_font_size)
            title_font = ('Microsoft YaHei', title_font_size, 'bold')
                                 

            tab_font_size = int(FONT_CONFIG['title_size'] * FONT_CONFIG['scale_factor'])
            tab_font = ('Microsoft YaHei', tab_font_size, 'bold')
            

            style = ttk.Style()
            

            style.configure('TLabel', font=default_font)
            style.configure('Title.TLabel', font=title_font)
            

            style.configure('TButton', font=default_font)
            

            style.configure('TLabelframe.Label', font=title_font)
            

            entry_font_size = max(base_font_size, 16)            
            entry_font = ('Microsoft YaHei', entry_font_size)
            style.configure('TEntry', font=entry_font)
            

            combobox_font_size = max(base_font_size, 16)            
            combobox_font = ('Microsoft YaHei', combobox_font_size)
            style.configure('TCombobox', font=combobox_font)
            

            style.configure('TCheckbutton', font=default_font)
            

            style.configure('TRadiobutton', font=default_font)
            

            treeview_font_size = max(base_font_size - 1, 14)                
            treeview_font = ('Microsoft YaHei', treeview_font_size)
            treeview_heading_font = ('Microsoft YaHei', treeview_font_size, 'bold')
                            

            treeview_row_height = int(treeview_font_size * 2.2)                
            style.configure('Treeview', font=treeview_font, rowheight=treeview_row_height)
            style.configure('Treeview.Heading', font=treeview_heading_font)
            

            listbox_font_size = max(base_font_size - 1, 14)                 
            listbox_font = ('Microsoft YaHei', listbox_font_size)
            self.root.option_add('*Listbox.font', listbox_font)
            

            style.configure('TNotebook.Tab', font=tab_font)
            

            style.configure('TNotebook', borderwidth=2, relief='solid')
            style.configure('TNotebook.Tab', padding=[12, 8], background='#f0f0f0')
            style.map('TNotebook.Tab', 

                     background=[('selected', '#e1e1e1'), ('active', '#e8e8e8')],

                     foreground=[('selected', 'black'), ('active', 'black')])
            

            self.root.option_add('*Text.font', default_font)
            

            self.root.option_add('*Menu.font', default_font)
            

            self.root.option_add('*Dialog.msg.font', default_font)
            self.root.option_add('*Dialog*Button.font', default_font)
            

            self.root.option_add('*Dialog*Entry.font', entry_font)
            self.root.option_add('*Dialog*Label.font', default_font)
            

        except Exception as e:
            logging.error(f"应用字体配置失败: {e}")
    

    def set_language(self, lang):
        

        self.current_lang = lang
        self.lang_strings = LANGUAGES[self.current_lang]
        

        self.root.title(self.lang_strings['app_title'])
        if hasattr(self.root, 'title_font'):
                           

            try:
                                             

                import ctypes
                hwnd = self.root.winfo_id()
                if hwnd:
                            

                    font_name, font_size, font_weight = self.root.title_font
                                              

                    pass
            except Exception as e:
                logging.error(f"应用窗口标题字体失败: {e}")
        

        self.update_ui_texts()
        self.update_lang_button_state()
                

        self.create_tray_icon()
    

    def _apply_font_to_widget(self, widget, font):
        

        try:
            if hasattr(widget, 'configure'):
                widget.configure(font=font)
            

            if hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    self._apply_font_to_widget(child, font)
        except Exception as e:
            logging.error(f"应用字体到控件失败: {e}")
    

    def update_ui_texts(self):
        

        for i, tab_text in enumerate([self.lang_strings['tab_backup'], self.lang_strings['tab_manage'], self.lang_strings['tab_settings']]):
            self.tab_control.tab(i, text=tab_text)
        

        self.frame_settings.config(text=self.lang_strings['frame_settings'])
        self.source_frame.config(text=self.lang_strings['source_paths'])
        self.dest_frame.config(text=self.lang_strings['dest_dir'])
        self.hotkey_label.config(text=self.lang_strings['hotkey'])
        self.interval_label.config(text=self.lang_strings['interval'])
        self.max_backups_label.config(text=self.lang_strings['max_backups'])
        self.check_hotkey.config(text=self.lang_strings['enable_hotkey'])
        self.restore_hotkey_label.config(text=self.lang_strings['restore_hotkey'])
        self.check_restore_hotkey.config(text=self.lang_strings['enable_restore_hotkey'])
        self.btn_add_file.config(text=self.lang_strings['add_file'])
        self.btn_add_dir.config(text=self.lang_strings['add_dir'])
        self.btn_remove_source.config(text=self.lang_strings['remove'])
        self.btn_browse_dest.config(text=self.lang_strings['browse'])
        self.btn_open_source.config(text=self.lang_strings['open_location'])
        self.btn_open_dest.config(text=self.lang_strings['open_location'])
        

        if self.running:
            self.start_btn.config(text=self.lang_strings['stop_backup'])
        else:
            self.start_btn.config(text=self.lang_strings['start_backup'])
        

        self.btn_manual_backup.config(text=self.lang_strings['manual_backup'])
        

        self.backup_list.heading("date", text=self.lang_strings['backup_time'])
        self.backup_list.heading("name", text=self.lang_strings['backup_name'])
        self.backup_list.heading("keep", text=self.lang_strings['keep_backup'])
        self.btn_refresh.config(text=self.lang_strings['refresh_list'])
        self.btn_delete.config(text=self.lang_strings['delete_selected'])
        self.btn_restore.config(text=self.lang_strings['restore_selected'])
        self.btn_open_backup.config(text=self.lang_strings['open_backup'])
        self.btn_rename.config(text=self.lang_strings['rename_backup'])
        

        self.frame_backup_options.config(text=self.lang_strings['backup_options'])
        self.suffix_type_label.config(text=self.lang_strings['suffix_type'])
        self.suffix_timestamp_radio.config(text=self.lang_strings['suffix_timestamp'])
        self.suffix_number_radio.config(text=self.lang_strings['suffix_number'])
        self.suffix_custom_radio.config(text=self.lang_strings['suffix_custom'])
        self.custom_suffix_label.config(text=self.lang_strings['custom_suffix'])
        self.duplicates_label.config(text=self.lang_strings['handle_duplicates'])
        self.duplicates_overwrite_radio.config(text=self.lang_strings['duplicates_overwrite'])
        self.duplicates_rename_radio.config(text=self.lang_strings['duplicates_rename'])
        self.duplicates_skip_radio.config(text=self.lang_strings['duplicates_skip'])
        self.backup_count_label.config(text=self.lang_strings['backup_count'].format(count=self.backup_counter))
        self.btn_set_start_number.config(text=self.lang_strings['set_start_number'])
        

        if hasattr(self, 'number_mode_label'):
            self.number_mode_label.config(text=self.lang_strings['number_mode'])
        if hasattr(self, 'auto_number_radio'):
            self.auto_number_radio.config(text=self.lang_strings['auto_number'])
        if hasattr(self, 'manual_number_radio'):
            self.manual_number_radio.config(text=self.lang_strings['manual_number'])
            

        self.btn_save_settings.config(text=self.lang_strings['save_settings'] if 'save_settings' in self.lang_strings else 'Save Settings')
        

        if hasattr(self, 'check_skip_hidden'):
            self.check_skip_hidden.config(text=self.lang_strings['skip_hidden'])
        

        if hasattr(self, 'btn_restore_defaults'):
            self.btn_restore_defaults.config(text=self.lang_strings['restore_defaults'])
        if hasattr(self, 'btn_export_settings'):
            self.btn_export_settings.config(text=self.lang_strings['export_settings'])
        if hasattr(self, 'btn_import_settings'):
            self.btn_import_settings.config(text=self.lang_strings['import_settings'])
        

        self.btn_zh.config(text=self.lang_strings['chinese'])
        self.btn_en.config(text=self.lang_strings['english'])
        self.update_lang_button_state()
        

        if self.status_var.get() == LANGUAGES['zh']['status_ready'] or self.status_var.get() == LANGUAGES['en']['status_ready']:
            self.status_var.set(self.lang_strings['status_ready'])
    

    def check_path_relationship(self, path1, path2):
        

        if not path1 or not path2:
            return False
            

        path1 = os.path.abspath(os.path.normpath(path1))
        path2 = os.path.abspath(os.path.normpath(path2))
        

        if path1 == path2:
            return True
            

        path1_parts = Path(path1).parts
        path2_parts = Path(path2).parts
        

        min_length = min(len(path1_parts), len(path2_parts))
        for i in range(min_length):
            if path1_parts[i] != path2_parts[i]:
                return False
        

        return True
    

    def set_icon(self, active=False):
        try:
                        

            icon_path = self.active_icon if active else self.normal_icon
            

            if os.path.exists(icon_path):
                          

                logging.info(f"尝试设置图标，路径: {icon_path}")
                

                try:
                    import ctypes
                    import win32gui
                    import win32con
                    import win32api
                    from win32api import GetSystemMetrics
                    

                    hwnd = self.root.winfo_id()
                    if hwnd:
                                              

                        app_id = "Local.Auto.Backup.Assistant"
                        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
                        

                        hicon = None
                        

                        try:
                            hicon = win32gui.LoadImage(

                                None, 

                                icon_path, 

                                win32con.IMAGE_ICON, 

                                0, 

                                0, 

                                win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE

                            )
                            if hicon:
                                logging.info(f"使用LoadImage加载图标成功")
                        except Exception as e:
                            logging.warning(f"使用LoadImage加载图标失败: {e}")
                        

                        if not hicon:
                            try:
                                                        

                                hicon_large = ctypes.c_void_p()
                                hicon_small = ctypes.c_void_p()
                                result = ctypes.windll.shell32.ExtractIconExW(

                                    icon_path, 0, 

                                    ctypes.byref(hicon_large), 

                                    ctypes.byref(hicon_small), 1

                                )
                                if result > 0:
                                    hicon = hicon_large
                                    logging.info(f"使用ExtractIconEx加载图标成功")
                            except Exception as e:
                                logging.warning(f"使用ExtractIconEx加载图标失败: {e}")
                        

                        if not hicon:
                            try:
                                         

                                hinstance = win32gui.GetModuleHandle(None)
                                hicon = win32gui.LoadImage(

                                    hinstance, 

                                    icon_path, 

                                    win32con.IMAGE_ICON, 

                                    GetSystemMetrics(win32con.SM_CXICON), 

                                    GetSystemMetrics(win32con.SM_CYICON), 

                                    win32con.LR_LOADFROMFILE

                                )
                                if hicon:
                                    logging.info(f"使用LoadImage(指定大小)加载图标成功")
                            except Exception as e:
                                logging.warning(f"使用LoadImage(指定大小)加载图标失败: {e}")
                        

                        if hicon:
                                                       

                            ctypes.windll.user32.SendMessageW(hwnd, win32con.WM_SETICON, win32con.ICON_BIG, hicon)
                                                       

                            ctypes.windll.user32.SendMessageW(hwnd, win32con.WM_SETICON, win32con.ICON_SMALL, hicon)
                                     

                            ctypes.windll.user32.SetClassLongPtrW(hwnd, win32con.GCL_HICON, hicon)
                            ctypes.windll.user32.SetClassLongPtrW(hwnd, win32con.GCL_HICONSM, hicon)
                            

                            ctypes.windll.user32.InvalidateRect(hwnd, None, True)
                            ctypes.windll.user32.UpdateWindow(hwnd)
                            

                            ctypes.windll.user32.SendMessageW(hwnd, win32con.WM_SETTINGCHANGE, 0, 0)
                            

                            logging.info(f"成功设置窗口图标: {icon_path}")
                        else:
                            logging.warning(f"无法加载图标文件: {icon_path}")
                except Exception as e:
                    logging.warning(f"设置Windows任务栏图标失败: {e}")
                    

                try:
                    self.root.iconbitmap(icon_path)
                    logging.info(f"使用iconbitmap设置图标成功")
                except Exception as e:
                    logging.warning(f"使用iconbitmap设置图标失败: {e}")
            else:
                self.root.iconbitmap(default="")
                logging.warning(f"图标文件不存在: {icon_path}")
        except Exception as e:
            logging.error(f"设置图标失败: {e}")
    

    def create_tray_icon(self, active=False):
        

        try:
                                  

            try:
                import pystray
                from PIL import Image, ImageDraw
                

                if hasattr(self, 'tray_icon') and self.tray_icon:
                    try:
                        self.tray_icon.stop()
                        logging.info("已停止并清理现有托盘图标")
                    except Exception as e:
                        logging.warning(f"清理现有托盘图标时出错: {e}")
                

                app_title = self.lang_strings.get('app_title', 'Auto Backup Tool')
                tray_show_text = self.lang_strings.get('tray_show', 'Show Window')
                tray_disable_auto_text = self.lang_strings.get('tray_disable_auto', 'Disable Auto Backup')
                tray_enable_auto_text = self.lang_strings.get('tray_enable_auto', 'Enable Auto Backup')
                tray_exit_text = self.lang_strings.get('tray_exit', 'Exit')
                

                try:
                    icon_path = self.active_icon if active else self.normal_icon
                    if os.path.exists(icon_path):
                                          

                        file_ext = os.path.splitext(icon_path)[1].lower()
                        logging.info(f"尝试加载图标文件: {icon_path} (扩展名: {file_ext})")
                        

                        icon_image = Image.open(icon_path)
                        logging.info(f"成功加载图标文件: {icon_path}, 大小: {icon_image.size}, 模式: {icon_image.mode}")
                    else:
                                          

                        def create_image(color):
                            width, height = 16, 16
                            image = Image.new('RGB', (width, height), color)
                            draw = ImageDraw.Draw(image)
                                         

                            draw.rectangle([(3, 3), (13, 13)], fill='white')
                            draw.rectangle([(6, 6), (10, 10)], fill=color)
                            return image
                        

                        color = 'green' if active else 'gray'
                        icon_image = create_image(color)
                        logging.warning(f"图标文件不存在，使用默认图标: {icon_path}")
                except Exception as e:
                            

                    def create_image(color):
                        width, height = 16, 16
                        image = Image.new('RGB', (width, height), color)
                        draw = ImageDraw.Draw(image)
                                     

                        draw.rectangle([(3, 3), (13, 13)], fill='white')
                        draw.rectangle([(6, 6), (10, 10)], fill=color)
                        return image
                    

                    color = 'green' if active else 'gray'
                    icon_image = create_image(color)
                    logging.error(f"加载图标文件失败，使用备用图标: {e}")
                

                def on_double_click(icon, event):
                              

                    if event.button == pystray.MouseButton.LEFT:
                        logging.info("pystray托盘图标左键双击事件触发")
                                         

                        self.root.after(0, self.restore_from_tray)
                

                menu_items = [

                    pystray.MenuItem(tray_show_text, lambda: (logging.info("pystray托盘图标默认菜单项触发"), self.root.after(0, self.restore_from_tray)), default=True),

                    pystray.MenuItem(

                        tray_disable_auto_text if self.running else tray_enable_auto_text,

                        lambda: self.root.after(0, self.toggle_backup)

                    ),

                    pystray.Menu.SEPARATOR,

                    pystray.MenuItem(tray_exit_text, lambda: self.root.after(0, self.exit_application))

                ]
                

                self.tray_icon = pystray.Icon(app_title, icon_image, app_title)
                self.tray_icon.menu = pystray.Menu(*menu_items)
                            

                self.tray_icon.on_double_click = on_double_click              
                

                import threading
                def run_tray():
                    try:
                        self.tray_icon.run()
                    except Exception as e:
                        logging.error(f"托盘图标运行错误: {e}")
                

                self.tray_thread = threading.Thread(target=run_tray, daemon=True)
                self.tray_thread.start()
                logging.info("使用pystray创建托盘图标成功")
                

            except ImportError:
                                            

                logging.warning("pystray库不可用，使用win32gui实现托盘图标")
                self._create_win32_tray_icon(active)
                

            except Exception as e:
                logging.error(f"创建pystray托盘图标失败: {e}")
                                  

                self._create_win32_tray_icon(active)
        except Exception as e:
            logging.error(f"创建/更新托盘图标失败: {e}")
    

    def _create_win32_tray_icon(self, active=False):
        

        try:
            import win32gui
            import win32con
            

            hicon = None
            try:
                          

                icon_path = self.active_icon if active else self.normal_icon
                if os.path.exists(icon_path):
                                      

                    file_ext = os.path.splitext(icon_path)[1].lower()
                    logging.info(f"尝试加载托盘图标文件: {icon_path} (扩展名: {file_ext})")
                    

                    hicon = win32gui.LoadImage(0, icon_path, win32con.IMAGE_ICON, 0, 0, 

                                             win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE)
                    

                    if hicon:
                        logging.info(f"成功加载托盘图标文件: {icon_path}")
                    else:
                        logging.warning(f"加载托盘图标文件失败，返回空句柄: {icon_path}")
                else:
                    logging.warning(f"托盘图标文件不存在: {icon_path}")
            except Exception as e:
                logging.warning(f"加载托盘图标文件失败: {e}")
            

            self.TRAY_ICON_ID = 1001
            self.TRAY_CALLBACK_MESSAGE = win32con.WM_USER + 1
            

            app_title = self.lang_strings.get('app_title', 'Auto Backup Tool')
            

            wc = win32gui.WNDCLASS()
            wc.lpszClassName = "AutoBackupTrayIcon"
            wc.lpfnWndProc = self._tray_icon_callback
            wc.hInstance = win32gui.GetModuleHandle(None)
            

            try:
                                  

                self._class_atom = win32gui.RegisterClass(wc)
            except Exception:
                            

                self._class_atom = win32gui.FindWindow(wc.lpszClassName, None)
                

            self.tray_window = win32gui.CreateWindow(

                self._class_atom, "Auto Backup Tray Window", 0, 0, 0, 1, 1, 0, 0, wc.hInstance, None

            )
            

            nid = (self.tray_window, self.TRAY_ICON_ID, win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,

                  self.TRAY_CALLBACK_MESSAGE, hicon, app_title)
            

            if hasattr(self, 'tray_icon_info'):
                win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, nid)
                logging.info("托盘图标更新成功")
            else:
                win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
                logging.info("托盘图标创建成功")
            

            self.tray_icon_info = nid
        except Exception as e:
            logging.error(f"创建win32托盘图标失败: {e}")
    

    def _create_default_icon(self, active):
        

        try:
            import win32gui
            import win32con
            

            hdc = win32gui.GetDC(0)
            hdcMem = win32gui.CreateCompatibleDC(hdc)
            hbm = win32gui.CreateCompatibleBitmap(hdc, 16, 16)
            hbmOld = win32gui.SelectObject(hdcMem, hbm)
            

            color = 0x00FF00 if active else 0xAAAAAA         
            brush = win32gui.CreateSolidBrush(color)
            win32gui.FillRect(hdcMem, (0, 0, 16, 16), brush)
            win32gui.DeleteObject(brush)
            

            import ctypes
            class ICONINFO(ctypes.Structure):
                _fields_ = [

                    ('fIcon', ctypes.c_bool),

                    ('xHotspot', ctypes.c_long),

                    ('yHotspot', ctypes.c_long),

                    ('hbmMask', ctypes.c_void_p),

                    ('hbmColor', ctypes.c_void_p)

                ]
            

            icon_info = ICONINFO()
            icon_info.fIcon = True
            icon_info.xHotspot = 0
            icon_info.yHotspot = 0
            icon_info.hbmMask = hbm
            icon_info.hbmColor = hbm
            

            hicon = win32gui.CreateIconIndirect(icon_info)
            

            win32gui.SelectObject(hdcMem, hbmOld)
            win32gui.DeleteObject(hbm)
            win32gui.DeleteDC(hdcMem)
            win32gui.ReleaseDC(0, hdc)
            

            return hicon
        except Exception as e:
            logging.error(f"创建默认图标失败: {e}")
            return None
    

    def _tray_icon_callback(self, hwnd, msg, wparam, lparam):
        

        import win32gui
        import win32con
        

        if msg == self.TRAY_CALLBACK_MESSAGE:
            if lparam == win32con.wm_lbuttonup:        
                logging.info("win32gui托盘图标左键单击事件触发")
                                 

                self.root.after(0, self.restore_from_tray)
            elif lparam == win32con.WM_LBUTTONDBLCLK:        
                logging.info("win32gui托盘图标左键双击事件触发")
                        

                self.root.after(0, self.restore_from_tray)
            elif lparam == win32con.WM_RBUTTONUP:
                logging.info("win32gui托盘图标右键单击事件触发")
                        

                self.root.after(0, lambda: self._show_tray_menu(hwnd))
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
    

    def _show_tray_menu(self, hwnd):
        

        try:
            import win32gui
            import win32con
            

            menu = win32gui.CreatePopupMenu()
            

            show_id = 1001
            toggle_id = 1002
            exit_id = 1003
            

            tray_show_text = self.lang_strings.get('tray_show', 'Show Window')
            tray_minimize_text = self.lang_strings.get('tray_minimize', 'Minimize to Tray')
            tray_disable_auto_text = self.lang_strings.get('tray_disable_auto', 'Disable Auto Backup')
            tray_enable_auto_text = self.lang_strings.get('tray_enable_auto', 'Enable Auto Backup')
            tray_exit_text = self.lang_strings.get('tray_exit', 'Exit')
            

            win32gui.AppendMenu(menu, win32con.MF_STRING, show_id, tray_show_text)
            

            if self.running:
                win32gui.AppendMenu(menu, win32con.MF_STRING, toggle_id, tray_disable_auto_text)
            else:
                win32gui.AppendMenu(menu, win32con.MF_STRING, toggle_id, tray_enable_auto_text)
            

            win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, None)
            win32gui.AppendMenu(menu, win32con.MF_STRING, exit_id, tray_exit_text)
            

            x, y = win32gui.GetCursorPos()
            win32gui.SetForegroundWindow(hwnd)
            

            cmd = win32gui.TrackPopupMenu(

                menu, win32con.TPM_LEFTALIGN | win32con.TPM_LEFTBUTTON | win32con.TPM_BOTTOMALIGN | win32con.TPM_RETURNCMD,

                x, y, 0, hwnd, None

            )
            

            if cmd == show_id:
                self.root.after(0, self.restore_from_tray)
            elif cmd == toggle_id:
                self.root.after(0, self.toggle_backup)
            elif cmd == exit_id:
                self.root.after(0, self.exit_application)
            

            win32gui.PostMessage(hwnd, win32con.WM_NULL, 0, 0)
            

        except Exception as e:
            logging.error(f"显示托盘菜单失败: {e}")
    

    def minimize_to_tray(self):
        

        self.root.withdraw()
    

    def restore_from_tray(self):
        

        try:
            logging.info("开始从托盘恢复窗口")
                    

            self.root.deiconify()
            logging.info("窗口deiconify调用完成")
                     

            self.root.lift()
            logging.info("窗口lift调用完成")
            

            try:
                import win32gui
                import win32con
                        

                hwnd = self.root.winfo_id()
                           

                win32gui.SetForegroundWindow(hwnd)
                logging.info("使用SetForegroundWindow强制窗口显示到前台")
                      

                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                logging.info("使用ShowWindow恢复窗口")
            except Exception as e:
                logging.warning(f"使用Windows API恢复窗口失败: {e}")
            

            self.root.attributes('-topmost', True)
            logging.info("设置topmost为True完成")
            self.root.after_idle(self.root.attributes, '-topmost', False)
            logging.info("设置topmost为False的延迟任务已添加")
            self.root.focus_set()
            logging.info("窗口focus_set调用完成")
                    

            self.root.update_idletasks()
            logging.info("窗口update_idletasks调用完成")
            logging.info("窗口从托盘恢复成功")
        except Exception as e:
            logging.error(f"恢复窗口失败: {e}")
    

    def exit_application(self):
        

        try:
                    

            self.running = False
            self.stop_event.set()
            

            if self.hotkey_registered:
                try:
                    keyboard.remove_hotkey(self.hotkey)
                except:
                    pass
            

            if self.restore_hotkey_registered:
                try:
                    keyboard.remove_hotkey(self.restore_hotkey)
                except:
                    pass
            

            if hasattr(self, 'tray_icon'):
                try:
                    self.tray_icon.stop()
                except:
                    pass
            elif hasattr(self, 'tray_icon_info'):
                try:
                    win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, self.tray_icon_info)
                except:
                    pass
            

            self.root.destroy()
        except Exception as e:
            logging.error(self.lang_strings['error_exit_app'].format(error=e))
            os._exit(0)
    

    def load_config(self):
        config = configparser.ConfigParser()
        if os.path.exists(self.config_file):
            try:
                                   

                try:
                    config.read(self.config_file, encoding='utf-8')
                except UnicodeDecodeError:
                                                       

                    config.read(self.config_file)
                

                if 'SourcePaths' in config:
                    self.source_paths = []
                    for key in config['SourcePaths']:
                        if key.startswith('path_'):
                            self.source_paths.append(config['SourcePaths'][key])
                                   

                elif config.has_option('Settings', 'SourcePaths'):
                    try:
                        self.source_paths = eval(config.get('Settings', 'SourcePaths'))
                    except:
                        self.source_paths = []
                else:
                    self.source_paths = []
                

                if config.has_option('Settings', 'dest_dir'):
                    self.dest_dir = config.get('Settings', 'dest_dir')
                else:
                    self.dest_dir = ''
                

                if config.has_option('Settings', 'interval'):
                    self.interval = config.getint('Settings', 'interval')
                

                if config.has_option('Settings', 'maxbackups'):
                    self.max_backups = config.getint('Settings', 'maxbackups')
                

                if config.has_option('Settings', 'hotkeyenabled'):
                    hotkey_enabled = config.getboolean('Settings', 'hotkeyenabled')
                    self.hotkey_var.set(hotkey_enabled)
                

                if config.has_option('Settings', 'hotkey'):
                    self.hotkey = config.get('Settings', 'hotkey')
                else:
                    self.hotkey = "ctrl+F1"
                

                if config.has_option('Settings', 'restorehotkeyenabled'):
                    restore_hotkey_enabled = config.getboolean('Settings', 'restorehotkeyenabled')
                    self.restore_hotkey_var.set(restore_hotkey_enabled)
                

                if config.has_option('Settings', 'restorehotkey'):
                    self.restore_hotkey = config.get('Settings', 'restorehotkey')
                else:
                    self.restore_hotkey = "ctrl+F2"
                

                if config.has_option('Settings', 'language'):
                    lang = config.get('Settings', 'language')
                    if lang in LANGUAGES:
                        self.current_lang = lang
                        self.lang_strings = LANGUAGES[self.current_lang]
                

                if config.has_option('Settings', 'suffixtype'):
                    self.suffix_type = config.get('Settings', 'suffixtype')
                

                if config.has_option('Settings', 'customsuffix'):
                    self.custom_suffix = config.get('Settings', 'customsuffix')
                

                if config.has_option('Settings', 'duplicatehandling'):
                    self.duplicate_handling = config.get('Settings', 'duplicatehandling')
                

                if config.has_option('Settings', 'backupcounter'):
                    self.backup_counter = config.getint('Settings', 'backupcounter')
                

                if config.has_option('Settings', 'startnumber'):
                    self.start_number = config.getint('Settings', 'startnumber')
                

                if config.has_option('Settings', 'autonumbermode'):
                    self.auto_number_mode = config.getboolean('Settings', 'autonumbermode')
                else:
                                

                    self.auto_number_mode = True
                

                if config.has_option('Settings', 'skiphidden'):
                    self.skip_hidden = config.getboolean('Settings', 'skiphidden')
                    self.skip_hidden_var.set(self.skip_hidden)
                

                self.root.after(100, self._enable_hotkeys)
                

                self.root.after(150, self.update_source_list)
                self.root.after(150, self.update_backup_list)
                

                if self.dest_dir:
                    self.root.after(200, self.update_source_path_file)
                    

            except Exception as e:
                logging.error(self.lang_strings['error_load_config'].format(error=e))
    

    def save_config(self):
        config = configparser.ConfigParser()
        config['Settings'] = {

            'dest_dir': self.dest_dir,

            'interval': str(self.interval),

            'maxbackups': str(self.max_backups),

            'hotkeyenabled': str(self.hotkey_var.get()),

            'hotkey': self.hotkey,

            'restorehotkeyenabled': str(self.restore_hotkey_var.get()),

            'restorehotkey': self.restore_hotkey,

            'language': self.current_lang,

            'suffixtype': self.suffix_type,

            'customsuffix': self.custom_suffix,

            'duplicatehandling': self.duplicate_handling,

            'backupcounter': str(self.backup_counter),

            'startnumber': str(self.start_number),

            'autonumbermode': str(self.auto_number_mode),

            'skiphidden': str(self.skip_hidden_var.get())

        }
        

        config['SourcePaths'] = {}
        if self.source_paths:
            for i, path in enumerate(self.source_paths):
                config['SourcePaths'][f'path_{i+1}'] = path
        

        try:
            with open(self.config_file, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
        except Exception as e:
            logging.error(self.lang_strings['error_save_config'].format(error=e))
    

    def create_widgets(self):
                           

        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=10, pady=5)
        

        lang_btn_frame = ttk.Frame(top_frame)
        lang_btn_frame.pack(side="right")
        

        self.btn_zh = ttk.Button(lang_btn_frame, text=self.lang_strings['chinese'], command=lambda: self.set_language('zh'))
        self.btn_zh.pack(side="left", padx=2)
        self.btn_en = ttk.Button(lang_btn_frame, text=self.lang_strings['english'], command=lambda: self.set_language('en'))
        self.btn_en.pack(side="left", padx=2)
        

        self.root.after(100, self.update_lang_button_state)
        

        self.tab_control = ttk.Notebook(self.root)
        tab_backup = ttk.Frame(self.tab_control)
        tab_manage = ttk.Frame(self.tab_control)
        tab_settings = ttk.Frame(self.tab_control)             
        self.tab_control.add(tab_backup, text=self.lang_strings['tab_backup'])
        self.tab_control.add(tab_manage, text=self.lang_strings['tab_manage'])
        self.tab_control.add(tab_settings, text=self.lang_strings['tab_settings'])
        self.tab_control.pack(expand=1, fill="both", padx=10, pady=5)
        

        self.frame_settings = ttk.LabelFrame(tab_backup, text=self.lang_strings['frame_settings'])
        self.frame_settings.pack(fill="both", expand=True, padx=10, pady=5)
        

        self.frame_settings.columnconfigure(1, weight=1)
        self.frame_settings.rowconfigure(0, weight=1)             
        

        main_frame = ttk.Frame(self.frame_settings)
        main_frame.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)
        

        self.source_frame = ttk.LabelFrame(main_frame, text=self.lang_strings['source_paths'])
        self.source_frame.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        self.source_frame.columnconfigure(1, weight=1)
        self.source_frame.rowconfigure(0, weight=1)
        

        self.source_listbox = tk.Listbox(self.source_frame, width=50, height=8, selectmode=tk.EXTENDED)                  
        self.source_listbox.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        

        self.source_btn_frame = ttk.Frame(self.source_frame)
        self.source_btn_frame.grid(row=0, column=2, padx=5, pady=5, sticky="nw")
        

        self.btn_add_file = ttk.Button(self.source_btn_frame, text=self.lang_strings['add_file'], command=lambda: self.add_source('file'))
        self.btn_add_file.pack(fill="x", padx=2, pady=2)
        self.btn_add_dir = ttk.Button(self.source_btn_frame, text=self.lang_strings['add_dir'], command=lambda: self.add_source('dir'))
        self.btn_add_dir.pack(fill="x", padx=2, pady=2)
        self.btn_remove_source = ttk.Button(self.source_btn_frame, text=self.lang_strings['remove'], command=self.remove_source)
        self.btn_remove_source.pack(fill="x", padx=2, pady=2)
        self.btn_open_source = ttk.Button(self.source_btn_frame, text=self.lang_strings['open_location'], command=self.open_source_location)
        self.btn_open_source.pack(fill="x", padx=2, pady=2)
        

        self.dest_frame = ttk.LabelFrame(main_frame, text=self.lang_strings['dest_dir'])
        self.dest_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        self.dest_frame.columnconfigure(1, weight=1)
        

        base_font_size = int(FONT_CONFIG['base_size'] * FONT_CONFIG['scale_factor'])
        entry_font_size = max(base_font_size, 16)            
        entry_font = ('Microsoft YaHei', entry_font_size)
        

        self.dest_entry = ttk.Entry(self.dest_frame, font=entry_font)
        self.dest_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.dest_entry.insert(0, self.dest_dir)
        

        self.dest_btn_frame = ttk.Frame(self.dest_frame)
        self.dest_btn_frame.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.btn_browse_dest = ttk.Button(self.dest_btn_frame, text=self.lang_strings['browse'], command=self.select_dest_dir)
        self.btn_browse_dest.pack(fill="x", padx=2, pady=2)
        self.btn_open_dest = ttk.Button(self.dest_btn_frame, text=self.lang_strings['open_location'], command=self.open_dest_location)
        self.btn_open_dest.pack(fill="x", padx=2, pady=2)
        

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=10, sticky="ew")
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        

        self.start_btn = ttk.Button(button_frame, text=self.lang_strings['start_backup'], command=self.toggle_backup)
        self.start_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        

        self.btn_manual_backup = ttk.Button(button_frame, text=self.lang_strings['manual_backup'], command=self.manual_backup)
        self.btn_manual_backup.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        

        manage_frame = ttk.Frame(tab_manage)
        manage_frame.pack(fill="both", expand=True, padx=10, pady=5)
        

        tree_container = ttk.Frame(manage_frame)
        tree_container.pack(fill="both", expand=True, padx=5, pady=5)
        

        self.backup_list = ttk.Treeview(tree_container, columns=("name", "date", "keep"), show="headings")
        self.backup_list.heading("name", text=self.lang_strings['backup_name'])
        self.backup_list.heading("date", text=self.lang_strings['backup_time'])
        self.backup_list.heading("keep", text=self.lang_strings['keep_backup'])
        self.backup_list.column("name", width=250)
        self.backup_list.column("date", width=200, stretch=True)
        self.backup_list.column("keep", width=50, anchor="center")
        

        self._apply_font_scaling_to_treeview()
        

        v_scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.backup_list.yview)
        v_scrollbar.pack(side="right", fill="y")
        

        h_scrollbar = ttk.Scrollbar(tree_container, orient="horizontal", command=self.backup_list.xview)
        h_scrollbar.pack(side="bottom", fill="x")
        

        self.backup_list.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        

        self.backup_list.bind("<ButtonRelease-1>", self.on_backup_list_click)
                

        self.backup_list.bind("<Double-Button-1>", self.on_backup_list_double_click)
        

        self.backup_list.pack(side="left", fill="both", expand=True)
        

        btn_frame = ttk.Frame(tab_manage)
        btn_frame.pack(fill="x", padx=10, pady=10)
        

        self.btn_refresh = ttk.Button(btn_frame, text=self.lang_strings['refresh_list'], command=self.update_backup_list)
        self.btn_refresh.pack(side="left", padx=5, pady=5)
        self.btn_delete = ttk.Button(btn_frame, text=self.lang_strings['delete_selected'], command=self.delete_selected)
        self.btn_delete.pack(side="left", padx=5, pady=5)
        self.btn_restore = ttk.Button(btn_frame, text=self.lang_strings['restore_selected'], command=self.restore_selected)
        self.btn_restore.pack(side="left", padx=5, pady=5)
        self.btn_open_backup = ttk.Button(btn_frame, text=self.lang_strings['open_backup'], command=self.open_backup_location)
        self.btn_open_backup.pack(side="left", padx=5, pady=5)
        

        self.btn_rename = ttk.Button(btn_frame, text=self.lang_strings['rename_backup'], command=self.rename_selected_backup)
        self.btn_rename.pack(side="left", padx=5, pady=5)
        

        settings_frame = ttk.Frame(tab_settings)
        settings_frame.pack(fill="both", expand=True, padx=10, pady=5)
        

        self.frame_backup_options = ttk.LabelFrame(settings_frame, text=self.lang_strings['backup_options'])
        self.frame_backup_options.pack(fill="x", padx=10, pady=5)
        

        self.frame_backup_options.columnconfigure(1, weight=1)
        

        interval_max_frame = ttk.Frame(self.frame_backup_options)
        interval_max_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        

        base_font_size = int(FONT_CONFIG['base_size'] * FONT_CONFIG['scale_factor'])
        entry_font_size = max(base_font_size, 16)            
        entry_font = ('Microsoft YaHei', entry_font_size)
        

        self.interval_label = ttk.Label(interval_max_frame, text=self.lang_strings['interval'])
        self.interval_label.pack(side="left", padx=5, pady=5)
        

        self.interval_entry = ttk.Entry(interval_max_frame, width=15, font=entry_font)
        self.interval_entry.insert(0, str(self.interval))
        self.interval_entry.pack(side="left", padx=5, pady=5)
        

        ttk.Separator(interval_max_frame, orient='vertical').pack(side="left", fill="y", padx=10, pady=5)
        

        self.max_backups_label = ttk.Label(interval_max_frame, text=self.lang_strings['max_backups'])
        self.max_backups_label.pack(side="left", padx=5, pady=5)
        

        self.max_backups_entry = ttk.Entry(interval_max_frame, width=15, font=entry_font)
        self.max_backups_entry.insert(0, str(self.max_backups))
        self.max_backups_entry.pack(side="left", padx=5, pady=5)
        

        backup_hotkey_frame = ttk.Frame(self.frame_backup_options)
        backup_hotkey_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        

        self.hotkey_label = ttk.Label(backup_hotkey_frame, text=self.lang_strings['hotkey'])
        self.hotkey_label.pack(side="left", padx=5, pady=5)
        

        self.hotkey_entry = ttk.Entry(backup_hotkey_frame, width=15, font=entry_font)
        self.hotkey_entry.insert(0, self.hotkey)
        self.hotkey_entry.pack(side="left", padx=5, pady=5)
        

        self.check_hotkey = ttk.Checkbutton(backup_hotkey_frame, text=self.lang_strings['enable_hotkey'], 

                        variable=self.hotkey_var, command=self.toggle_hotkey)
        self.check_hotkey.pack(side="left", padx=5, pady=5)
        

        restore_hotkey_frame = ttk.Frame(self.frame_backup_options)
        restore_hotkey_frame.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        

        self.restore_hotkey_label = ttk.Label(restore_hotkey_frame, text=self.lang_strings['restore_hotkey'])
        self.restore_hotkey_label.pack(side="left", padx=5, pady=5)
        

        self.restore_hotkey_entry = ttk.Entry(restore_hotkey_frame, width=15, font=entry_font)
        self.restore_hotkey_entry.insert(0, self.restore_hotkey)
        self.restore_hotkey_entry.pack(side="left", padx=5, pady=5)
        

        self.check_restore_hotkey = ttk.Checkbutton(restore_hotkey_frame, text=self.lang_strings['enable_restore_hotkey'], 

                        variable=self.restore_hotkey_var, command=self.toggle_restore_hotkey)
        self.check_restore_hotkey.pack(side="left", padx=5, pady=5)
        

        separator = ttk.Separator(self.frame_backup_options, orient='horizontal')
        separator.grid(row=4, column=0, columnspan=3, padx=5, pady=10, sticky="ew")
        

        self.suffix_type_label = ttk.Label(self.frame_backup_options, text=self.lang_strings['suffix_type'])
        self.suffix_type_label.grid(row=5, column=0, padx=5, pady=5, sticky="e")
        

        suffix_frame = ttk.Frame(self.frame_backup_options)
        suffix_frame.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        

        self.suffix_type_var = tk.StringVar(value=self.suffix_type)
        self.suffix_timestamp_radio = ttk.Radiobutton(suffix_frame, text=self.lang_strings['suffix_timestamp'], 

                                                      variable=self.suffix_type_var, value="timestamp")
        self.suffix_timestamp_radio.pack(side="left", padx=5)
        

        self.suffix_number_radio = ttk.Radiobutton(suffix_frame, text=self.lang_strings['suffix_number'], 

                                                   variable=self.suffix_type_var, value="number")
        self.suffix_number_radio.pack(side="left", padx=5)
        

        self.suffix_custom_radio = ttk.Radiobutton(suffix_frame, text=self.lang_strings['suffix_custom'], 

                                                   variable=self.suffix_type_var, value="custom")
        self.suffix_custom_radio.pack(side="left", padx=5)
        

        self.number_mode_frame = ttk.Frame(self.frame_backup_options)
                                      

        self.custom_suffix_label = ttk.Label(self.frame_backup_options, text=self.lang_strings['custom_suffix'])
        self.custom_suffix_label.grid(row=6, column=0, padx=5, pady=5, sticky="e")
        

        base_font_size = int(FONT_CONFIG['base_size'] * FONT_CONFIG['scale_factor'])
        entry_font_size = max(base_font_size, 16)            
        entry_font = ('Microsoft YaHei', entry_font_size)
        

        self.custom_suffix_entry = ttk.Entry(self.frame_backup_options, font=entry_font)
        self.custom_suffix_entry.insert(0, self.custom_suffix)
        self.custom_suffix_entry.grid(row=6, column=1, padx=5, pady=5, sticky="we")
        

        self.suffix_type_var.trace_add('write', self.update_suffix_entry_state)
        

        self.duplicates_label = ttk.Label(self.frame_backup_options, text=self.lang_strings['handle_duplicates'])
        self.duplicates_label.grid(row=7, column=0, padx=5, pady=5, sticky="e")
        

        duplicates_frame = ttk.Frame(self.frame_backup_options)
        duplicates_frame.grid(row=7, column=1, padx=5, pady=5, sticky="w")
        

        self.duplicates_var = tk.StringVar(value=self.duplicate_handling)
        self.duplicates_overwrite_radio = ttk.Radiobutton(duplicates_frame, text=self.lang_strings['duplicates_overwrite'], 

                                                          variable=self.duplicates_var, value="overwrite")
        self.duplicates_overwrite_radio.pack(side="left", padx=5)
        

        self.duplicates_rename_radio = ttk.Radiobutton(duplicates_frame, text=self.lang_strings['duplicates_rename'], 

                                                       variable=self.duplicates_var, value="rename")
        self.duplicates_rename_radio.pack(side="left", padx=5)
        

        self.duplicates_skip_radio = ttk.Radiobutton(duplicates_frame, text=self.lang_strings['duplicates_skip'], 

                                                     variable=self.duplicates_var, value="skip")
        self.duplicates_skip_radio.pack(side="left", padx=5)
        

        counter_frame = ttk.Frame(self.frame_backup_options)
        counter_frame.grid(row=8, column=0, columnspan=2, padx=5, pady=5, sticky="we")
        

        self.counter_var = tk.StringVar(value=str(self.backup_counter))
        self.backup_count_label = ttk.Label(counter_frame, text="")
                   

        def update_counter_display(*args):
            self.backup_count_label.config(text=self.lang_strings['backup_count'].format(count=self.counter_var.get()))
        

        self.counter_var.trace_add('write', update_counter_display)
        

        self.number_mode_label = ttk.Label(counter_frame, text=self.lang_strings['number_mode'])
                  

        number_mode_radio_frame = ttk.Frame(counter_frame)
                    

        self.auto_number_mode = True
        self.auto_number_mode_var = tk.BooleanVar(value=self.auto_number_mode)
        self.auto_number_radio = ttk.Radiobutton(number_mode_radio_frame, text=self.lang_strings['auto_number'], 

                                                variable=self.auto_number_mode_var, value=True)
                    

        self.manual_number_radio = ttk.Radiobutton(number_mode_radio_frame, text=self.lang_strings['manual_number'], 

                                                  variable=self.auto_number_mode_var, value=False)
                    

        self.suffix_type_var.trace_add('write', self.update_number_mode_visibility)
        

        self.btn_set_start_number = ttk.Button(counter_frame, text=self.lang_strings['set_start_number'], command=self.set_start_number)
                    

        skip_hidden_frame = ttk.Frame(self.frame_backup_options)
        skip_hidden_frame.grid(row=9, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        

        self.skip_hidden_var = tk.BooleanVar(value=self.skip_hidden)
        self.check_skip_hidden = ttk.Checkbutton(skip_hidden_frame, text=self.lang_strings['skip_hidden'], 

                        variable=self.skip_hidden_var)
        self.check_skip_hidden.pack(side="left", padx=5, pady=5)
        

        save_btn_frame = ttk.Frame(settings_frame)
        save_btn_frame.pack(fill="x", padx=10, pady=10)
        

        buttons_container = ttk.Frame(save_btn_frame)
        buttons_container.pack(fill="x", expand=True)
        

        buttons_container.columnconfigure(0, weight=1)
        buttons_container.columnconfigure(1, weight=1)
        

        self.btn_save_settings = ttk.Button(buttons_container, text=self.lang_strings['save_settings'], command=self.save_advanced_settings)
        self.btn_save_settings.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        

        self.btn_restore_defaults = ttk.Button(buttons_container, text=self.lang_strings['restore_defaults'], command=self.restore_default_settings)
        self.btn_restore_defaults.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        

        export_import_frame = ttk.Frame(tab_settings)
        export_import_frame.pack(fill="x", padx=10, pady=10)
        

        self.btn_export_settings = ttk.Button(export_import_frame, text=self.lang_strings['export_settings'], command=self.export_settings)
        self.btn_export_settings.pack(side="left", padx=5, pady=5, fill="x", expand=True)
        

        self.btn_import_settings = ttk.Button(export_import_frame, text=self.lang_strings['import_settings'], command=self.import_settings)
        self.btn_import_settings.pack(side="right", padx=5, pady=5, fill="x", expand=True)
        

        self.status_var = tk.StringVar()
        self.status_var.set(self.lang_strings['status_ready'])
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(side="bottom", fill="x")
        

        self.update_source_list()
        

        self.update_suffix_entry_state()
    

    def export_settings(self):
        

        try:
                          

            default_dir = os.path.dirname(os.path.abspath(__file__))
            

            file_path = filedialog.asksaveasfilename(

                title=self.lang_strings['export_settings'],

                defaultextension=".ini",

                filetypes=[("INI Files", "*.ini"), ("All Files", "*")],

                initialdir=default_dir,

                initialfile="backup_config_export.ini"

            )
            

            if not file_path:
                return           
            

            temp_config = configparser.ConfigParser()
            

            temp_config['Settings'] = {}
            

            interval_value = self.interval_entry.get().strip()
            temp_config['Settings']['interval'] = interval_value
            

            max_backups_value = self.max_backups_entry.get().strip()
            temp_config['Settings']['maxbackups'] = max_backups_value
            

            suffix_type_value = self.suffix_type_var.get()
            temp_config['Settings']['suffixtype'] = suffix_type_value
            

            custom_suffix_value = self.custom_suffix_entry.get().strip()
            temp_config['Settings']['customsuffix'] = custom_suffix_value
            

            duplicates_value = self.duplicates_var.get()
            temp_config['Settings']['duplicatehandling'] = duplicates_value
            

            temp_config['Settings']['language'] = self.current_lang
            

            temp_config['Settings']['hotkeyenabled'] = str(self.hotkey_var.get())
            

            temp_config['Settings']['restorehotkeyenabled'] = str(self.restore_hotkey_var.get())
            

            hotkey_value = self.hotkey_entry.get().strip()
            temp_config['Settings']['hotkey'] = hotkey_value
            

            restore_hotkey_value = self.restore_hotkey_entry.get().strip()
            temp_config['Settings']['restorehotkey'] = restore_hotkey_value
            

            if hasattr(self, 'auto_number_mode_var'):
                temp_config['Settings']['autonumbermode'] = str(self.auto_number_mode_var.get())
            else:
                temp_config['Settings']['autonumbermode'] = ''
            

            temp_config['Settings']['dest_dir'] = self.dest_dir if self.dest_dir else ''
            

            if hasattr(self, 'skip_hidden_var'):
                temp_config['Settings']['skiphidden'] = str(self.skip_hidden_var.get())
            else:
                temp_config['Settings']['skiphidden'] = 'False'
            

            temp_config['Settings']['backupcounter'] = str(self.backup_counter)
            

            temp_config['Settings']['startnumber'] = str(self.start_number)
            

            temp_config['SourcePaths'] = {}
            if self.source_paths:
                for i, path in enumerate(self.source_paths):
                    temp_config['SourcePaths'][f'path_{i+1}'] = path
            

            with open(file_path, 'w', encoding='utf-8') as configfile:
                temp_config.write(configfile)
            

            messagebox.showinfo(self.lang_strings['info'], 

                              f"{self.lang_strings.get('settings_exported_success', 'Settings successfully exported to')}: {file_path}")
        except Exception as e:
            logging.error(self.lang_strings['error_export_settings'].format(error=e))
            messagebox.showerror(self.lang_strings['error'], 

                               f"{self.lang_strings.get('export_settings_failed', 'Failed to export settings')}: {str(e)}")
    

    def merge_config(self, imported_config):
        

        try:
                       

            backup_config = configparser.ConfigParser()
            

            backup_config['Settings'] = {

                'dest_dir': self.dest_dir,

                'interval': str(self.interval),

                'maxbackups': str(self.max_backups),

                'hotkeyenabled': str(self.hotkey_var.get()),

                'hotkey': self.hotkey,

                'restorehotkeyenabled': str(self.restore_hotkey_var.get()),

                'restorehotkey': self.restore_hotkey,

                'language': self.current_lang,

                'suffixtype': self.suffix_type,

                'customsuffix': self.custom_suffix,

                'duplicatehandling': self.duplicate_handling,

                'skiphidden': str(self.skip_hidden_var.get()) if hasattr(self, 'skip_hidden_var') else 'False',

                'backupcounter': str(self.backup_counter),

                'startnumber': str(self.start_number),

                'autonumbermode': str(self.auto_number_mode)

            }
            

            if self.source_paths:
                backup_config['SourcePaths'] = {}
                for i, path in enumerate(self.source_paths):
                    backup_config['SourcePaths'][f'path_{i+1}'] = path
            

            merged_config = configparser.ConfigParser()
            merged_config['Settings'] = {}
            

            settings_items = [

                ('dest_dir', 'dest_dir'),

                ('interval', 'interval'),

                ('maxbackups', 'maxbackups'),

                ('hotkeyenabled', 'hotkeyenabled'),

                ('hotkey', 'hotkey'),

                ('restorehotkeyenabled', 'restorehotkeyenabled'),

                ('restorehotkey', 'restorehotkey'),

                ('language', 'language'),

                ('suffixtype', 'suffixtype'),

                ('customsuffix', 'customsuffix'),

                ('duplicatehandling', 'duplicatehandling'),

                ('skiphidden', 'skiphidden'),

                ('backupcounter', 'backupcounter'),

                ('startnumber', 'startnumber'),

                ('autonumbermode', 'autonumbermode')

            ]
            

            for imported_key, backup_key in settings_items:
                if imported_config.has_option('Settings', imported_key):
                    imported_value = imported_config.get('Settings', imported_key)
                                              

                    if backup_key == 'dest_dir' and not imported_value:
                                        

                        if backup_config.has_option('Settings', backup_key):
                            merged_config['Settings'][backup_key] = backup_config.get('Settings', backup_key)
                    else:
                                          

                        merged_config['Settings'][backup_key] = imported_value
                elif backup_config.has_option('Settings', backup_key):
                    merged_config['Settings'][backup_key] = backup_config.get('Settings', backup_key)
            

            merged_config['SourcePaths'] = {}
            

            if 'SourcePaths' in backup_config:
                for key in backup_config['SourcePaths']:
                    if key.startswith('path_'):
                        merged_config['SourcePaths'][key] = backup_config['SourcePaths'][key]
            

            if 'SourcePaths' in imported_config:
                for key in imported_config['SourcePaths']:
                    if key.startswith('path_'):
                        merged_config['SourcePaths'][key] = imported_config['SourcePaths'][key]
            

            with open(self.config_file, 'w', encoding='utf-8') as configfile:
                merged_config.write(configfile)
            

            self.load_config()
            

            return True
        except Exception as e:
            logging.error(f"合并配置失败: {e}")
            return False
    

    def import_settings(self):
        

        try:
                       

            file_path = filedialog.askopenfilename(

                title=self.lang_strings['import_settings'],

                filetypes=[("INI Files", "*.ini"), ("All Files", "*")]

            )
            

            if not file_path:
                return           
            

            if not os.path.exists(file_path):
                messagebox.showerror(self.lang_strings.get('error', "Error"), 

                                   self.lang_strings.get('file_not_found', "File not found"))
                return
            

            imported_config = configparser.ConfigParser()
            

            backup_config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backup_config_backup.ini")
            

            try:
                                                   

                try:
                    imported_config.read(file_path, encoding='utf-8')
                except UnicodeDecodeError:
                                            

                    imported_config.read(file_path)
                

                if 'Settings' not in imported_config:
                    raise ValueError(self.lang_strings.get('invalid_config_format', "Invalid configuration file format"))
                

                if os.path.exists(self.config_file):
                    shutil.copy2(self.config_file, backup_config_path)
                

                if self.merge_config(imported_config):
                              

                    self.update_ui_texts()
                    

                    self.update_source_list()
                    

                    self.update_backup_list()
                    

                    self.update_source_path_file()
                    

                    self.dest_entry.delete(0, tk.END)
                    if self.dest_dir:
                        self.dest_entry.insert(0, self.dest_dir)
                    

                    self.interval_entry.delete(0, tk.END)
                    self.interval_entry.insert(0, str(self.interval))
                    

                    self.max_backups_entry.delete(0, tk.END)
                    self.max_backups_entry.insert(0, str(self.max_backups))
                    

                    if hasattr(self, 'skip_hidden_var'):
                        self.skip_hidden_var.set(self.skip_hidden)
                    

                    messagebox.showinfo(self.lang_strings.get('info', "Info"), 

                                      self.lang_strings.get('settings_imported_success', "Settings successfully imported and applied"))
                else:
                    raise Exception(self.lang_strings.get('config_merge_failed', "Configuration merge failed"))
            except Exception as e:
                                

                if os.path.exists(backup_config_path):
                    try:
                        shutil.copy2(backup_config_path, self.config_file)
                        self.load_config()
                        self.update_ui_texts()
                    except:
                        pass
                raise e
        except Exception as e:
            logging.error(self.lang_strings['error_import_settings'].format(error=e))
            messagebox.showerror(self.lang_strings.get('error', "Error"), 

                               f"{self.lang_strings.get('import_settings_failed', 'Failed to import settings')}: {str(e)}")
    

    def _apply_font_scaling_to_treeview(self):
        

        try:
                              

            base_font_size = int(FONT_CONFIG['base_size'] * FONT_CONFIG['scale_factor'])
            

            treeview_font_size = max(base_font_size - 1, 14)
            

            row_height = int(treeview_font_size * 2.2)                
            

            style = ttk.Style()
            style.configure('Treeview', 

                          font=('Microsoft YaHei', treeview_font_size), 

                          rowheight=row_height)
            style.configure('Treeview.Heading', 

                          font=('Microsoft YaHei', treeview_font_size, 'bold'))
        except Exception as e:
            logging.error(f"应用字体缩放到表格失败: {e}")
    

    def update_lang_button_state(self):
        

        if self.current_lang == 'zh':
            self.btn_zh.config(state='disabled')
            self.btn_en.config(state='normal')
        else:
            self.btn_zh.config(state='normal')
            self.btn_en.config(state='disabled')
    

    def update_source_list(self):
        self.source_listbox.delete(0, tk.END)
        for path in self.source_paths:
            self.source_listbox.insert(tk.END, path)
    

    def add_source(self, source_type):
        if source_type == 'file':
            files = filedialog.askopenfilenames(title=self.lang_strings['add_file'])
            if files:
                new_paths = []
                for file in files:
                                

                    if self.dest_dir and self.check_path_relationship(file, self.dest_dir):
                        messagebox.showerror("Error", self.lang_strings['error_path_contains'])
                        continue
                    new_paths.append(file)
                self.source_paths.extend(new_paths)
        else:
            directory = filedialog.askdirectory(title=self.lang_strings['add_dir'])
            if directory:
                            

                if self.dest_dir and self.check_path_relationship(directory, self.dest_dir):
                    messagebox.showerror("Error", self.lang_strings['error_path_contains'])
                    return
                self.source_paths.append(directory)
        

        self.source_paths = list(set(self.source_paths))
        self.update_source_list()
        self.save_config()
    

    def remove_source(self):
        selected = self.source_listbox.curselection()
        if selected:
                                        

            selected_indices = sorted(selected, reverse=True)
            

            for index in selected_indices:
                self.source_paths.pop(index)
            

            self.update_source_list()
            self.save_config()
    

    def open_source_location(self):
        selected = self.source_listbox.curselection()
        if not selected:
            return
        

        for index in selected:
            path = self.source_paths[index]
            if os.path.exists(path):
                if os.path.isfile(path):
                    directory = os.path.dirname(path)
                    os.startfile(directory)
                else:
                    os.startfile(path)
    

    def open_dest_location(self):
        if self.dest_dir and os.path.exists(self.dest_dir):
            os.startfile(self.dest_dir)
    

    def open_backup_location(self):
        selected = self.backup_list.selection()
        if not selected:
            return
        

        item = selected[0]
        values = self.backup_list.item(item, "values")
        if values and len(values) >= 2:
                           

            backup_name = values[0]
            backup_path = os.path.join(self.dest_dir, backup_name)
            if os.path.exists(backup_path):
                os.startfile(backup_path)
    

    def select_dest_dir(self):
        directory = filedialog.askdirectory(title=self.lang_strings['dest_dir'])
        if directory:
                         

            for source in self.source_paths:
                if self.check_path_relationship(source, directory):
                    messagebox.showerror("Error", self.lang_strings['error_path_contains'])
                    return
            

            self.dest_dir = directory
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, directory)
            self.update_backup_list()
            self.save_config()
    

    def is_valid_hotkey(self, hotkey_str):
        if not hotkey_str:
            return False
        

        valid_modifiers = ["ctrl", "shift", "alt", "win"]
        

        keys = hotkey_str.split('+')
        

        if len(keys) < 1:
            return False
        

        for i, key in enumerate(keys):
            key_lower = key.lower()
            

            if i == len(keys) - 1:
                return len(key) > 0
                

            if key_lower not in valid_modifiers:
                return False
        

        return True
    

    def toggle_hotkey(self):
        new_hotkey = self.hotkey_entry.get().strip()
        

        if not self.is_valid_hotkey(new_hotkey):
            messagebox.showerror("Error", self.lang_strings['error_invalid_hotkey'])
            self.hotkey_var.set(False)
            return
        

        self.hotkey = new_hotkey
        

        if self.hotkey_var.get():
            try:
                if self.hotkey_registered:
                    try:
                        keyboard.remove_hotkey(self.hotkey)
                    except:
                        pass
                

                keyboard.add_hotkey(self.hotkey, self.manual_backup)
                self.hotkey_registered = True
                self.status_var.set(self.lang_strings['hotkey_enabled'].format(hotkey=self.hotkey))
            except Exception as e:
                logging.error(f"注册快捷键失败: {e}")
                messagebox.showerror("Error", self.lang_strings['error_hotkey_reg'].format(error=str(e)))
                self.hotkey_var.set(False)
                self.hotkey_registered = False
        else:
            if self.hotkey_registered:
                try:
                    keyboard.remove_hotkey(self.hotkey)
                except:
                    pass
                self.hotkey_registered = False
                self.status_var.set(self.lang_strings['hotkey_disabled'])
    

    def toggle_restore_hotkey(self):
        

        new_restore_hotkey = self.restore_hotkey_entry.get().strip()
        

        if not self.is_valid_hotkey(new_restore_hotkey):
            messagebox.showerror("Error", self.lang_strings['error_invalid_hotkey'])
            self.restore_hotkey_var.set(False)
            return
        

        self.restore_hotkey = new_restore_hotkey
        

        if self.restore_hotkey_var.get():
            try:
                if self.restore_hotkey_registered:
                    try:
                        keyboard.remove_hotkey(self.restore_hotkey)
                    except:
                        pass
                

                keyboard.add_hotkey(self.restore_hotkey, lambda: self.restore_selected(skip_confirm=True, skip_success_dialog=True))
                self.restore_hotkey_registered = True
                self.status_var.set(self.lang_strings['restore_hotkey_enabled'].format(hotkey=self.restore_hotkey))
            except Exception as e:
                logging.error(self.lang_strings['error_register_restore_hotkey'].format(error=e))
                messagebox.showerror("Error", self.lang_strings['error_hotkey_reg'].format(error=str(e)))
                self.restore_hotkey_var.set(False)
                self.restore_hotkey_registered = False
        else:
            if self.restore_hotkey_registered:
                try:
                    keyboard.remove_hotkey(self.restore_hotkey)
                except:
                    pass
                self.restore_hotkey_registered = False
                self.status_var.set(self.lang_strings['restore_hotkey_disabled'])
    

    def update_suffix_entry_state(self, *args):
        

        if self.suffix_type_var.get() == "custom":
            self.custom_suffix_entry.config(state="normal")
        else:
            self.custom_suffix_entry.config(state="disabled")
    

    def toggle_backup(self):
        if self.running:
                  

            self.running = False
            self.stop_event.set()
            self.start_btn.config(text=self.lang_strings['start_backup'])
            self.status_var.set(self.lang_strings['backup_stopped'])
            self.set_icon(active=False)
            self.create_tray_icon(active=False)
        else:
                  

            if not self.source_paths:
                messagebox.showerror("Error", self.lang_strings['error_no_source'])
                return
            

            if not self.dest_dir:
                messagebox.showerror("Error", self.lang_strings['error_no_dest'])
                return
            

            try:
                interval = int(self.interval_entry.get())
                if interval <= 0:
                    messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                    return
                self.interval = interval
            except ValueError:
                messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                return
            

            try:
                max_backups = int(self.max_backups_entry.get())
                if max_backups <= 0:
                    messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                    return
                self.max_backups = max_backups
            except ValueError:
                messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                return
            

            self.save_config()
            

            self.running = True
            self.stop_event.clear()
            self.start_btn.config(text=self.lang_strings['stop_backup'])
            self.status_var.set(self.lang_strings['backup_started'].format(interval=self.interval))
            self.set_icon(active=True)
            self.create_tray_icon(active=True)
            

            self.backup_thread = threading.Thread(target=self.backup_loop)
            self.backup_thread.daemon = True
            self.backup_thread.start()
    

    def is_hidden(self, filepath):
        

        try:
            attrs = win32file.GetFileAttributesW(filepath)
            return attrs & (win32con.FILE_ATTRIBUTE_HIDDEN | win32con.FILE_ATTRIBUTE_SYSTEM) != 0
        except Exception as e:
            logging.error(self.lang_strings['error_checking_file_attributes'].format(error=e))
            return False
    

    def copy_with_attributes(self, src, dst, is_source_item=False, is_restore=False):
        try:
                            

            if is_restore:
                filename = os.path.basename(src)
                if filename in [".keep", ".file", ".source_path", ".folder"] or \
                   (filename.startswith(".") and (filename.endswith("keep") or filename.endswith("file"))):
                    return
            

            if self.skip_hidden and self.is_hidden(src) and not is_source_item and not is_restore:
                logging.info(self.lang_strings['skipping_hidden_file'].format(path=src))
                return
                

            if os.path.isfile(src):
                             

                shutil.copy2(src, dst)
                

                src_attrs = win32file.GetFileAttributesW(src)
                

                win32file.SetFileAttributesW(dst, src_attrs)
                

            elif os.path.isdir(src):
                                              

                if os.path.exists(dst):
                                    

                    self._remove_item_with_readonly_handling(dst)
                

                os.makedirs(dst, exist_ok=True)
                

                for item in os.listdir(src):
                    src_item = os.path.join(src, item)
                    dst_item = os.path.join(dst, item)
                    

                    self.copy_with_attributes(src_item, dst_item, is_source_item=False)
                

                src_attrs = win32file.GetFileAttributesW(src)
                

                win32file.SetFileAttributesW(dst, src_attrs)
        except Exception as e:
            logging.error(self.lang_strings['error_copying_file_attributes'].format(error=e))
            raise
    

    def backup_loop(self):
        

        while self.running and not self.stop_event.is_set():
            try:
                self.perform_backup()
                

                if self.stop_event.wait(self.interval):
                    break
                    

            except Exception as e:
                logging.error(self.lang_strings['error_during_backup'].format(error=e))
                self.status_var.set(self.lang_strings['error_backup'].format(error=str(e)))
                

                if self.stop_event.wait(10):
                    break
        

        if self.running:
            self.running = False
            self.root.after(0, lambda: self.start_btn.config(text=self.lang_strings['start_backup']))
            self.root.after(0, lambda: self.status_var.set(self.lang_strings['backup_stopped']))
            self.root.after(0, lambda: self.set_icon(active=False))
            self.root.after(0, lambda: self.create_tray_icon(active=False))
    

    def get_auto_number_suffix(self):
        

        try:
            source_path_file = os.path.join(self.dest_dir, ".source_path")
            if not os.path.exists(source_path_file):
                                         

                return "_1"
            

            current_source_name = os.path.basename(self.source_paths[0])
            

            backup_items = []                             
            with open(source_path_file, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line:
                        parts = line.split('|')
                        if len(parts) >= 6 and parts[1] == current_source_name:
                            backup_time = parts[3]        
                            backup_name = parts[2]         
                            backup_items.append((backup_time, backup_name))
            

            backup_items.sort(key=lambda x: x[0])
            

            if len(backup_items) >= self.max_backups:
                                     

                for i, (backup_time, backup_name) in enumerate(backup_items):
                                  

                    new_number = i + 1
                    

                    base_name = backup_name
                    match = re.match(r'(.+)_\d+$', backup_name)
                    if match:
                        base_name = match.group(1)
                    

                    new_backup_name = f"{base_name}_{new_number}"
                    

                    if new_backup_name != backup_name:
                        old_path = os.path.join(self.dest_dir, backup_name)
                        new_path = os.path.join(self.dest_dir, new_backup_name)
                        

                        if os.path.exists(old_path) and not os.path.exists(new_path):
                            try:
                                os.rename(old_path, new_path)
                                logging.info(self.lang_strings['renaming_backup'].format(old_name=backup_name, new_name=new_backup_name))
                            except Exception as e:
                                logging.error(self.lang_strings['rename_backup_failed'].format(error=e))
                                

            return f"_{len(backup_items) + 1}"
        except Exception as e:
            logging.error(self.lang_strings['error_get_auto_number_suffix'].format(error=e))
                    

            return "_1"
    

    def get_backup_suffix(self):
        

        suffix = ""
        if self.suffix_type == "timestamp":
            suffix = datetime.datetime.now().strftime("_%Y%m%d_%H%M%S")
        elif self.suffix_type == "number":
            if self.auto_number_mode:
                                                    

                suffix = self.get_auto_number_suffix()
            else:
                                  

                suffix = f"_{self.backup_counter}"
        elif self.suffix_type == "custom" and self.custom_suffix:
            suffix = f"_{self.custom_suffix}"
        return suffix
    

    def perform_backup(self):
        

        if not self.source_paths or not self.dest_dir:
            return
        

        removed_paths = []
        for source_path in self.source_paths[:]:            
            if not os.path.exists(source_path):
                self.source_paths.remove(source_path)
                removed_paths.append(source_path)
                logging.warning(f"源路径不存在，已移除: {source_path}")
        

        if removed_paths:
            self.save_config()            
            removed_paths_str = "\n".join(removed_paths)
            messagebox.showwarning(

                "Warning", 

                self.lang_strings['source_path_removed'].format(paths=removed_paths_str)

            )
            

            self.root.after(0, self.update_source_list)
            

            if not self.source_paths:
                logging.warning("所有源路径都不存在，备份已取消")
                return
        

        if not os.path.exists(self.dest_dir):
            try:
                os.makedirs(self.dest_dir)
            except Exception as e:
                logging.error(f"创建备份目录失败: {e}")
                raise
        

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        

        prefix = ""
        suffix = self.get_backup_suffix()
        

        try:
            for source_path in self.source_paths:
                source_name = os.path.basename(source_path)
                

                backup_folder_name = f"{prefix}{source_name}{suffix}"
                backup_folder_path = os.path.join(self.dest_dir, backup_folder_name)
                

                if os.path.exists(backup_folder_path):
                    if self.duplicate_handling == "overwrite":
                        self._remove_item_with_readonly_handling(backup_folder_path)
                    elif self.duplicate_handling == "rename":
                                        

                        counter = 1
                        while True:
                            unique_suffix = f"_{counter}"
                            new_backup_folder_name = f"{prefix}{source_name}{suffix}{unique_suffix}"
                            new_backup_folder_path = os.path.join(self.dest_dir, new_backup_folder_name)
                            if not os.path.exists(new_backup_folder_path):
                                backup_folder_name = new_backup_folder_name
                                backup_folder_path = new_backup_folder_path
                                break
                            counter += 1
                    elif self.duplicate_handling == "skip":
                        logging.info(f"跳过备份，文件夹已存在: {backup_folder_path}")
                        continue
                

                os.makedirs(backup_folder_path)
                

                if os.path.isdir(source_path):
                                           

                    dest_path = os.path.join(backup_folder_path, source_name)
                    self.copy_with_attributes(source_path, dest_path, is_source_item=True)
                    

                    folder_marker_path = os.path.join(backup_folder_path, ".folder")
                    with open(folder_marker_path, 'w', encoding='utf-8') as f:
                        f.write(f"source_name={source_name}\n")
                        f.write(f"source_path={source_path}\n")
                else:
                                       

                    dest_path = os.path.join(backup_folder_path, source_name)
                    self.copy_with_attributes(source_path, dest_path, is_source_item=True)
                    

                    file_marker_path = os.path.join(backup_folder_path, ".file")
                    with open(file_marker_path, 'w', encoding='utf-8') as f:
                        f.write(f"source_name={source_name}\n")
                        f.write(f"source_path={source_path}\n")
            

            if self.suffix_type == "number":
                self.backup_counter += 1
            

            self.update_source_path_file()
            

            self.cleanup_old_backups_by_item()
            

            self.status_var.set(self.lang_strings['backup_success'].format(timestamp=timestamp))
            

            self.save_config()
            

            self.root.after(0, self.update_backup_list)
            

            if self.suffix_type == "number":
                self.root.after(0, lambda: self.counter_var.set(str(self.backup_counter)))
            

        except Exception as e:
            logging.error(f"备份过程中出错: {e}")
            raise
    

    def manual_backup(self):
        

        if not self.source_paths or not self.dest_dir:
            messagebox.showerror("Error", self.lang_strings['error_no_source'] if not self.source_paths else self.lang_strings['error_no_dest'])
            return
        

        try:
            self.status_var.set(self.lang_strings['manual_backup_running'])
            self.perform_backup()
            self.status_var.set(self.lang_strings['manual_backup_done'])
        except Exception as e:
            logging.error(f"手动备份失败: {e}")
            messagebox.showerror("Error", self.lang_strings['error_backup'].format(error=str(e)))
            self.status_var.set(self.lang_strings['status_ready'])
    

    def _remove_item_with_readonly_handling(self, item_path):
        

        try:
            if os.path.isdir(item_path):
                                   

                for root, dirs, files in os.walk(item_path, topdown=False):
                    for file in files:
                        file_path = os.path.join(root, file)
                                   

                        try:
                            attrs = win32file.GetFileAttributesW(file_path)
                            if attrs & win32con.FILE_ATTRIBUTE_READONLY:
                                win32file.SetFileAttributesW(file_path, attrs & ~win32con.FILE_ATTRIBUTE_READONLY)
                        except:
                            pass                   
                        try:
                            os.remove(file_path)
                        except Exception as e:
                            logging.warning(f"删除文件失败: {file_path}, 错误: {e}")
                    

                    for dir in dirs:
                        dir_path = os.path.join(root, dir)
                                   

                        try:
                            attrs = win32file.GetFileAttributesW(dir_path)
                            if attrs & win32con.FILE_ATTRIBUTE_READONLY:
                                win32file.SetFileAttributesW(dir_path, attrs & ~win32con.FILE_ATTRIBUTE_READONLY)
                        except:
                            pass                   
                        try:
                            os.rmdir(dir_path)
                        except Exception as e:
                            logging.warning(f"删除目录失败: {dir_path}, 错误: {e}")
                

                try:
                    attrs = win32file.GetFileAttributesW(item_path)
                    if attrs & win32con.FILE_ATTRIBUTE_READONLY:
                        win32file.SetFileAttributesW(item_path, attrs & ~win32con.FILE_ATTRIBUTE_READONLY)
                except:
                    pass                   
                

                try:
                    os.rmdir(item_path)
                except Exception as e:
                    logging.warning(f"删除根目录失败: {item_path}, 错误: {e}")
            else:
                      

                try:
                    attrs = win32file.GetFileAttributesW(item_path)
                    if attrs & win32con.FILE_ATTRIBUTE_READONLY:
                        win32file.SetFileAttributesW(item_path, attrs & ~win32con.FILE_ATTRIBUTE_READONLY)
                except:
                    pass                   
                

                try:
                    os.remove(item_path)
                except Exception as e:
                    logging.warning(f"删除文件失败: {item_path}, 错误: {e}")
        except Exception as e:
            logging.error(f"删除项目失败: {item_path}, 错误: {e}")
    

    def cleanup_old_backups(self):
        

        try:
                                  

            backup_items = []
            for item in os.listdir(self.dest_dir):
                item_path = os.path.join(self.dest_dir, item)
                if os.path.isfile(item_path) or os.path.isdir(item_path):
                                            

                    mtime = os.path.getmtime(item_path)
                    backup_items.append((mtime, item_path))
            

            backup_items.sort(reverse=True, key=lambda x: x[0])
            

            if len(backup_items) > self.max_backups:
                for _, item_path in backup_items[self.max_backups:]:
                    self._remove_item_with_readonly_handling(item_path)
        except Exception as e:
            logging.error(self.lang_strings['error_cleanup_old_backups'].format(error=e))
            raise
    

    def update_number_mode_visibility(self, *args):
        

        if hasattr(self, 'number_mode_label'):
            self.number_mode_label.pack_forget()
        if hasattr(self, 'auto_number_radio'):
            self.auto_number_radio.pack_forget()
        if hasattr(self, 'manual_number_radio'):
            self.manual_number_radio.pack_forget()
    

    def update_source_path_file(self):
        

        try:
                        

            if not self.dest_dir:
                                        

                source_path_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".source_path")
                

                if not self.source_paths:
                    with open(source_path_file, 'w', encoding='utf-8') as f:
                        pass
                    return
                

                with open(source_path_file, 'w', encoding='utf-8') as f:
                    for i, path in enumerate(self.source_paths):
                                      

                        f.write(f"source_path_{i+1}|{os.path.basename(path)}|||{path}|0\n")
                return
                

            source_path_file = os.path.join(self.dest_dir, ".source_path")
            

            existing_records = {}
            if os.path.exists(source_path_file):
                with open(source_path_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line:
                            parts = line.split('|')
                            if len(parts) >= 6:
                                backup_type = parts[0]
                                source_name = parts[1]
                                backup_name = parts[2]
                                backup_time = parts[3]
                                source_path = parts[4]
                                backup_count = int(parts[5])
                                

                                backup_path = os.path.join(self.dest_dir, backup_name)
                                if os.path.exists(backup_path):
                                    key = f"{backup_type}|{source_name}|{backup_name}"                        
                                    existing_records[key] = {

                                        'type': backup_type,

                                        'source_name': source_name,

                                        'backup_name': backup_name,

                                        'backup_time': backup_time,

                                        'source_path': source_path,

                                        'backup_count': backup_count

                                    }
            

            current_backups = {}
            for item in os.listdir(self.dest_dir):
                if item == ".source_path":                      
                    continue
                    

                item_path = os.path.join(self.dest_dir, item)
                if os.path.isfile(item_path) or os.path.isdir(item_path):
                            

                    ctime = os.path.getctime(item_path)
                    timestamp = datetime.datetime.fromtimestamp(ctime).strftime("%Y-%m-%d %H:%M:%S")
                    

                    original_name = item
                               

                    if hasattr(self, 'prefix_type') and self.prefix_type == "number":
                        match = re.match(r'^\d+_(.+)', item)
                        if match:
                            original_name = match.group(1)
                    elif hasattr(self, 'prefix_type') and self.prefix_type == "custom" and hasattr(self, 'custom_prefix') and self.custom_prefix:
                        prefix = f"{self.custom_prefix}_"
                        if item.startswith(prefix):
                            original_name = item[len(prefix):]
                    

                    if hasattr(self, 'suffix_type') and self.suffix_type == "timestamp":
                                                  

                        match = re.match(r'(.+)_\d{8}_\d{6}$', original_name)
                        if match:
                            original_name = match.group(1)
                    elif hasattr(self, 'suffix_type') and self.suffix_type == "number":
                                       

                        match = re.match(r'(.+)_v\d+$', original_name)
                        if match:
                            original_name = match.group(1)
                    elif hasattr(self, 'suffix_type') and self.suffix_type == "custom" and hasattr(self, 'custom_suffix') and self.custom_suffix:
                        suffix = f"_{self.custom_suffix}"
                        if original_name.endswith(suffix):
                            original_name = original_name[:-len(suffix)]
                    

                    source_path = ""
                    for path in self.source_paths:
                        if os.path.basename(path) == original_name:
                            source_path = path
                            break
                    

                    backup_type = "unknown"           
                    source_name_from_marker = ""                 
                    source_path_from_marker = ""                
                    

                    if os.path.isdir(item_path):
                                                  

                        file_marker_path = os.path.join(item_path, ".file")
                        folder_marker_path = os.path.join(item_path, ".folder")
                        

                        if os.path.exists(file_marker_path):
                            backup_type = "file"
                                                   

                            try:
                                with open(file_marker_path, 'r', encoding='utf-8') as f:
                                    for line in f:
                                        line = line.strip()
                                        if line.startswith("source_name="):
                                            source_name_from_marker = line[12:]                      
                                        elif line.startswith("source_path="):
                                            source_path_from_marker = line[12:]                      
                            except Exception as e:
                                logging.warning(f"读取.file文件失败: {e}")
                        elif os.path.exists(folder_marker_path):
                            backup_type = "folder"
                                                     

                            try:
                                with open(folder_marker_path, 'r', encoding='utf-8') as f:
                                    for line in f:
                                        line = line.strip()
                                        if line.startswith("source_name="):
                                            source_name_from_marker = line[12:]                      
                                        elif line.startswith("source_path="):
                                            source_path_from_marker = line[12:]                      
                            except Exception as e:
                                logging.warning(f"读取.folder文件失败: {e}")
                        

                        if backup_type == "unknown":
                            continue
                        

                        if source_name_from_marker:
                            original_name = source_name_from_marker
                        if source_path_from_marker:
                            source_path = source_path_from_marker
                    else:
                                         

                        backup_type = "file"
                    

                    key = f"{backup_type}|{original_name}|{item}"
                    

                    current_backups[key] = {

                        'type': backup_type,

                        'source_name': original_name,

                        'backup_name': item,

                        'backup_time': timestamp,

                        'source_path': source_path,

                        'backup_count': 0        

                    }
            

            backup_counts = {}
            for key, backup in current_backups.items():
                backup_type, source_name, _ = key.split('|')
                source_key = f"{backup_type}|{source_name}"
                if source_key not in backup_counts:
                    backup_counts[source_key] = 0
                backup_counts[source_key] += 1
            

            for key, backup in current_backups.items():
                backup_type, source_name, _ = key.split('|')
                source_key = f"{backup_type}|{source_name}"
                backup['backup_count'] = backup_counts[source_key]
            

            all_records = {}
            

            for key, backup in current_backups.items():
                all_records[key] = backup
            

            for key, record in existing_records.items():
                if key not in all_records:
                    all_records[key] = record
            

            with open(source_path_file, 'w', encoding='utf-8') as f:
                for record in all_records.values():
                    line = f"{record['type']}|{record['source_name']}|{record['backup_name']}|{record['backup_time']}|{record['source_path']}|{record['backup_count']}"
                    f.write(line + "\n")
                    

        except Exception as e:
            logging.error(self.lang_strings['error_update_source_path_file'].format(error=e))
    

    def cleanup_old_backups_by_item(self):
        

        try:
                                     

            source_path_file = os.path.join(self.dest_dir, ".source_path")
            if not os.path.exists(source_path_file):
                                              

                self.cleanup_old_backups()
                return
            

            backup_groups = {}
            

            for item in os.listdir(self.dest_dir):
                if item == ".source_path":                      
                    continue
                    

                item_path = os.path.join(self.dest_dir, item)
                if os.path.isfile(item_path) or os.path.isdir(item_path):
                            

                    ctime = os.path.getctime(item_path)
                    

                    original_name = item
                               

                    if hasattr(self, 'prefix_type') and self.prefix_type == "number":
                        match = re.match(r'^\d+_(.+)', item)
                        if match:
                            original_name = match.group(1)
                    elif hasattr(self, 'prefix_type') and self.prefix_type == "custom" and hasattr(self, 'custom_prefix') and self.custom_prefix:
                        prefix = f"{self.custom_prefix}_"
                        if item.startswith(prefix):
                            original_name = item[len(prefix):]
                    

                    if hasattr(self, 'suffix_type') and self.suffix_type == "timestamp":
                                                  

                        match = re.match(r'(.+)_\d{8}_\d{6}$', original_name)
                        if match:
                            original_name = match.group(1)
                    elif hasattr(self, 'suffix_type') and self.suffix_type == "number":
                                       

                        match = re.match(r'(.+)_v\d+$', original_name)
                        if match:
                            original_name = match.group(1)
                    elif hasattr(self, 'suffix_type') and self.suffix_type == "custom" and hasattr(self, 'custom_suffix') and self.custom_suffix:
                        suffix = f"_{self.custom_suffix}"
                        if original_name.endswith(suffix):
                            original_name = original_name[:-len(suffix)]
                    

                    backup_type = "unknown"           
                    source_name_from_marker = ""                 
                    

                    if os.path.isdir(item_path):
                                                  

                        file_marker_path = os.path.join(item_path, ".file")
                        folder_marker_path = os.path.join(item_path, ".folder")
                        

                        if os.path.exists(file_marker_path):
                            backup_type = "file"
                                               

                            try:
                                with open(file_marker_path, 'r', encoding='utf-8') as f:
                                    for line in f:
                                        line = line.strip()
                                        if line.startswith("source_name="):
                                            source_name_from_marker = line[12:]                      
                                            break
                            except Exception as e:
                                logging.warning(f"读取.file文件失败: {e}")
                        elif os.path.exists(folder_marker_path):
                            backup_type = "folder"
                                                 

                            try:
                                with open(folder_marker_path, 'r', encoding='utf-8') as f:
                                    for line in f:
                                        line = line.strip()
                                        if line.startswith("source_name="):
                                            source_name_from_marker = line[12:]                      
                                            break
                            except Exception as e:
                                logging.warning(f"读取.folder文件失败: {e}")
                        

                        if backup_type == "unknown":
                            continue
                        

                        if source_name_from_marker:
                            original_name = source_name_from_marker
                    else:
                                         

                        backup_type = "file"
                    

                    key = f"{backup_type}|{original_name}"
                    

                    if key not in backup_groups:
                        backup_groups[key] = []
                    

                    backup_groups[key].append((ctime, item_path))
            

            for key, backups in backup_groups.items():
                              

                backups.sort(reverse=True, key=lambda x: x[0])
                

                kept_backups = []
                normal_backups = []
                

                for ctime, item_path in backups:
                                  

                    keep_file_path = os.path.join(item_path, ".keep") if os.path.isdir(item_path) else None
                    

                    if keep_file_path and os.path.exists(keep_file_path):
                        kept_backups.append((ctime, item_path))
                    else:
                        normal_backups.append((ctime, item_path))
                

                total_backups = len(backups)
                max_backups = self.max_backups
                kept_count = len(kept_backups)
                

                if total_backups > max_backups + kept_count:
                                   

                    deletable_count = total_backups - (max_backups + kept_count)
                    

                    if deletable_count > 0 and len(normal_backups) > deletable_count:
                                          

                        normal_backups.sort(key=lambda x: x[0])
                        

                        for _, item_path in normal_backups[:deletable_count]:
                            if os.path.exists(item_path):
                                self._remove_item_with_readonly_handling(item_path)
                                logging.info(f"已删除旧备份: {os.path.basename(item_path)}")
                        

                        normal_backups = normal_backups[deletable_count:]
                

                remaining_backups = kept_backups + normal_backups
                

                remaining_backups.sort(key=lambda x: x[0])
                

                for i, (ctime, item_path) in enumerate(remaining_backups):
                                  

                    new_number = i + 1
                    

                    backup_name = os.path.basename(item_path)
                    base_name = backup_name
                    match = re.match(r'(.+)_\d+$', backup_name)
                    if match:
                        base_name = match.group(1)
                    

                    new_backup_name = f"{base_name}_{new_number}"
                    

                    if new_backup_name != backup_name:
                        new_path = os.path.join(os.path.dirname(item_path), new_backup_name)
                        

                        if os.path.exists(item_path) and not os.path.exists(new_path):
                            try:
                                os.rename(item_path, new_path)
                                logging.info(self.lang_strings['renaming_backup'].format(old_name=backup_name, new_name=new_backup_name))
                            except Exception as e:
                                logging.error(self.lang_strings['rename_backup_failed'].format(error=e))
            

                self.update_source_path_file()
                

        except Exception as e:
            logging.error(self.lang_strings['error_cleanup_old_backup_item'].format(error=e))
    

    def update_backup_list(self):
        

        for item in self.backup_list.get_children():
            self.backup_list.delete(item)
        

        if not self.dest_dir or not os.path.exists(self.dest_dir):
            return
        

        try:
            backup_items = []
            

            for item in os.listdir(self.dest_dir):
                item_path = os.path.join(self.dest_dir, item)
                

                if os.path.isdir(item_path):
                                 

                    if item.startswith('.'):
                        continue
                        

                    has_file_marker = os.path.exists(os.path.join(item_path, ".file"))
                    has_folder_marker = os.path.exists(os.path.join(item_path, ".folder"))
                    

                    if not (has_file_marker or has_folder_marker):
                        continue
                        

                    ctime = os.path.getctime(item_path)
                    timestamp = datetime.datetime.fromtimestamp(ctime).strftime("%Y-%m-%d %H:%M:%S")
                    

                    keep_file_path = os.path.join(item_path, ".keep")
                    keep_status = 1 if os.path.exists(keep_file_path) else 0
                    

                    if keep_status == 0:
                        for subitem in os.listdir(item_path):
                            if subitem.startswith(".keep_"):
                                keep_status = 1
                                break
                    

                    backup_items.append((ctime, timestamp, item_path, item, keep_status))          
            

            if hasattr(self, 'sort_type'):
                if self.sort_type == "name":
                           

                    if hasattr(self, 'name_sort_order') and self.name_sort_order == "asc":
                                   

                        backup_items.sort(key=lambda x: x[3].lower())
                    else:
                                     

                        backup_items.sort(key=lambda x: x[3].lower(), reverse=True)
                elif self.sort_type == "keep":
                             

                    if hasattr(self, 'keep_sort_order') and self.keep_sort_order == "asc":
                                     

                        backup_items.sort(key=lambda x: x[4])
                    else:
                                       

                        backup_items.sort(key=lambda x: x[4], reverse=True)
                else:           
                           

                    if hasattr(self, 'sort_order') and self.sort_order == "asc":
                                     

                        backup_items.sort(key=lambda x: x[0])
                    else:
                                       

                        backup_items.sort(key=lambda x: x[0], reverse=True)
            else:
                         

                if hasattr(self, 'sort_order') and self.sort_order == "asc":
                                 

                    backup_items.sort(key=lambda x: x[0])
                else:
                                   

                    backup_items.sort(key=lambda x: x[0], reverse=True)
            

            for _, timestamp, path, folder_name, keep_status in backup_items:
                              

                keep_display = "✓" if keep_status == 1 else ""
                self.backup_list.insert("", "end", values=(folder_name, timestamp, keep_display))
        except Exception as e:
            logging.error(self.lang_strings['error_update_backup_list'].format(error=e))
    

    def delete_selected(self):
        

        selected = self.backup_list.selection()
        if not selected:
            messagebox.showinfo("Info", self.lang_strings['select_backup'])
            return
        

        try:
                                

            deleted_groups = set()
            

            for item in selected:
                values = self.backup_list.item(item, "values")
                if values and len(values) >= 2:
                                              

                    backup_name = values[0]            
                    backup_path = os.path.join(self.dest_dir, backup_name)          
                    

                    base_name = backup_name
                    match = re.match(r'(.+)_\d+$', backup_name)
                    if match:
                        base_name = match.group(1)
                                     

                        deleted_groups.add(base_name)
                    

                    if os.path.exists(backup_path):
                        self._remove_item_with_readonly_handling(backup_path)
                        logging.info(f"已删除备份: {os.path.basename(backup_path)}")
                        self.status_var.set(self.lang_strings['deleted_backup'].format(name=os.path.basename(backup_path)))
            

            self.update_source_path_file()
            

            for base_name in deleted_groups:
                self.reassign_backup_numbers(base_name)
            

            self.update_backup_list()
        except Exception as e:
            logging.error(self.lang_strings['error_delete'].format(error=e))
            messagebox.showerror("Error", self.lang_strings['error_delete'].format(error=str(e)))
    

    def restore_selected(self, skip_confirm=False, skip_success_dialog=False):
        

        selected = self.backup_list.selection()
        if not selected:
            messagebox.showinfo("Info", self.lang_strings['select_backup'])
            return
        

        if len(selected) > 1:
            messagebox.showinfo("Info", self.lang_strings['select_one_backup'])
            return
        

        item = selected[0]
        values = self.backup_list.item(item, "values")
        if not values or len(values) < 2:
            messagebox.showerror("Error", self.lang_strings['error_backup_content'])
            return
        

        backup_name = values[0]            
        backup_path = os.path.join(self.dest_dir, backup_name)          
        

        if not os.path.exists(backup_path):
            messagebox.showerror("Error", self.lang_strings['error_backup_content'])
            return
        

        original_source_path = ""
        original_name = ""
        

        if os.path.isdir(backup_path):
                                      

            file_marker_path = os.path.join(backup_path, ".file")
            folder_marker_path = os.path.join(backup_path, ".folder")
            

            marker_path = None
            if os.path.exists(file_marker_path):
                marker_path = file_marker_path
            elif os.path.exists(folder_marker_path):
                marker_path = folder_marker_path
            

            if marker_path:
                try:
                    with open(marker_path, 'r', encoding='utf-8') as f:
                        for line in f:
                            line = line.strip()
                            if line.startswith("source_name="):
                                original_name = line[12:]                      
                            elif line.startswith("source_path="):
                                original_source_path = line[12:]                      
                except Exception as e:
                    logging.error(f"读取标记文件失败: {e}")
            

            if not original_name:
                for item in os.listdir(backup_path):
                    if item not in [".file", ".folder"]:
                        original_name = item
                        break
        

        if not original_name:
            original_name = backup_name
        

        if not skip_confirm:
            if not messagebox.askyesno("Confirm", self.lang_strings['confirm_restore'].format(name=original_name)):
                return
        

        try:
                      

            restore_path = ""
            

            if original_source_path and os.path.exists(os.path.dirname(original_source_path)):
                restore_path = original_source_path
            else:
                                              

                for source_path in self.source_paths:
                    if os.path.basename(source_path) == original_name:
                        restore_path = source_path
                        break
                

                if not restore_path and original_source_path:
                    parent_dir = os.path.dirname(original_source_path)
                    if parent_dir and os.path.exists(parent_dir):
                        restore_path = os.path.join(parent_dir, original_name)
            

            if not restore_path:
                messagebox.showerror("Error", self.lang_strings['error_no_source_path'].format(name=original_name))
                return
            

            keep_file_path = os.path.join(backup_path, ".keep")
                                

            if os.path.isdir(backup_path):
                                    

                file_marker_path = os.path.join(backup_path, ".file")
                folder_marker_path = os.path.join(backup_path, ".folder")
                

                if os.path.exists(file_marker_path):
                                                  

                    backup_items = [item for item in os.listdir(backup_path) if item not in [".file", ".folder"]]
                    

                    if backup_items:
                                            

                        if len(backup_items) == 1:
                            source_item = os.path.join(backup_path, backup_items[0])
                                      

                            os.makedirs(os.path.dirname(restore_path), exist_ok=True)
                                            

                            if os.path.exists(restore_path):
                                self._remove_item_with_readonly_handling(restore_path)
                            self.copy_with_attributes(source_item, restore_path, is_restore=True)
                        else:
                                                                     

                            target_dir = os.path.dirname(restore_path)
                            os.makedirs(target_dir, exist_ok=True)
                            for item in backup_items:
                                                                 

                                if item.startswith(".") and (item.endswith("keep") or item.endswith("file") or item in [".keep", ".file", ".source_path"]):
                                    continue
                                source_item = os.path.join(backup_path, item)
                                dest_item = os.path.join(target_dir, item)
                                                

                                if os.path.exists(dest_item):
                                    self._remove_item_with_readonly_handling(dest_item)
                                self.copy_with_attributes(source_item, dest_item, is_restore=True)
                elif os.path.exists(folder_marker_path):
                                                   

                    backup_items = [item for item in os.listdir(backup_path) if item not in [".file", ".folder", ".source_path"]]
                    

                    if backup_items:
                                      

                        parent_dir = os.path.dirname(restore_path)
                        os.makedirs(parent_dir, exist_ok=True)
                        

                        if os.path.exists(restore_path):
                            self._remove_item_with_readonly_handling(restore_path)
                        

                        actual_folders = [item for item in backup_items if not (item.startswith(".") and (item.endswith("keep") or item.endswith("file") or item in [".keep", ".file", ".source_path"]))]
                        

                        if actual_folders:
                                             

                            source_folder = actual_folders[0]
                            source_folder_path = os.path.join(backup_path, source_folder)
                            

                            if os.path.exists(restore_path):
                                self._remove_item_with_readonly_handling(restore_path)
                            os.makedirs(restore_path, exist_ok=True)
                            

                            items_to_copy = []
                            for item in os.listdir(source_folder_path):
                                                                 

                                if item.startswith(".") and (item.endswith("keep") or item.endswith("file") or item in [".keep", ".file", ".source_path"]):
                                    continue
                                source_item = os.path.join(source_folder_path, item)
                                dest_item = os.path.join(restore_path, item)
                                items_to_copy.append((source_item, dest_item))
                            

                            for source_item, dest_item in items_to_copy:
                                try:
                                    self.copy_with_attributes(source_item, dest_item, is_source_item=False, is_restore=True)
                                except Exception as e:
                                    logging.error(f"复制失败: {source_item} -> {dest_item}, 错误: {e}")
                                    raise
                            

                        else:
                            logging.warning(f"未找到实际文件夹用于还原，备份项: {backup_items}")
                else:
                                       

                    backup_items = [item for item in os.listdir(backup_path) if item not in [".file", ".folder", ".source_path"]]
                    

                    if backup_items:
                                      

                        parent_dir = os.path.dirname(restore_path)
                        os.makedirs(parent_dir, exist_ok=True)
                        

                        if os.path.exists(restore_path):
                            self._remove_item_with_readonly_handling(restore_path)
                        

                        actual_folders = [item for item in backup_items if not (item.startswith(".") and (item.endswith("keep") or item.endswith("file") or item in [".keep", ".file", ".source_path"]))]
                        

                        if actual_folders:
                                             

                            source_folder = actual_folders[0]
                            source_folder_path = os.path.join(backup_path, source_folder)
                            

                            if os.path.exists(restore_path):
                                self._remove_item_with_readonly_handling(restore_path)
                            os.makedirs(restore_path, exist_ok=True)
                            

                            items_to_copy = []
                            for item in os.listdir(source_folder_path):
                                                                 

                                if item.startswith(".") and (item.endswith("keep") or item.endswith("file") or item in [".keep", ".file", ".source_path"]):
                                    continue
                                source_item = os.path.join(source_folder_path, item)
                                dest_item = os.path.join(restore_path, item)
                                items_to_copy.append((source_item, dest_item))
                            

                            for source_item, dest_item in items_to_copy:
                                try:
                                    self.copy_with_attributes(source_item, dest_item, is_source_item=False, is_restore=True)
                                except Exception as e:
                                    logging.error(f"复制失败: {source_item} -> {dest_item}, 错误: {e}")
                                    raise
                            

                        else:
                            logging.warning(f"无标记文件情况下未找到实际文件夹用于还原，备份项: {backup_items}")
                

            else:
                                 

                self.copy_with_attributes(backup_path, restore_path, is_restore=True)
                

            self.status_var.set(self.lang_strings['restored_backup'].format(name=original_name))
                                   

            if not skip_success_dialog:
                messagebox.showinfo("Success", self.lang_strings['restore_success'])
        except Exception as e:
            logging.error(self.lang_strings['error_restore'].format(error=e))
            messagebox.showerror("Error", self.lang_strings['error_restore'].format(error=str(e)))
    

    def rename_selected_backup(self, backup_name=None):
        

        if backup_name is None:
            selected = self.backup_list.selection()
            if not selected:
                messagebox.showinfo("Info", self.lang_strings['select_backup'])
                return
            

            if len(selected) > 1:
                messagebox.showinfo("Info", self.lang_strings['select_one_backup'])
                return
            

            item = selected[0]
            values = self.backup_list.item(item, "values")
            if not values or len(values) < 2:
                messagebox.showerror("Error", self.lang_strings['error_backup_content'])
                return
            

            backup_name = values[0]
        

        backup_path = os.path.join(self.dest_dir, backup_name)
        

        if not os.path.exists(backup_path):
            messagebox.showerror("Error", self.lang_strings['error_backup_content'])
            return
        

        current_name = os.path.basename(backup_path)
        

        base_font_size = int(FONT_CONFIG['base_size'] * FONT_CONFIG['scale_factor'])
        entry_font_size = max(base_font_size, 16)            
        

        dialog = tk.Toplevel(self.root)
        dialog.title(self.lang_strings['rename_backup'])
        dialog.geometry("400x200")
        dialog.resizable(True, True)            
        dialog.transient(self.root)               
        dialog.grab_set()         
        

        default_font = ('Microsoft YaHei', base_font_size)
        entry_font = ('Microsoft YaHei', entry_font_size)
        

        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill="both", expand=True)
        

        label = ttk.Label(frame, text=self.lang_strings['new_backup_name'], font=default_font)
        label.pack(pady=10)
        

        entry = ttk.Entry(frame, font=entry_font)
        entry.insert(0, current_name)
        entry.pack(pady=10, fill="x")
        entry.select_range(0, tk.END)        
        entry.focus()        
        

        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x")
        

        result = [None]                     
        

        def on_ok():
            new_name = entry.get().strip()
            if not new_name or new_name == current_name:
                dialog.destroy()
                return
            result[0] = new_name
            dialog.destroy()
        

        ok_button = ttk.Button(button_frame, text=self.lang_strings['ok_button'], command=on_ok)
        ok_button.pack(side="left", padx=5)
        

        def on_cancel():
            dialog.destroy()
        

        cancel_button = ttk.Button(button_frame, text=self.lang_strings['cancel_button'], command=on_cancel)
        cancel_button.pack(side="left", padx=5)
        

        entry.bind('<Return>', lambda event: on_ok())
        

        self.root.wait_window(dialog)
        

        new_name = result[0]
        if new_name is None:
            return
        

        try:
                        

            new_path = os.path.join(os.path.dirname(backup_path), new_name)
            if os.path.exists(new_path):
                if not messagebox.askyesno("Confirm", self.lang_strings['confirm_overwrite']):
                    return
                

                self._remove_item_with_readonly_handling(new_path)
            

            os.rename(backup_path, new_path)
            

            if os.path.isdir(new_path):
                keep_file_path = os.path.join(new_path, ".keep")
                if not os.path.exists(keep_file_path):
                    with open(keep_file_path, 'w') as f:
                        f.write("")
                    logging.info(f"已为重命名的备份创建保留标记: {new_name}")
                

            self.update_source_path_file()
            

            self.status_var.set(self.lang_strings['rename_success'])
            self.update_backup_list()
        except Exception as e:
            logging.error(self.lang_strings['error_rename'].format(error=e))
            messagebox.showerror("Error", self.lang_strings['error_rename'].format(error=str(e)))
    

    def on_backup_list_click(self, event):
        

        item_id = self.backup_list.identify_row(event.y)
        column = self.backup_list.identify_column(event.x)
        region = self.backup_list.identify_region(event.x, event.y)
        

        if region == "heading":
                                         

            if column == "#2":
                             

                if not hasattr(self, 'sort_order'):
                    self.sort_order = "desc"               
                else:
                            

                    self.sort_order = "asc" if self.sort_order == "desc" else "desc"
                

                self.sort_type = "time"
                

                self.update_backup_list()
                return
            

            if column == "#1":
                             

                if not hasattr(self, 'name_sort_order'):
                    self.name_sort_order = "asc"             
                else:
                            

                    self.name_sort_order = "desc" if self.name_sort_order == "asc" else "asc"
                

                self.sort_type = "name"
                

                self.update_backup_list()
                return
            

            if column == "#3":
                             

                if not hasattr(self, 'keep_sort_order'):
                    self.keep_sort_order = "asc"               
                else:
                            

                    self.keep_sort_order = "desc" if self.keep_sort_order == "asc" else "asc"
                

                self.sort_type = "keep"
                

                self.update_backup_list()
                return
        

        if not item_id:
            return
        

        if column == "#3":
            try:
                        

                values = self.backup_list.item(item_id, "values")
                if not values or len(values) < 3:
                    return
                

                backup_name = values[0]
                backup_path = os.path.join(self.dest_dir, backup_name)
                

                if not os.path.exists(backup_path):
                    return
                

                keep_file_path = os.path.join(backup_path, ".keep")
                

                if os.path.exists(keep_file_path):
                                    

                    os.remove(keep_file_path)
                    new_keep_status = ""
                    logging.info(f"已移除备份的保留状态: {backup_name}")
                else:
                                     

                    with open(keep_file_path, 'w') as f:
                        f.write("")
                    new_keep_status = "✓"
                    logging.info(f"已设置备份为保留状态: {backup_name}")
                

                self.backup_list.item(item_id, values=(values[0], values[1], new_keep_status))
            except Exception as e:
                logging.error(f"切换备份保留状态时出错: {e}")
    

    def on_backup_list_double_click(self, event):
        

        item_id = self.backup_list.identify_row(event.y)
        column = self.backup_list.identify_column(event.x)
        

        if not item_id:
            return
        

        if column == "#1":
            try:
                        

                values = self.backup_list.item(item_id, "values")
                if not values or len(values) < 1:
                    return
                

                backup_name = values[0]
                backup_path = os.path.join(self.dest_dir, backup_name)
                

                if not os.path.exists(backup_path):
                    return
                

                self.rename_selected_backup(backup_name)
            except Exception as e:
                logging.error(f"双击备份名称时出错: {e}")
    

    def save_advanced_settings(self):
        

        try:


            if self.suffix_type_var.get() == "custom" and not self.custom_suffix_entry.get().strip():
                messagebox.showerror("Error", self.lang_strings['error_suffix_empty'])
                return
            

            try:
                interval = int(self.interval_entry.get())
                if interval <= 0:
                    messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                    return
                self.interval = interval
            except ValueError:
                messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                return
            

            try:
                max_backups = int(self.max_backups_entry.get())
                if max_backups <= 0:
                    messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                    return
                self.max_backups = max_backups
            except ValueError:
                messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                return
            

            self.suffix_type = self.suffix_type_var.get()
            self.custom_suffix = self.custom_suffix_entry.get().strip()
            self.duplicate_handling = self.duplicates_var.get()
            

            if hasattr(self, 'auto_number_mode_var'):
                self.auto_number_mode = self.auto_number_mode_var.get()
            

            if hasattr(self, 'skip_hidden_var'):
                self.skip_hidden = self.skip_hidden_var.get()
            

            self.save_config()
            

            self.update_ui_texts()
            

            self.update_source_list()
            self.update_backup_list()
            

            if self.running:
                self.status_var.set(self.lang_strings['backup_started'].format(interval=self.interval))
            

            messagebox.showinfo("Info", self.lang_strings['save_settings'])
        except Exception as e:
            logging.error(self.lang_strings['error_save_advanced_settings'].format(error=e))
            messagebox.showerror("Error", self.lang_strings['error_save_settings'])
    

    def restore_default_settings(self):
        

        try:
                   

            result = messagebox.askyesno(

                self.lang_strings.get('confirm', 'Confirm'),

                self.lang_strings.get('confirm_restore_defaults', 'Are you sure you want to restore all settings to defaults? This action cannot be undone.')

            )
            

            if not result:
                return           
            

            self.interval = 300
            self.max_backups = 5
            self.suffix_type = "number"
            self.custom_suffix = ""
            self.duplicate_handling = "rename"
            self.backup_counter = 1
            self.start_number = 1
            self.auto_number_mode = True
            self.skip_hidden = False
            self.hotkey = "ctrl+F1"
            self.restore_hotkey = "ctrl+F2"
            

            self.interval_entry.delete(0, tk.END)
            self.interval_entry.insert(0, str(self.interval))
            

            self.max_backups_entry.delete(0, tk.END)
            self.max_backups_entry.insert(0, str(self.max_backups))
            

            self.suffix_type_var.set(self.suffix_type)
            

            self.custom_suffix_entry.delete(0, tk.END)
            self.custom_suffix_entry.insert(0, self.custom_suffix)
            

            self.duplicates_var.set(self.duplicate_handling)
            

            if hasattr(self, 'auto_number_mode_var'):
                self.auto_number_mode_var.set(self.auto_number_mode)
            

            if hasattr(self, 'skip_hidden_var'):
                self.skip_hidden_var.set(self.skip_hidden)
            

            self.hotkey_entry.delete(0, tk.END)
            self.hotkey_entry.insert(0, self.hotkey)
            

            self.hotkey_var.set(False)
            

            self.restore_hotkey_entry.delete(0, tk.END)
            self.restore_hotkey_entry.insert(0, self.restore_hotkey)
            

            self.restore_hotkey_var.set(False)
            

            if hasattr(self, 'counter_var'):
                self.counter_var.set(str(self.backup_counter))
            

            self.update_suffix_entry_state()
            

            self.update_number_mode_visibility()
            

            self.save_config()
            

            self.update_ui_texts()
            

            self.update_source_list()
            self.update_backup_list()
            

            if self.running:
                self.status_var.set(self.lang_strings['backup_started'].format(interval=self.interval))
            

            messagebox.showinfo(

                self.lang_strings.get('info', 'Info'),

                self.lang_strings.get('settings_restored_success', 'All settings have been restored to defaults.')

            )
        except Exception as e:
            logging.error(self.lang_strings['error_restore_defaults'].format(error=e))
            messagebox.showerror(

                self.lang_strings.get('error', 'Error'),

                f"{self.lang_strings.get('restore_defaults_failed', 'Failed to restore defaults')}: {str(e)}"

            )
    

    def _enable_hotkeys(self):
        

        try:
            if self.hotkey_var.get():
                self.toggle_hotkey()
            

            if self.restore_hotkey_var.get():
                self.toggle_restore_hotkey()
        except Exception as e:
            logging.error(self.lang_strings['error_enable_hotkeys'].format(error=e))
    

    def reassign_backup_numbers(self, base_name):
        

        try:
                           

            backup_items = []                             
            

            for item in os.listdir(self.dest_dir):
                if item == ".source_path":                      
                    continue
                    

                item_path = os.path.join(self.dest_dir, item)
                if os.path.isfile(item_path) or os.path.isdir(item_path):
                                

                    match = re.match(rf"^{re.escape(base_name)}_\d+$", item)
                    if match:
                                

                        ctime = os.path.getctime(item_path)
                        backup_items.append((ctime, item_path, item))
            

            backup_items.sort(key=lambda x: x[0])
            

            for i, (ctime, item_path, item_name) in enumerate(backup_items):
                              

                new_number = i + 1
                

                new_backup_name = f"{base_name}_{new_number}"
                

                if new_backup_name != item_name:
                                               

                    new_path = os.path.join(self.dest_dir, new_backup_name)
                    

                    if os.path.exists(item_path) and not os.path.exists(new_path):
                        try:
                            os.rename(item_path, new_path)
                            logging.info(self.lang_strings['renaming_backup'].format(old_name=item_name, new_name=new_backup_name))
                        except Exception as e:
                            logging.error(self.lang_strings['rename_backup_failed'].format(error=e))
            

            self.update_source_path_file()
            

        except Exception as e:
            logging.error(self.lang_strings['error_reassign_backup_numbers'].format(error=e))
    

    def set_start_number(self):
        

        try:
                           

            base_font_size = int(FONT_CONFIG['base_size'] * FONT_CONFIG['scale_factor'])
            entry_font_size = max(base_font_size, 16)            
            

            dialog = tk.Toplevel(self.root)
            dialog.title(self.lang_strings['set_start_number'])
            dialog.geometry("400x200")
            dialog.resizable(True, True)            
            dialog.transient(self.root)               
            dialog.grab_set()         
            

            default_font = ('Microsoft YaHei', base_font_size)
            entry_font = ('Microsoft YaHei', entry_font_size)
            

            frame = ttk.Frame(dialog, padding="20")
            frame.pack(fill="both", expand=True)
            

            label = ttk.Label(frame, text=self.lang_strings['confirm_set_start_number'], font=default_font)
            label.pack(pady=10)
            

            entry = ttk.Entry(frame, font=entry_font)
            entry.insert(0, str(self.start_number))
            entry.pack(pady=10, fill="x")
            entry.select_range(0, tk.END)        
            entry.focus()        
            

            button_frame = ttk.Frame(frame)
            button_frame.pack(fill="x")
            

            result = [None]                     
            

            def on_ok():
                try:
                    new_number = int(entry.get())
                    if new_number < 1:
                        messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                        return
                    result[0] = new_number
                    dialog.destroy()
                except ValueError:
                    messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
            

            ok_button = ttk.Button(button_frame, text=self.lang_strings['ok_button'], command=on_ok)
            ok_button.pack(side="right", padx=5)
            

            def on_cancel():
                dialog.destroy()
            

            cancel_button = ttk.Button(button_frame, text=self.lang_strings['cancel_button'], command=on_cancel)
            cancel_button.pack(side="right", padx=5)
            

            entry.bind('<Return>', lambda event: on_ok())
            

            self.root.wait_window(dialog)
            

            new_number = result[0]
            if new_number is None:
                return
            

            self.start_number = new_number
            self.backup_counter = new_number
            

            self.counter_var.set(str(self.backup_counter))
            

            self.save_config()
            

            messagebox.showinfo("Info", self.lang_strings['start_number_updated'])
        except Exception as e:
            logging.error(self.lang_strings['error_set_start_number'].format(error=e))
            messagebox.showerror("Error", self.lang_strings['error_set_start_number'])


def main():
                                   

    try:
        import ctypes
        import win32gui
        import win32con
        

        app_id = "Local.Auto.Backup.Assistant"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
        logging.info(f"设置应用程序ID成功: {app_id}")
        

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        

        logging.info("COM库初始化已临时禁用，用于测试文件对话框卡死问题")
    except Exception as e:
        logging.warning(f"初始化Windows环境失败: {e}")
    

    root = tk.Tk()
    app = AutoBackupTool(root)
    

    root.after(100, lambda: app.set_icon())
    

    def after_window_shown():
        app.set_icon()
                 

        try:
            import ctypes
            import win32con
            hwnd = app.root.winfo_id()
            if hwnd:
                ctypes.windll.user32.SendMessageW(hwnd, win32con.WM_SETTINGCHANGE, 0, 0)
                logging.info("发送WM_SETTINGCHANGE消息刷新任务栏")
        except Exception as e:
            logging.warning(f"刷新任务栏失败: {e}")
    

    root.after(500, after_window_shown)
    

    root.mainloop()


if __name__ == "__main__":
    main()
