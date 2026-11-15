"""Windows 右键菜单管理器 (Context Menu Manager).

通过修改 HKCU 注册表，在无需管理员权限的情况下安装 Windows 原生二级级联右键菜单：
1. Excel 文件右键 (.xlsx / .xlsm / .xls):
   [ExcelRelay] 投研资产引擎 ▶
     ├─ [📋 注入“_excel_relay”元数据模板]  -> 执行 `excel-relay pack init "%1"`
     ├─ [📂 注册为项目专属数据源]   -> 执行 `excel-relay add "%1" --project`
     └─ [🌐 注册为用户全局数据源]   -> 执行 `excel-relay add "%1" --global`

2. 文件夹图标及文件夹内部空白处右键 (Directory / Directory\\Background):
   [ExcelRelay] 投研资产引擎 ▶
     ├─ [🚀 初始化为项目资产空间]   -> 执行 `excel-relay init "%1"` (或 `"%V"`)
     ├─ [📂 批量注册为项目数据源]   -> 执行 `excel-relay add "%1" --project` (或 `"%V"`)
     └─ [🌐 批量注册为全局数据源]   -> 执行 `excel-relay add "%1" --global` (或 `"%V"`)
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import List, Tuple

if sys.platform == "win32":
    import winreg
else:
    winreg = None

SUPPORTED_EXTS = [".xlsx", ".xlsm", ".xls"]

DIR_TARGETS = [
    ("文件夹图标 (Directory)", "Software\\Classes\\Directory\\shell", "%1"),
    ("文件夹空白处 (Background)", "Software\\Classes\\Directory\\Background\\shell", "%V"),
]

MAIN_MENU_KEY = "ExcelRelay.Main"
MAIN_MENU_LABEL = "[ExcelRelay] 投研资产引擎"
MAIN_MENU_ICON = "shell32.dll,260"

SUB_MENU_ITEMS = [
    {
        "key_name": "01_PackInit",
        "label": "[📋 注入“_excel_relay”元数据模板]",
        "command_args": 'pack init "%1"',
        "icon": "shell32.dll,134",
    },
    {
        "key_name": "02_RegisterProject",
        "label": "[📂 注册为项目专属数据源]",
        "command_args": 'add "%1" --project',
        "icon": "shell32.dll,4",
    },
    {
        "key_name": "03_RegisterGlobal",
        "label": "[🌐 注册为用户全局数据源]",
        "command_args": 'add "%1" --global',
        "icon": "shell32.dll,15",
    },
]

DIR_SUB_MENU_ITEMS = [
    {
        "key_name": "01_InitProject",
        "label": "[🚀 初始化为项目资产空间]",
        "command_template": 'init "{placeholder}"',
        "icon": "shell32.dll,43",
    },
    {
        "key_name": "02_BatchRegisterProject",
        "label": "[📂 批量注册为项目数据源]",
        "command_template": 'add "{placeholder}" --project',
        "icon": "shell32.dll,4",
    },
    {
        "key_name": "03_BatchRegisterGlobal",
        "label": "[🌐 批量注册为全局数据源]",
        "command_template": 'add "{placeholder}" --global',
        "icon": "shell32.dll,15",
    },
]


def get_command_string(subcommand: str) -> str:
    """获取右键点击时执行的完整 Shell 命令."""
    scripts_dir = Path(sys.executable).parent
    excel_relay_cmd = scripts_dir / "excel-relay.cmd"
    if not excel_relay_cmd.exists():
        excel_relay_cmd = scripts_dir / "excel-relay.bat"

    if excel_relay_cmd.exists():
        invoke_cmd = f'"{excel_relay_cmd}" {subcommand}'
    else:
        invoke_cmd = f'"{sys.executable}" -m excel_relay.interface.cli {subcommand}'

    return f'cmd.exe /c "chcp 65001 >nul & set PYTHONIOENCODING=utf-8 & {invoke_cmd} & echo. & pause"'


def _delete_reg_key_tree(root, path: str) -> None:
    """递归删除注册表键树."""
    if winreg is None:
        return
    try:
        with winreg.OpenKey(root, path, 0, winreg.KEY_ALL_ACCESS) as key:
            while True:
                try:
                    sub = winreg.EnumKey(key, 0)
                    _delete_reg_key_tree(root, f"{path}\\{sub}")
                except OSError:
                    break
        winreg.DeleteKey(root, path)
    except FileNotFoundError:
        pass
    except Exception:
        pass


def _install_cascading_menu_at(base_shell: str, items: List[dict]) -> None:
    """在指定注册表 shell 路径下安装 ExcelRelay 二级级联菜单."""
    if winreg is None:
        return
    _delete_reg_key_tree(winreg.HKEY_CURRENT_USER, f"{base_shell}\\{MAIN_MENU_KEY}")

    main_path = f"{base_shell}\\{MAIN_MENU_KEY}"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, main_path) as main_key:
        winreg.SetValueEx(main_key, "MUIVerb", 0, winreg.REG_SZ, MAIN_MENU_LABEL)
        winreg.SetValueEx(main_key, "SubCommands", 0, winreg.REG_SZ, "")
        winreg.SetValueEx(main_key, "Icon", 0, winreg.REG_SZ, MAIN_MENU_ICON)

    sub_shell_path = f"{main_path}\\shell"
    for item in items:
        item_path = f"{sub_shell_path}\\{item['key_name']}"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, item_path) as item_key:
            winreg.SetValueEx(item_key, "MUIVerb", 0, winreg.REG_SZ, item["label"])
            winreg.SetValueEx(item_key, "Icon", 0, winreg.REG_SZ, item["icon"])

        cmd_path = f"{item_path}\\command"
        cmd_str = get_command_string(item["command_args"])
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, cmd_path) as cmd_key:
            winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, cmd_str)


def install_context_menu() -> List[str]:
    """为所有支持的 Excel 后缀以及文件夹（图标与空白处）安装二级级联右键菜单."""
    if sys.platform != "win32" or winreg is None:
        raise OSError("右键菜单功能仅支持 Windows 系统。")

    installed = []

    # 1. 安装 Excel 文件右键菜单
    for ext in SUPPORTED_EXTS:
        try:
            base_shell = f"Software\\Classes\\SystemFileAssociations\\{ext}\\shell"
            _install_cascading_menu_at(base_shell, SUB_MENU_ITEMS)
            installed.append(ext)
        except Exception as e:
            raise RuntimeError(f"安装右键菜单到 {ext} 失败: {e}") from e

    # 2. 安装文件夹图标与文件夹内部空白处右键菜单
    for target_label, base_shell, placeholder in DIR_TARGETS:
        try:
            resolved_items = [
                {
                    "key_name": item["key_name"],
                    "label": item["label"],
                    "command_args": item["command_template"].format(placeholder=placeholder),
                    "icon": item["icon"],
                }
                for item in DIR_SUB_MENU_ITEMS
            ]
            _install_cascading_menu_at(base_shell, resolved_items)
            installed.append(target_label)
        except Exception as e:
            raise RuntimeError(f"安装右键菜单到 {target_label} 失败: {e}") from e

    return installed


def uninstall_context_menu() -> List[str]:
    """卸载并清理所有 ExcelRelay 右键菜单项 (包括 Excel 文件与文件夹级联菜单)."""
    if sys.platform != "win32" or winreg is None:
        raise OSError("右键菜单功能仅支持 Windows 系统。")

    uninstalled = []

    for ext in SUPPORTED_EXTS:
        base_shell = f"Software\\Classes\\SystemFileAssociations\\{ext}\\shell"
        try:
            _delete_reg_key_tree(winreg.HKEY_CURRENT_USER, f"{base_shell}\\{MAIN_MENU_KEY}")
            uninstalled.append(ext)
        except Exception as e:
            raise RuntimeError(f"从 {ext} 卸载右键菜单失败: {e}") from e

    for target_label, base_shell, _placeholder in DIR_TARGETS:
        try:
            _delete_reg_key_tree(winreg.HKEY_CURRENT_USER, f"{base_shell}\\{MAIN_MENU_KEY}")
            uninstalled.append(target_label)
        except Exception as e:
            raise RuntimeError(f"从 {target_label} 卸载右键菜单失败: {e}") from e

    return uninstalled


def get_menu_status() -> List[Tuple[str, bool, List[str]]]:
    """查询当前各后缀及文件夹的右键菜单安装状态.

    Returns:
        List[(target_name, all_enabled, active_labels)]
    """
    if sys.platform != "win32" or winreg is None:
        return [(ext, False, []) for ext in SUPPORTED_EXTS] + [
            (label, False, []) for label, _, _ in DIR_TARGETS
        ]

    status: List[Tuple[str, bool, List[str]]] = []

    for ext in SUPPORTED_EXTS:
        active_labels: List[str] = []
        base_shell = f"Software\\Classes\\SystemFileAssociations\\{ext}\\shell\\{MAIN_MENU_KEY}\\shell"
        for item in SUB_MENU_ITEMS:
            item_cmd_path = f"{base_shell}\\{item['key_name']}\\command"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, item_cmd_path):
                    active_labels.append(item["label"])
            except Exception:
                pass

        all_ok = (len(active_labels) == len(SUB_MENU_ITEMS))
        status.append((ext, all_ok, active_labels))

    for target_label, base_shell_root, _placeholder in DIR_TARGETS:
        active_labels = []
        base_shell = f"{base_shell_root}\\{MAIN_MENU_KEY}\\shell"
        for item in DIR_SUB_MENU_ITEMS:
            item_cmd_path = f"{base_shell}\\{item['key_name']}\\command"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, item_cmd_path):
                    active_labels.append(item["label"])
            except Exception:
                pass

        all_ok = (len(active_labels) == len(DIR_SUB_MENU_ITEMS))
        status.append((target_label, all_ok, active_labels))

    return status
