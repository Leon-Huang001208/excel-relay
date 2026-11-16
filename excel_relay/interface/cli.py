"""excel-relay 命令行客户端 (CLI Entrypoint).

基于 Typer 和 Rich 实现极简、专业、优雅的终端体验。
"""

from __future__ import annotations

import json
import sys
from contextlib import nullcontext
from pathlib import Path
from typing import List, Optional

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from excel_relay.core import (
    CATALOG_PROTOCOL_VERSION,
    CascadingCatalog,
    CatalogManager,
    CatalogScope,
    DescriptionOrigin,
    ExcelInspector,
    HeadlessExcelRefresher,
    RefreshError,
    find_project_root,
    get_cascading_catalog,
    get_version,
    get_version_info,
    is_sensitive_dir,
)
from excel_relay.core.contracts import configure_source_contract
from excel_relay.core.refresher import RefreshBusyError
from excel_relay.interface.menu import (
    DIR_SUB_MENU_ITEMS,
    SUB_MENU_ITEMS,
    get_menu_status,
    install_context_menu,
    uninstall_context_menu,
)

app = typer.Typer(
    name="excel-relay",
    help="Excel Relay：让 AI 发现、路由并后台刷新 Excel 模型。",
    add_completion=False,
)
pack_app = typer.Typer(name="pack", help="数据源包 _excel_relay 模板注入、校验与离线剖析工具。")
app.add_typer(pack_app)
menu_app = typer.Typer(name="menu", help="Windows 资源管理器右键快捷菜单管理。")
app.add_typer(menu_app)

if sys.platform == "win32":
    try:
        # 显式激活 Windows 控制台原生 ANSI / 虚拟终端支持 (ENABLE_VIRTUAL_TERMINAL_PROCESSING)
        import ctypes

        kernel32 = ctypes.windll.kernel32
        for std_id in (-11, -12):  # STD_OUTPUT_HANDLE, STD_ERROR_HANDLE
            h = kernel32.GetStdHandle(std_id)
            mode = ctypes.c_ulong()
            if kernel32.GetConsoleMode(h, ctypes.byref(mode)):
                kernel32.SetConsoleMode(h, mode.value | 0x0004)
    except Exception:
        pass
    try:
        import colorama

        colorama.just_fix_windows_console()
    except Exception:
        pass
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

console = Console(safe_box=False)


def _version_callback(value: bool):
    if value:
        console.print(f"excel-relay [bold green]v{get_version()}[/bold green]")
        raise typer.Exit()


@app.callback()
def main(
    version: Optional[bool] = typer.Option(
        None,
        "--version",
        "-v",
        help="显示 excel-relay 引擎版本号并退出。",
        callback=_version_callback,
        is_eager=True,
    ),
):
    """Excel Relay：让 AI 发现、路由并后台刷新 Excel 模型。"""
    pass


# -----------------------------------------------------------------------------
# 0. 版本与运行环境信息 (Version)
# -----------------------------------------------------------------------------
@app.command(name="version")
def show_version(
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出版本与环境详情。"),
):
    """查看 excel-relay 引擎版本与运行环境详细信息。"""
    info = get_version_info()
    if json_output:
        console.print(json.dumps(info, ensure_ascii=False, indent=2))
        return

    table = Table(title=f"[bold blue]excel-relay 引擎版本信息[/bold blue] (v{info['version']})")
    table.add_column("组件 / 环境", style="cyan")
    table.add_column("版本 / 状态", style="bold green")
    table.add_row("excel-relay 引擎版本", f"v{info['version']}")
    table.add_row("Catalog 仓储格式版本", f"v{info['catalog_version']}")
    table.add_row("Python 环境", info["python_version"])
    table.add_row("操作系统平台", info["platform"])
    table.add_row("系统详细标识", info["system"])
    console.print(table)


# -----------------------------------------------------------------------------
# 1. 环境诊断 (Doctor) 与空间初始化 (Init)
# -----------------------------------------------------------------------------
@app.command()
def doctor(
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出诊断详情。"),
):
    """检查 Excel、金融加载项和项目或全局数据目录。"""
    env = HeadlessExcelRefresher.probe_environment()
    cascading = get_cascading_catalog()
    is_macos = sys.platform == "darwin"

    if json_output:
        data = {
            "engine": {
                "version": get_version(),
                "catalog_protocol_version": CATALOG_PROTOCOL_VERSION,
            },
            "python": {
                "version": env["python_version"],
                "platform": env["platform"],
            },
            "excel": {
                "available": env["excel_ok"],
                "version": env.get("excel_version"),
                "error": env.get("excel_error"),
            },
            "plugins": env.get("plugins", []),
            "context": {
                "mode": "project" if cascading.fallback else "global",
                "primary_catalog": str(cascading.primary.catalog_file),
                "primary_scope": cascading.primary.store.scope.value,
                "primary_project_name": cascading.primary.store.project_name,
                "primary_sources_count": len(cascading.primary.store.sources),
                "fallback_catalog": str(cascading.fallback.catalog_file) if cascading.fallback else None,
                "fallback_sources_count": len(cascading.fallback.store.sources) if cascading.fallback else 0,
            },
        }
        if is_macos:
            data["excel"]["backend"] = "applescript"
            data["excel"]["refresh_formats"] = [".xlsx"]
            data["excel"]["refresh_vendors"] = ["wind"]
        print(json.dumps(data, ensure_ascii=False, indent=2))
        return

    table = Table(title="[bold blue]excel-relay 环境与插件诊断[/bold blue]")
    table.add_column("检查项", style="cyan")
    table.add_column("状态", style="bold")
    table.add_column("详情说明", style="dim")

    table.add_row("excel-relay 引擎", "[green][OK] 正常[/green]", f"v{get_version()} (Catalog 协议 v{CATALOG_PROTOCOL_VERSION})")

    table.add_row("Python 环境", "[green][OK] 正常[/green]", f"Python {env['python_version']} ({env['platform']})")

    if env["excel_ok"]:
        table.add_row("Microsoft Excel", "[green][OK] 已安装[/green]", f"Excel 版本号: {env['excel_version']}")
    else:
        backend_name = "AppleScript" if is_macos else "COM"
        table.add_row("Microsoft Excel", "[red][FAIL] 异常[/red]", f"无法通过 {backend_name} 调起 Excel: {env['excel_error']}")

    if env["plugins"]:
        plugin_status = "[green][OK] 已检测到[/green]" if is_macos else "[green][OK] 已激活就绪[/green]"
        table.add_row("金融数据插件", plugin_status, ", ".join(env["plugins"]))
    else:
        plugin_name = "Wind" if is_macos else "万得或同花顺"
        table.add_row("金融数据插件", "[yellow]! 未检测到[/yellow]", f"未在 Excel 运行时发现{plugin_name}插件")
    if is_macos:
        table.add_row("刷新范围", "macOS", "Excel + Wind .xlsx，生成保留公式的副本")

    # 探测当前上下文
    if cascading.fallback:
        # 说明当前位于项目内，且持有全局兜底仓
        proj_mgr = cascading.primary
        proj_count = len(proj_mgr.store.sources)
        table.add_row(
            "本项目数据源",
            "[green][OK] 已激活[/green]",
            f"{proj_mgr.catalog_file} (项目: {proj_mgr.store.project_name}, 登记 {proj_count} 个专属数据源)",
        )
        glob_mgr = cascading.fallback
        glob_count = len(glob_mgr.store.sources)
        table.add_row(
            "用户全局数据源",
            "[green][OK] 就绪[/green]",
            f"{glob_mgr.catalog_file} (登记 {glob_count} 个通用数据源)",
        )
    else:
        # 当前仅处于全局上下文
        count = len(cascading.primary.store.sources)
        table.add_row(
            "数据源注册表",
            "[green][OK] 全局模式[/green]",
            f"{cascading.primary.catalog_file} (未在项目中，已注册 {count} 个全局数据源)",
        )

    console.print(table)


@app.command(name="init")
def init_project(
    path: Optional[Path] = typer.Argument(None, help="拟初始化为 excel-relay 项目资产空间的目录路径 (默认当前工作目录)"),
    name: Optional[str] = typer.Option(None, "--name", "-n", help="项目中文名称或别名 (默认取目录名)"),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出初始化结果。"),
):
    """在指定目录初始化一个独立的 excel-relay 项目资产空间 (.excel-relay/catalog.json)."""
    try:
        mgr, created = CatalogManager.init_project(path=path, project_name=name)
    except ValueError as e:
        if json_output:
            print(json.dumps({"status": "error", "error": str(e)}, ensure_ascii=False, indent=2))
        else:
            console.print(f"[bold red]拒绝初始化:[/bold red] {e}")
            console.print("[yellow]提示: 请在具体的投研课题或工程文件夹内执行初始化。[/yellow]")
        raise typer.Exit(1)

    if json_output:
        print(
            json.dumps(
                {
                    "status": "success",
                    "created_new": created,
                    "project_name": mgr.store.project_name,
                    "project_root": str(mgr.base_dir),
                    "catalog_file": str(mgr.catalog_file),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return

    if not created:
        console.print(f"[yellow]该目录已是 excel-relay 项目空间:[/yellow] {mgr.base_dir}")
        console.print(f"资产注册表: {mgr.catalog_file}")
        return

    console.print("[bold green][OK] 成功初始化 excel-relay 项目资产空间！[/bold green]")
    console.print(f"项目名称: [bold cyan]{mgr.store.project_name}[/bold cyan]")
    console.print(f"项目根目录: {mgr.base_dir}")
    console.print(f"资产注册表: {mgr.catalog_file}")
    console.print("[dim]提示: 放置在项目内的 Excel 底稿执行 excel-relay add 将默认注册为本项目专属数据源。[/dim]")


# -----------------------------------------------------------------------------
# 2. 数据源管理 (Add / Set / Rm / Ls / Find)
# -----------------------------------------------------------------------------
@app.command(name="add")
def add_source(
    path: Path = typer.Argument(..., help="拟注册的 Excel 文件或目录路径"),
    tag: Optional[List[str]] = typer.Option(None, "--tag", "-t", help="附加标签，如 --tag 食品饮料"),
    project_flag: bool = typer.Option(False, "--project", "-p", help="强制注册为【当前项目专属】数据源"),
    global_flag: bool = typer.Option(False, "--global", "-g", help="强制注册为【用户全局】数据源"),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出注册结果。"),
):
    """注册单个 Excel 文件或递归扫描目录添加为数据源 (自动感知项目/全局)."""
    if project_flag and global_flag:
        if json_output:
            print(json.dumps({"status": "error", "error": "--project 与 --global 为互斥选项，不能同时指定。"}, ensure_ascii=False, indent=2))
        else:
            console.print("[bold red]错误: --project 与 --global 为互斥选项，不能同时指定。[/bold red]")
        raise typer.Exit(1)

    target_abs = path.resolve()
    target_folder = target_abs.parent if target_abs.is_file() else target_abs
    project_root = find_project_root(target_abs)

    if global_flag:
        target_mgr = CatalogManager()
    elif project_flag:
        if project_root is not None:
            target_mgr = CatalogManager(base_dir=project_root)
        else:
            # 目标路径及其父级均未发现 .excel-relay 空间
            is_interactive = sys.stdin.isatty() and not json_output
            if not is_interactive:
                msg = f"目标路径所在目录 '{target_folder}' 及其父级目录中未发现 excel-relay 项目资产空间 (.excel-relay)。"
                if json_output:
                    print(json.dumps({"status": "error", "error": msg}, ensure_ascii=False, indent=2))
                else:
                    console.print(
                        f"[bold red]错误: {msg}[/bold red]\n"
                        "[dim]若要作为项目使用，请先运行 `excel-relay init`；或使用 `--global` 注册为全局数据源。[/dim]"
                    )
                raise typer.Exit(1)

            # 友好控制台交互提示
            console.print(f"\n[bold yellow][提示][/bold yellow] 未在目标路径上级目录中发现 excel-relay 项目资产空间 (.excel-relay)。")
            console.print(f"目标目录: [cyan]{target_folder}[/cyan]\n")
            console.print("请选择注册目标：")
            console.print(f"  [bold]1[/bold] 在当前文件夹 ({target_folder.name}) 初始化为新项目并注册")
            console.print(f"  [bold]2[/bold] 注册为【用户全局数据源】[green](推荐)[/green]")
            console.print(f"  [bold]3[/bold] 取消操作\n")

            choice = typer.prompt("请输入选项", default="2").strip()
            if choice == "1":
                if is_sensitive_dir(target_folder):
                    confirm = typer.confirm(
                        f"[警告] 当前文件夹为系统常用公共目录 ({target_folder.name})，确定要在此处初始化项目吗？",
                        default=False,
                    )
                    if not confirm:
                        console.print("[yellow]操作已取消。[/yellow]")
                        raise typer.Exit(0)

                # 初始化新项目并创建管理器
                CatalogManager.init_project(path=target_folder, allow_sensitive=True)
                console.print(f"[green]已在 {target_folder} 成功初始化 excel-relay 项目资产空间。[/green]")
                target_mgr = CatalogManager(base_dir=target_folder)
            elif choice == "2":
                target_mgr = CatalogManager()
            else:
                console.print("[yellow]操作已取消。[/yellow]")
                raise typer.Exit(0)
    else:
        # 默认自动感知：文件在项目内则注册为项目专属，否则注册为用户全局
        target_mgr = CatalogManager(base_dir=project_root) if project_root else CatalogManager()

    scope_name = "【项目专属】" if target_mgr.store.scope == CatalogScope.PROJECT else "【用户全局】"
    if json_output:
        try:
            results = target_mgr.register_path(target_abs, custom_tags=tag)
        except Exception as e:
            print(json.dumps({"status": "error", "error": str(e)}, ensure_ascii=False, indent=2))
            raise typer.Exit(1)
        res_list = []
        for r in results:
            res_list.append({
                "source_id": r.source_id,
                "scope": r.scope.value,
                "name": r.profile.name,
                "summary": r.profile.summary,
                "metadata_status": r.profile.metadata_status.value,
                "description_origin": r.profile.description_origin.value,
                "vendor": r.physical.vendor,
                "has_live_plugins": r.physical.has_live_plugins,
                "inputs_count": len(r.inputs),
                "inputs": [i.model_dump() for i in r.inputs],
                "sheets": [s.name for s in r.physical.sheets],
                "file_path": r.file_path,
                "absolute_path": str(target_mgr.resolve_absolute_path(r)),
            })
        print(json.dumps({
            "status": "success",
            "scope": target_mgr.store.scope.value,
            "registered_count": len(results),
            "sources": res_list,
        }, ensure_ascii=False, indent=2))
        return

    with console.status(f"[bold green]正在扫描并注册到 {scope_name}: {path}..."):
        try:
            results = target_mgr.register_path(target_abs, custom_tags=tag)
        except Exception as e:
            console.print(f"[bold red]注册失败:[/bold red] {e}")
            raise typer.Exit(1)

    if not results:
        console.print(f"[yellow]未找到有效的 Excel 文件 (.xlsx/.xlsm/.xls): {path}[/yellow]")
        return

    table = Table(title=f"[bold green]成功注册/更新 {len(results)} 个数据源到 {scope_name}[/bold green]")
    table.add_column("数据源 ID (source_id)", style="cyan", no_wrap=True)
    table.add_column("模型名称", style="bold")
    table.add_column("作用域", justify="center")
    table.add_column("状态分层", style="magenta")
    table.add_column("插件", style="yellow")
    table.add_column("可调参数", justify="center")
    table.add_column("摘要说明", style="white")

    for r in results:
        status_color = {"declared": "green", "enriched": "cyan", "raw": "dim"}.get(
            r.profile.metadata_status.value, "white"
        )
        badge = f"[{status_color}]{r.profile.metadata_status.value}[/{status_color}]"
        io_str = str(len(r.inputs))
        scope_badge = "[green][项目][/green]" if r.scope == CatalogScope.PROJECT else "[cyan][全局][/cyan]"
        table.add_row(
            r.source_id,
            r.profile.name,
            scope_badge,
            badge,
            r.physical.vendor,
            io_str,
            r.profile.summary[:32] + ("..." if len(r.profile.summary) > 32 else ""),
        )

    console.print(table)


@app.command(name="ls")
def list_sources(
    tag: Optional[str] = typer.Option(None, "--tag", "-t", help="按标签过滤"),
    project_only: bool = typer.Option(False, "--project-only", "-p", help="仅列出当前项目专属数据源"),
    global_only: bool = typer.Option(False, "--global-only", "-g", help="仅列出用户全局数据源"),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出已注册数据源清单。"),
):
    """列出当前已注册的所有 Excel 数据源 (支持项目与全局分层展示)."""
    if project_only and global_only:
        if json_output:
            print(json.dumps({"status": "error", "error": "--project-only 与 --global-only 互斥。"}, ensure_ascii=False, indent=2))
        else:
            console.print("[bold red]错误: --project-only 与 --global-only 互斥。[/bold red]")
        raise typer.Exit(1)

    scope_filter = None
    if project_only:
        scope_filter = CatalogScope.PROJECT
    elif global_only:
        scope_filter = CatalogScope.GLOBAL

    cascading = get_cascading_catalog()
    sources = cascading.list_sources(tag_filter=tag, scope_filter=scope_filter)

    if json_output:
        res = []
        for s in sources:
            abs_p = cascading.resolve_absolute_path(s)
            res.append({
                "source_id": s.source_id,
                "scope": s.scope.value,
                "project_name": s.project_name,
                "name": s.profile.name,
                "summary": s.profile.summary,
                "tags": s.profile.tags,
                "questions": s.profile.questions,
                "notes": s.profile.notes,
                "metadata_status": s.profile.metadata_status.value,
                "description_origin": s.profile.description_origin.value,
                "vendor": s.physical.vendor,
                "has_live_plugins": s.physical.has_live_plugins,
                "inputs_count": len(s.inputs),
                "inputs": [i.model_dump() for i in s.inputs],
                "file_path": s.file_path,
                "absolute_path": str(abs_p),
                "is_file_valid": abs_p.is_file(),
            })
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return

    if not sources:
        console.print("[yellow]当前未检索到任何数据源。请使用 excel-relay add <path> 添加。[/yellow]")
        return

    table = Table(title=f"[bold blue]已注册数据源清单 (共 {len(sources)} 个)[/bold blue]")
    table.add_column("ID (支持短Hash调用)", style="cyan", no_wrap=True)
    table.add_column("名称", style="bold")
    table.add_column("作用域", justify="center")
    table.add_column("分层", style="magenta")
    table.add_column("插件", style="yellow")
    table.add_column("可调参数", justify="center")
    table.add_column("标签", style="white")
    table.add_column("物理状态", style="dim")

    for s in sources:
        status_color = {"declared": "green", "enriched": "cyan", "raw": "dim"}.get(
            s.profile.metadata_status.value, "white"
        )
        badge = f"[{status_color}]{s.profile.metadata_status.value}[/{status_color}]"
        io_str = str(len(s.inputs))
        scope_tag = (
            "[bold green][项目专属][/bold green]"
            if s.scope == CatalogScope.PROJECT
            else "[bold cyan][用户全局][/bold cyan]"
        )
        abs_p = cascading.resolve_absolute_path(s)
        exists_str = "[green]有效[/green]" if abs_p.is_file() else "[red]失效(可运行 excel-relay clean)[/red]"

        table.add_row(
            s.source_id,
            s.profile.name,
            scope_tag,
            badge,
            s.physical.vendor,
            io_str,
            ",".join(s.profile.tags[:3]),
            exists_str,
        )

    console.print(table)


@app.command(name="set")
def set_source_metadata(
    source_id: str = typer.Argument(..., help="数据源 ID (支持完整 ID、前 6 位短 Hash 或文件路径)"),
    name: Optional[str] = typer.Option(None, "--name", help="设置模型中文名称"),
    summary: Optional[str] = typer.Option(None, "--summary", "-s", help="补充业务功能摘要"),
    tag: Optional[List[str]] = typer.Option(None, "--tag", "-t", help="设置标签(多选)"),
    question: Optional[List[str]] = typer.Option(None, "--question", "-q", help="设置推荐问法(多选)"),
    notes: Optional[str] = typer.Option(None, "--notes", "-n", help="补充背景备注说明"),
    input_def: Optional[List[str]] = typer.Option(
        None,
        "--input",
        "-i",
        help="登记可调输入参数坐标，如 -i '提价幅度=假设表!C4 (单位%)' 或 -i '假设表!C4'",
    ),
    contract_file: Optional[Path] = typer.Option(
        None, "--contract", help="读取项目底稿的日期参数和固定指标定义 JSON 文件。"
    ),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出更新后的数据源信息。"),
):
    """补充数据源语义、登记输入参数或配置项目输出指标。"""
    cascading = get_cascading_catalog()
    entry = cascading.get_source(source_id)
    if not entry:
        if json_output:
            print(json.dumps({"status": "error", "error": f"未找到对应数据源: {source_id}"}, ensure_ascii=False, indent=2))
        else:
            console.print(f"[bold red]未找到对应数据源:[/bold red] {source_id}")
        raise typer.Exit(1)

    has_profile_updates = any(v is not None for v in (name, summary, tag, question, notes))
    contract = None
    if contract_file is not None:
        try:
            if has_profile_updates or input_def is not None:
                raise ValueError("请先完成数据源或输入参数的更新，再单独登记指标配置。")
            spec = json.loads(contract_file.read_text(encoding="utf-8"))
            contract = configure_source_contract(cascading, entry.source_id, spec)
        except Exception as error:
            if json_output:
                print(json.dumps({"status": "error", "error": str(error)}, ensure_ascii=False, indent=2))
            else:
                console.print(f"登记底稿指标失败。{error}", style="bold red", markup=False)
            raise typer.Exit(1)

    if has_profile_updates:
        entry = cascading.update_source(
            source_id=source_id,
            name=name,
            summary=summary,
            tags=tag,
            questions=question,
            notes=notes,
            origin=DescriptionOrigin.USER_MANUAL,
        )

    if input_def is not None:
        try:
            entry = cascading.set_source_inputs(
                source_id=source_id,
                inputs=input_def,
            )
        except ValueError as ve:
            if json_output:
                print(json.dumps({"status": "error", "error": str(ve)}, ensure_ascii=False, indent=2))
            else:
                console.print(f"[bold red]参数坐标校验失败:[/bold red] {ve}")
            raise typer.Exit(1)

    if json_output:
        abs_p = cascading.resolve_absolute_path(entry)
        payload = {
            "status": "success",
            "source_id": entry.source_id,
            "scope": entry.scope.value,
            "project_name": entry.project_name,
            "name": entry.profile.name,
            "summary": entry.profile.summary,
            "tags": entry.profile.tags,
            "questions": entry.profile.questions,
            "notes": entry.profile.notes,
            "metadata_status": entry.profile.metadata_status.value,
            "description_origin": entry.profile.description_origin.value,
            "inputs": [i.model_dump() for i in entry.inputs],
            "file_path": entry.file_path,
            "absolute_path": str(abs_p),
        }
        if contract is not None:
            payload["contract"] = contract.model_dump()
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return

    scope_name = "【项目专属】" if entry.scope == CatalogScope.PROJECT else "【用户全局】"
    console.print(f"[bold green][OK] 成功更新数据源 {entry.source_id} ({scope_name})！[/bold green]")
    console.print(
        f"当前状态: [cyan]{entry.profile.metadata_status.value}[/cyan] "
        f"(来源: {entry.profile.description_origin.value})"
    )
    console.print(f"摘要: {entry.profile.summary}")
    if entry.inputs:
        in_desc = ", ".join([f"{i.name}({i.cell}={i.current_value})" for i in entry.inputs])
        console.print(f"已登记输入参数: [yellow]{in_desc}[/yellow]")
    if contract is not None:
        console.print(f"已登记 {len(contract.outputs)} 项固定指标。")


@app.command(name="rm")
def remove_source(
    source_id: str = typer.Argument(..., help="拟注销的数据源 ID、前 6 位短 Hash 或文件路径"),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出注销结果。"),
):
    """注销指定的数据源并删除其临时沙箱副本 (优先从当前项目注销，次选全局)."""
    cascading = get_cascading_catalog()
    result = cascading.unregister_detailed(source_id)
    ok = result is not None
    if json_output:
        payload = {
            "status": "success" if ok else "error",
            "source_id": source_id,
            "message": f"已成功注销数据源: {source_id}" if ok else f"未找到对应数据源: {source_id}",
        }
        if result is not None:
            payload["scope"] = result.scope.value
            payload["removed_artifacts_count"] = result.removed_artifacts_count
            payload["failed_artifact_paths"] = result.failed_artifact_paths
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        if not ok:
            raise typer.Exit(1)
        return

    if result is not None:
        console.print(f"[bold green]已成功注销数据源:[/bold green] {source_id}")
        console.print(f"[dim]已清理 {result.removed_artifacts_count} 个临时沙箱副本。[/dim]")
        if result.failed_artifact_paths:
            console.print("[bold yellow]以下文件可能被占用而未能删除，可稍后执行 `excel-relay clean` 清理:[/bold yellow]")
            for fp in result.failed_artifact_paths:
                console.print(f"  - {fp}")
    else:
        console.print(f"[bold red]未找到对应数据源:[/bold red] {source_id}")
        raise typer.Exit(1)


@app.command(name="find")
def search_sources(
    query: str = typer.Argument(..., help="搜索关键词或问法，如 '茅台 估值'"),
    project_only: bool = typer.Option(False, "--project-only", "-p", help="仅在当前项目专属数据源中检索"),
    global_only: bool = typer.Option(False, "--global-only", "-g", help="仅在用户全局数据源中检索"),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出检索结果。"),
):
    """检索最匹配的数据源 (跨仓级联检索，对本项目数据源赋予 Project Boost 优先加权)."""
    if project_only and global_only:
        if json_output:
            print(json.dumps({"status": "error", "error": "--project-only 与 --global-only 互斥。"}, ensure_ascii=False, indent=2))
        else:
            console.print("[bold red]错误: --project-only 与 --global-only 互斥。[/bold red]")
        raise typer.Exit(1)

    scope_filter = None
    if project_only:
        scope_filter = CatalogScope.PROJECT
    elif global_only:
        scope_filter = CatalogScope.GLOBAL

    cascading = get_cascading_catalog()
    matches = cascading.search(query, top_k=5, scope_filter=scope_filter)

    if json_output:
        res = []
        for entry, score, reason in matches:
            abs_p = str(cascading.resolve_absolute_path(entry))
            res.append({
                "source_id": entry.source_id,
                "scope": entry.scope.value,
                "project_name": entry.project_name,
                "name": entry.profile.name,
                "summary": entry.profile.summary,
                "tags": entry.profile.tags,
                "match_score": round(score, 1),
                "match_reason": reason,
                "vendor": entry.physical.vendor,
                "has_live_plugins": entry.physical.has_live_plugins,
                "recommended_questions": entry.profile.questions,
                "inputs": [i.model_dump() for i in entry.inputs],
                "file_path": entry.file_path,
                "absolute_path": abs_p,
            })
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return

    if not matches:
        console.print(f"[yellow]未找到与 '{query}' 匹配的数据源。[/yellow]")
        return

    console.print(f"[bold green]为您找到 {len(matches)} 个最匹配的数据源:[/bold green]\n")
    for entry, score, reason in matches:
        p = entry.profile
        scope_text = "[bold green][项目专属][/bold green]" if entry.scope == CatalogScope.PROJECT else "[bold cyan][用户全局][/bold cyan]"
        abs_p = cascading.resolve_absolute_path(entry)
        panel_content = (
            f"[bold cyan]ID:[/bold cyan] {entry.source_id}   {scope_text}   [bold]插件:[/bold] {entry.physical.vendor}\n"
            f"[bold]路径:[/bold] {entry.file_path} (物理绝对路径: {abs_p})\n"
            f"[bold]摘要:[/bold] {p.summary}\n"
            f"[bold]匹配原因:[/bold] [yellow]{reason}[/yellow] (匹配分: {round(score, 1)})\n"
        )
        if p.questions:
            panel_content += f"[bold]推荐问法:[/bold] {'; '.join(p.questions[:2])}\n"
        if entry.inputs:
            param_strs = [f"{inp.name} ({inp.cell}, 当前={inp.current_value})" for inp in entry.inputs]
            panel_content += f"[bold]可调参数:[/bold] {', '.join(param_strs)}\n"

        console.print(Panel(panel_content, title=f"[bold]{p.name}[/bold]", expand=False))


# -----------------------------------------------------------------------------
# 3. 刷新、沙箱重算与清理 (Run / History / Clean)
# -----------------------------------------------------------------------------
def _parse_cli_params(params_raw: Optional[List[str]]) -> Optional[dict]:
    """将 CLI 命令行传入的 ['k=v', ...] 字符串列表解析为标准参数字典."""
    if not params_raw:
        return None
    param_dict = {}
    for p_item in params_raw:
        if "=" in p_item:
            k, v = p_item.split("=", 1)
            v_strip = v.strip()
            try:
                v_num = float(v_strip) if "." in v_strip else int(v_strip)
                param_dict[k.strip()] = v_num
            except ValueError:
                param_dict[k.strip()] = v_strip
    return param_dict if param_dict else None


@app.command(name="run")
def run_refresh(
    source_id_or_file: str = typer.Argument(..., help="已注册的数据源 ID (或前6位短Hash) 或文件路径"),
    params: Optional[List[str]] = typer.Option(
        None, "--param", "-p", help="已登记参数赋值。已声明的日期参数接受 YYYY-MM-DD 或 Excel 序列值。"
    ),
    output: Optional[Path] = typer.Option(None, "--output", "-o", help="另存为输出文件路径"),
    timeout: int = typer.Option(120, "--timeout", "-t", help="刷新超时秒数，Windows 在插件持续返回数据时可顺延"),
    keep_formulas: bool = typer.Option(
        False, "--keep-formulas", help="保留结果副本中的公式。macOS 始终保留公式，Windows 默认固化为数值"
    ),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出计算与沙箱产物信息。"),
):
    """刷新已注册的 Excel 数据源。macOS 将 Wind .xlsx 的结果保存为副本。"""
    cascading = get_cascading_catalog()
    entry = cascading.get_source(source_id_or_file)

    if not entry:
        if json_output:
            print(json.dumps({"status": "error", "error": f"未在当前上下文或全局 Catalog 中找到对应数据源: {source_id_or_file}"}, ensure_ascii=False, indent=2))
        else:
            console.print(f"[bold red]未在当前上下文或全局 Catalog 中找到对应数据源:[/bold red] {source_id_or_file}")
            console.print("[yellow]提示: 所有刷新与改参操作均基于 Catalog 契约管理，请先执行 'excel-relay add <文件路径>' 注册数据源。[/yellow]")
        raise typer.Exit(1)

    param_dict = _parse_cli_params(params)
    allowed_inputs = entry.inputs
    is_macos = sys.platform == "darwin"
    is_sandbox = is_macos or bool(allowed_inputs) or bool(param_dict) or (output is not None)
    keep_formulas = is_macos or keep_formulas

    if not json_output:
        scope_str = "项目专属" if entry.scope == CatalogScope.PROJECT else "用户全局"
        engine = "桌面 Excel" if is_macos else "后台静默重算引擎"
        console.print(f"[bold cyan]启动{engine} ({scope_str})...[/bold cyan]")
        abs_src = cascading.resolve_absolute_path(entry)
        console.print(f"数据源: {entry.source_id} ({abs_src})")
        if allowed_inputs:
            full_p = HeadlessExcelRefresher.build_full_params(param_dict, allowed_inputs)
            console.print(f"参数配置 X: {full_p} (带参模型 f(t,X) 启用母版只读保护与沙箱隔离)")
        if is_sandbox:
            mode_text = "[yellow]保留全部公式[/yellow]" if keep_formulas else "[green]纯数值快照 (自动脱水)[/green]"
            console.print(f"产物策略: {mode_text}")

    try:
        status_display = (
            nullcontext()
            if json_output
            else console.status("[bold green]Excel 正在后台计算中，请稍候...[/bold green]")
        )
        with status_display:
            out_path, elapsed, artifact, outputs = HeadlessExcelRefresher.refresh_with_outputs(
                source_id_or_path=entry.source_id,
                params=param_dict,
                output_path=output,
                timeout_sec=timeout,
                keep_formulas=keep_formulas,
                catalog_mgr=cascading,
            )
    except RefreshError as re_err:
        busy = isinstance(re_err, RefreshBusyError)
        if json_output:
            print(json.dumps({
                "status": "busy" if busy else "error",
                "source_id": entry.source_id,
                "error": str(re_err),
            }, ensure_ascii=False, indent=2))
        elif busy:
            console.print(f"Excel Relay 忙碌。{re_err}", style="yellow", markup=False)
        else:
            console.print(f"[bold red]刷新被拦截或失败:[/bold red] {re_err}")
        raise typer.Exit(1)
    except Exception as ex:
        if json_output:
            print(json.dumps({"status": "error", "source_id": entry.source_id, "error": str(ex)}, ensure_ascii=False, indent=2))
        else:
            console.print(f"[bold red]发生未知错误:[/bold red] {ex}")
        raise typer.Exit(1)

    if json_output:
        payload = {
            "status": "success",
            "source_id": entry.source_id,
            "scope": entry.scope.value,
            "mode": "sandbox_copy" if artifact else "in_place_refresh",
            "artifact_id": artifact.artifact_id if artifact else None,
            "result_file_path": out_path,
            "params": artifact.params if artifact else param_dict,
            "params_hash": artifact.params_hash if artifact else "base",
            "keep_formulas": artifact.keep_formulas if artifact else True,
            "elapsed_seconds": elapsed,
        }
        if outputs is not None:
            payload["outputs"] = outputs
            payload["refreshed_at"] = artifact.created_at if artifact else None
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return

    console.print(f"[bold green][OK] 计算刷新成功！[/bold green] (耗时 {elapsed} 秒)")
    console.print(f"结果文件路径: [bold underline]{out_path}[/bold underline]")

    if artifact:
        console.print(
            f"产物唯一 ID (artifact_id): [bold cyan]{artifact.artifact_id}[/bold cyan] "
            f"(参数Hash: {artifact.params_hash}) 已由 core 引擎自动同步至归属仓储"
        )
    else:
        console.print("[dim]无参模型 f(t) 原地刷新完成，core 引擎已自动同步更新母版最后刷新时间。[/dim]")


@app.command(name="history")
def show_history(
    source_id_or_file: str = typer.Argument(..., help="数据源 ID (或前6位短Hash) 或文件路径"),
    limit: int = typer.Option(10, "--limit", "-n", help="最多展示条数"),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出测算产物清单。"),
):
    """查看数据源的不同参数配置测算产物清单 (按 artifact_id 唯一标识)."""
    cascading = get_cascading_catalog()
    entry = cascading.get_source(source_id_or_file)
    if not entry:
        if json_output:
            print(json.dumps({"status": "error", "error": f"未找到对应数据源: {source_id_or_file}"}, ensure_ascii=False, indent=2))
        else:
            console.print(f"[bold red]未找到对应数据源:[/bold red] {source_id_or_file}")
        raise typer.Exit(1)

    artifacts = cascading.list_artifacts(entry.source_id, limit=limit)
    if json_output:
        res = []
        for art in artifacts:
            res.append({
                "artifact_id": art.artifact_id,
                "created_at": art.created_at,
                "params": art.params,
                "params_hash": art.params_hash,
                "keep_formulas": art.keep_formulas,
                "elapsed_seconds": art.elapsed_seconds,
                "file_path": art.file_path,
                "absolute_path": str(cascading.resolve_absolute_path(art)),
            })
        print(json.dumps({
            "status": "success",
            "source_id": entry.source_id,
            "scope": entry.scope.value,
            "name": entry.profile.name,
            "count": len(res),
            "artifacts": res,
        }, ensure_ascii=False, indent=2))
        return

    if not artifacts:
        console.print(f"[yellow]数据源 {entry.source_id} 暂无沙箱测算产物记录。[/yellow]")
        return

    table = Table(
        title=f"[bold green]数据源 {entry.profile.name} ({entry.source_id}) 参数配置产物清单 (共 {len(artifacts)} 项)[/bold green]"
    )
    table.add_column("Artifact ID", style="cyan", no_wrap=True)
    table.add_column("最新刷新时间 t", style="magenta")
    table.add_column("全量参数配置 X", style="yellow")
    table.add_column("格式", style="blue")
    table.add_column("耗时", justify="right")
    table.add_column("产物路径", style="green")

    for art in artifacts:
        p_str = json.dumps(art.params, ensure_ascii=False) if art.params else "(无参基准)"
        mode_str = "保留公式" if art.keep_formulas else "纯数值"
        time_str = art.created_at[:19].replace("T", " ")
        table.add_row(
            art.artifact_id,
            time_str,
            p_str,
            mode_str,
            f"{art.elapsed_seconds}s",
            art.file_path,
        )
    console.print(table)


@app.command(name="clean")
def clean_catalog_cmd(
    source_id: Optional[str] = typer.Option(None, "--source", "-s", help="仅清理特定数据源的临时副本"),
    days: int = typer.Option(7, "--days", "-d", help="清理超过指定天数的临时沙箱副本"),
    keep_stale_sources: bool = typer.Option(False, "--keep-stale", help="保留源文件已移动/删除的失效数据源条目"),
    global_only: bool = typer.Option(False, "--global-only", "-g", help="仅清理用户全局数据源"),
    project_only: bool = typer.Option(False, "--project-only", "-p", help="仅清理当前项目专属数据源"),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出清理结果。"),
):
    """清理失效的数据源条目 (因文件改名/移动/删除)、孤儿产物目录及过期的临时沙箱副本."""
    if project_only and global_only:
        if json_output:
            print(json.dumps({"status": "error", "error": "--project-only 与 --global-only 互斥。"}, ensure_ascii=False, indent=2))
        else:
            console.print("[bold red]错误: --project-only 与 --global-only 互斥。[/bold red]")
        raise typer.Exit(1)

    scope_filter = None
    if project_only:
        scope_filter = CatalogScope.PROJECT
    elif global_only:
        scope_filter = CatalogScope.GLOBAL

    cascading = get_cascading_catalog()

    removed_sources = []
    if not keep_stale_sources and not source_id:
        removed_sources = cascading.clean_stale_sources(scope_filter=scope_filter)
        if not json_output:
            if removed_sources:
                console.print(f"[bold yellow]已自动移除 {len(removed_sources)} 个源文件不存在的失效数据源条目:[/bold yellow]")
                for sid, fp, sc in removed_sources:
                    sc_str = "[项目]" if sc == CatalogScope.PROJECT else "[全局]"
                    console.print(f"  - {sc_str} [cyan]{sid}[/cyan] ({fp})")
            else:
                console.print("[dim]未发现失效的数据源条目，所有已注册文件路径均有效。[/dim]")

    # 孤儿产物目录清理 (各仓 artifacts/ 下已不对应任何已注册数据源的 {source_id} 目录)
    removed_orphans = []
    if not source_id:
        removed_orphans = cascading.clean_orphan_artifact_dirs(scope_filter=scope_filter)
        if not json_output:
            if removed_orphans:
                console.print(f"[bold yellow]已移除 {len(removed_orphans)} 个孤儿产物目录:[/bold yellow]")
                for sid, dp, sc in removed_orphans:
                    sc_str = "[项目]" if sc == CatalogScope.PROJECT else "[全局]"
                    console.print(f"  - {sc_str} [cyan]{sid}[/cyan] ({dp})")
            else:
                console.print("[dim]未发现孤儿产物目录。[/dim]")

    cleaned = cascading.clean_expired_artifacts(
        source_id_or_path=source_id,
        max_age_days=days,
        scope_filter=scope_filter,
    )

    if json_output:
        print(json.dumps({
            "status": "success",
            "removed_stale_sources_count": len(removed_sources),
            "removed_stale_sources": [
                {"source_id": sid, "file_path": fp, "scope": sc.value}
                for sid, fp, sc in removed_sources
            ],
            "removed_orphan_dirs_count": len(removed_orphans),
            "removed_orphan_dirs": [
                {"source_id": sid, "dir_path": dp, "scope": sc.value}
                for sid, dp, sc in removed_orphans
            ],
            "cleaned_expired_artifacts_count": cleaned,
        }, ensure_ascii=False, indent=2))
        return

    console.print(f"[bold green][OK] 成功清理 {cleaned} 个超过 {days} 天(或已丢失)的临时沙箱副本！[/bold green]")


@app.command(name="open")
def open_file_cmd(
    target: str = typer.Argument(..., help="已登记的数据源 ID (或前6位短Hash)、已注册母版路径，或测算产物 artifact_id"),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出打开结果。"),
):
    """在系统默认程序 (如 Excel) 中直接打开已注册的数据源母版或沙箱测算产物."""
    cascading = get_cascading_catalog()
    file_path: Optional[Path] = None

    # 1. 优先尝试查找测算产物 (RunArtifact)
    art = cascading.find_artifact_by_id(target)
    if art:
        abs_art = cascading.resolve_absolute_path(art)
        if abs_art.is_file():
            file_path = abs_art
    if not file_path:
        # 2. 尝试查找已注册的数据源母版
        entry = cascading.get_source(target)
        if entry:
            abs_p = cascading.resolve_absolute_path(entry)
            if abs_p.is_file():
                file_path = abs_p

    # 3. 拦截未注册文件或不存在的目标
    if not file_path:
        p = Path(target)
        try:
            if p.is_file():
                if json_output:
                    print(json.dumps({"status": "error", "error": f"文件 '{target}' 尚未在 Catalog 中注册。"}, ensure_ascii=False, indent=2))
                else:
                    console.print(f"[bold red]拒绝打开:[/bold red] 文件 '{target}' 尚未在 Catalog 中注册。")
                    console.print("[yellow]提示: 所有资产操作均严格基于 Catalog 契约管理，请先执行 'excel-relay add <文件路径>' 注册数据源。[/yellow]")
                raise typer.Exit(1)
        except typer.Exit:
            raise
        except Exception:
            pass
        if json_output:
            print(json.dumps({"status": "error", "error": f"无法定位目标数据源或测算产物: {target}"}, ensure_ascii=False, indent=2))
        else:
            console.print(f"[bold red]无法定位目标数据源或测算产物:[/bold red] {target}")
        raise typer.Exit(1)

    if sys.platform == "darwin":
        import subprocess

        try:
            subprocess.run(["/usr/bin/open", str(file_path)], check=True, capture_output=True, text=True)
        except (OSError, subprocess.CalledProcessError) as exc:
            details = (getattr(exc, "stderr", None) or "").strip()
            error = f"系统无法打开文件 {file_path}：{details or exc}"
            if json_output:
                print(json.dumps({"status": "error", "error": error}, ensure_ascii=False, indent=2))
            else:
                console.print(f"[bold red]{error}[/bold red]")
            raise typer.Exit(1) from exc
        if json_output:
            print(json.dumps({"status": "success", "file_path": str(file_path)}, ensure_ascii=False, indent=2))
        else:
            console.print(f"[bold green]已交给系统打开文件：[/bold green] {file_path}")
        return

    if json_output:
        print(json.dumps({"status": "success", "file_path": str(file_path)}, ensure_ascii=False, indent=2))
    else:
        console.print(f"[bold green]正在调用系统程序打开文件:[/bold green] {file_path}")

    import os
    if sys.platform == "win32":
        os.startfile(str(file_path))
    else:
        import subprocess
        subprocess.run(["xdg-open", str(file_path)])


# -----------------------------------------------------------------------------
# 4. 数据源包工具 (Pack)
# -----------------------------------------------------------------------------
@pack_app.command("init")
def pack_init(
    file_path: Path = typer.Argument(..., help="目标 Excel 文件路径"),
    hidden: bool = typer.Option(
        False, "--hidden", help="是否将 _excel_relay 工作表设为隐藏 (默认可见并排在首个标签页，方便直接填写)"
    ),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出模板注入结果。"),
):
    """向目标 Excel 注入两段式全中文 _excel_relay 标准元数据模板表."""
    try:
        ok = ExcelInspector.inject_spec_sheet(file_path, hidden=hidden)
        if json_output:
            print(json.dumps({
                "status": "success",
                "injected": ok,
                "file_path": str(file_path.resolve()),
                "hidden": hidden,
                "message": "成功向 Excel 注入 _excel_relay 元数据模板表" if ok else "该文件已存在 _excel_relay 工作表，无需重复注入",
            }, ensure_ascii=False, indent=2))
            return

        if ok:
            vis_text = "已设为隐藏" if hidden else "位于首个工作表标签页"
            console.print(f"[bold green][OK] 成功向 {file_path.name} 注入 _excel_relay 元数据模板表 ({vis_text})！[/bold green]")
            console.print("您可以直接打开该 Excel 文件，在 [bold yellow]浅黄色填写区[/bold yellow] 补充摘要与可调参数位置，保存后执行 excel-relay add 注册。")
        else:
            console.print("[yellow]该文件已存在 _excel_relay 工作表，无需重复注入。[/yellow]")
    except Exception as e:
        if json_output:
            print(json.dumps({"status": "error", "error": str(e)}, ensure_ascii=False, indent=2))
        else:
            console.print(f"[bold red]注入失败:[/bold red] {e}")
        raise typer.Exit(1)


@pack_app.command("inspect")
def pack_inspect(
    file_path: Path = typer.Argument(..., help="Excel 文件路径"),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出离线剖析结果。"),
):
    """离线分析 Excel 文件，查看提取的物理结构、业务画像与可调参数接口."""
    try:
        inspected = ExcelInspector.inspect_file(file_path)
    except Exception as e:
        if json_output:
            print(json.dumps({"status": "error", "error": str(e)}, ensure_ascii=False, indent=2))
        else:
            console.print(f"[bold red]探测失败:[/bold red] {e}")
        raise typer.Exit(1)

    p = inspected.profile
    phys = inspected.physical

    if json_output:
        payload = {
            "status": "success",
            "file_path": str(file_path.resolve()),
            "has_official_spec": phys.has_official_spec,
            "vendor": phys.vendor,
            "has_live_plugins": phys.has_live_plugins,
            "profile": p.model_dump(),
            "inputs": [i.model_dump() for i in inspected.inputs],
            "sheets": [
                {
                    "name": s.name,
                    "row_count": s.row_count,
                    "col_count": s.col_count,
                    "header_row": s.header_row,
                    "columns": s.columns,
                    "row_headers": s.row_headers,
                    "has_formulas": s.has_formulas,
                }
                for s in phys.sheets
            ],
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return

    status = (
        f"[green]已内置 _excel_relay 规范表 (语义状态: {p.metadata_status.value})[/green]"
        if phys.has_official_spec
        else f"[yellow]普通文件 (语义状态: {p.metadata_status.value})[/yellow]"
    )

    in_lines = [
        f"{i.name} [坐标={i.cell}, 当前值={i.current_value}" + (f", 说明={i.description}]" if i.description else "]")
        for i in inspected.inputs
    ]

    sheet_lines = []
    for s in phys.sheets:
        col_preview = ", ".join(s.columns[:6])
        row_preview = ", ".join(s.row_headers[:6])
        line = f"• {s.name} ({s.row_count}x{s.col_count}, 表头行={s.header_row}): 横向=[{col_preview}]"
        if row_preview:
            line += f" | 纵向=[{row_preview}]"
        sheet_lines.append(line)

    panel = Panel(
        f"[bold]状态:[/bold] {status}\n"
        f"[bold]文件路径:[/bold] [cyan]{inspected.file_path}[/cyan]\n"
        f"[bold]模型名称:[/bold] {p.name}\n"
        f"[bold]功能摘要:[/bold] {p.summary}\n"
        f"[bold]插件依赖:[/bold] {phys.vendor} (含动态公式: {phys.has_live_plugins})\n"
        f"[bold]检索标签:[/bold] {', '.join(p.tags)}\n"
        f"[bold]典型问法:[/bold] {'; '.join(p.questions)}\n"
        f"[bold]口径备注:[/bold] {p.notes or '无'}\n"
        f"[bold]可调输入参数 ({len(inspected.inputs)}):[/bold] {'; '.join(in_lines) or '无 (禁止传参修改)'}\n"
        f"[bold]工作表结构 ({len(phys.sheets)}):[/bold]\n  " + "\n  ".join(sheet_lines),
        title=f"[bold cyan]数据源离线剖析卡片: {file_path.name}[/bold cyan]",
        expand=False,
    )
    console.print(panel)


@pack_app.command("validate")
def pack_validate(
    file_path: Path = typer.Argument(..., help="Excel 文件路径"),
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出校验结果。"),
):
    """校验 Excel 是否包含规范的 _excel_relay 元数据表及有效接口坐标."""
    _entry, errors, warnings = ExcelInspector.validate_package(file_path)

    if json_output:
        payload = {
            "status": "valid" if not errors else "invalid",
            "file_path": str(file_path.resolve()),
            "errors": errors,
            "warnings": warnings,
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        if errors:
            raise typer.Exit(1)
        return

    if not errors and not warnings:
        console.print(f"[bold green][OK] 恭喜！{file_path.name} 完全符合 excel-relay 规范要求！[/bold green]")
    else:
        if errors:
            console.print("[bold red]发现以下规范或坐标错误:[/bold red]")
            for err in errors:
                console.print(f"  [FAIL] {err}")
        if warnings:
            console.print("[bold yellow]发现以下优化建议:[/bold yellow]")
            for w in warnings:
                console.print(f"  ! {w}")


# -----------------------------------------------------------------------------
# 5. MCP 服务启动 (Serve)
# -----------------------------------------------------------------------------
@app.command(name="serve")
def serve_mcp():
    """以标准 stdio 模式启动 excel-relay MCP Server (供 AI Agent 连接)."""
    from excel_relay.interface.mcp import run_server
    run_server()


# -----------------------------------------------------------------------------
# 6. Windows 右键菜单管理 (Menu)
# -----------------------------------------------------------------------------
@menu_app.command("install")
def menu_install():
    """向 Windows 资源管理器添加 Excel 文件与文件夹二级级联右键菜单."""
    try:
        targets = install_context_menu()
        console.print("[bold green][OK] 成功为 Windows 资源管理器启用 ExcelRelay 级联右键菜单！[/bold green]")
        console.print(f"已绑定目标: [cyan]{', '.join(targets)}[/cyan]\n")
        console.print("1. 在任意 [bold]Excel 文件[/bold] 上点击右键，展开 [bold][ExcelRelay] 投研资产引擎[/bold]：")
        for idx, item in enumerate(SUB_MENU_ITEMS, start=1):
            console.print(f"   {idx}. [bold yellow]{item['label']}[/bold yellow]")
        console.print("\n2. 在任意 [bold]文件夹图标或文件夹空白处[/bold] 点击右键，展开 [bold][ExcelRelay] 投研资产引擎[/bold]：")
        for idx, item in enumerate(DIR_SUB_MENU_ITEMS, start=1):
            console.print(f"   {idx}. [bold yellow]{item['label']}[/bold yellow]")
    except Exception as e:
        console.print(f"[bold red]安装右键菜单失败:[/bold red] {e}")
        raise typer.Exit(1)


@menu_app.command("uninstall")
def menu_uninstall():
    """从 Windows 资源管理器中移除所有 ExcelRelay 右键菜单项."""
    try:
        uninstall_context_menu()
        console.print("[bold green][OK] 已成功从 Windows 资源管理器移除所有 ExcelRelay 右键菜单！[/bold green]")
    except Exception as e:
        console.print(f"[bold red]卸载右键菜单失败:[/bold red] {e}")
        raise typer.Exit(1)


@menu_app.command("status")
def menu_status(
    json_output: bool = typer.Option(False, "--json", help="以 JSON 格式输出右键菜单状态。"),
):
    """检查右键菜单在各文件类型与文件夹中的启用状态."""
    if sys.platform == "darwin":
        error = "右键菜单管理仅适用于 Windows 资源管理器。"
        if json_output:
            print(json.dumps({"status": "unsupported", "error": error}, ensure_ascii=False, indent=2))
        else:
            console.print(error)
        return
    status_list = get_menu_status()
    if json_output:
        res = [
            {"target": target_name, "enabled": enabled, "labels": labels}
            for target_name, enabled, labels in status_list
        ]
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return

    table = Table(title="ExcelRelay Windows 右键菜单启用状态")
    table.add_column("绑定目标", style="cyan")
    table.add_column("状态", style="bold")
    table.add_column("已启用菜单项", style="white")

    for target_name, enabled, labels in status_list:
        badge = "[green][OK] 全部启用[/green]" if enabled else ("[yellow]部分启用[/yellow]" if labels else "[dim]未启用[/dim]")
        table.add_row(target_name, badge, " | ".join(labels) if labels else "-")

    console.print(table)


@app.command("uninstall")
def self_uninstall(
    yes: bool = typer.Option(False, "--yes", "-y", help="跳过二次确认，直接执行完全卸载。"),
):
    """从系统中彻底卸载 excel-relay：注销 Windows 级联右键菜单并清理核心引擎与沙箱环境."""
    if not yes:
        console.print(
            Panel.fit(
                "[bold red]【完全卸载 excel-relay 投研资产引擎】[/bold red]\n\n"
                "此操作将执行以下清理：\n"
                "  1. 从 Windows 资源管理器注销所有 ExcelRelay 级联右键快捷菜单 (HKCU 注册表)\n"
                "  2. 从系统彻底卸载 excel-relay 核心引擎及其专用 Python 沙箱环境\n\n"
                "[dim](提示：您配置的 Agent MCP 设置与您的 Excel 投研底稿资产不会被删除)[/dim]",
                title="卸载确认",
                border_style="red",
            )
        )
        confirm = typer.confirm("确定要从本机彻底卸载 excel-relay 吗？", default=False)
        if not confirm:
            console.print("[yellow]已取消卸载操作。[/yellow]")
            raise typer.Exit(0)

    # 1. 注销 Windows 资源管理器右键菜单
    console.print("\n[1/2] 正在从 Windows 资源管理器注销右键菜单...")
    try:
        uninstall_context_menu()
        console.print("      [bold green][OK] 右键快捷菜单已注销。[/bold green]")
    except Exception as e:
        console.print(f"      [yellow][提示] 注销右键菜单: {e}[/yellow]")

    # 2. 调度自毁清理任务 (Detached Process)
    console.print("[2/2] 正在清理核心引擎与独立沙箱环境...")
    import importlib.util
    import shutil
    import subprocess

    prefix_norm = sys.prefix.replace("\\", "/").lower()
    is_uv_tool = "uv/tools/excel-relay" in prefix_norm

    uv_exe = (
        shutil.which("uv")
        or shutil.which("uv.exe")
        or (Path.home() / ".local" / "bin" / "uv.exe")
        or (Path.home() / ".cargo" / "bin" / "uv.exe")
    )
    if uv_exe and not Path(uv_exe).is_file():
        uv_exe = None

    has_pip = importlib.util.find_spec("pip") is not None

    if sys.platform == "win32":
        if is_uv_tool and uv_exe:
            # 模式 1：标准 uv tool 全局独立沙箱环境
            detach_cmd = f'cmd.exe /c "ping 127.0.0.1 -n 2 >nul & \"{uv_exe}\" tool uninstall excel-relay >nul 2>&1"'
        elif is_uv_tool:
            # 模式 2：uv tool 环境，但当前系统缺失 uv 二进制（直接物理清理沙箱与可执行垫片）
            tools_dir = Path.home() / ".local" / "share" / "uv" / "tools" / "excel-relay"
            shim_file = Path.home() / ".local" / "bin" / "excel-relay.exe"
            detach_cmd = (
                f'cmd.exe /c "ping 127.0.0.1 -n 2 >nul & '
                f'(if exist \"{tools_dir}\" rmdir /s /q \"{tools_dir}\") & '
                f'(if exist \"{shim_file}\" del /f /q \"{shim_file}\")"'
            )
        elif uv_exe:
            # 模式 3：通过 uv pip install 安装（即便虚拟环境中无 pip 命令，也能通过 uv pip uninstall 干净卸载）
            py_exe = sys.executable
            detach_cmd = f'cmd.exe /c "ping 127.0.0.1 -n 2 >nul & \"{uv_exe}\" pip uninstall --python \"{py_exe}\" excel-relay >nul 2>&1"'
        elif has_pip:
            # 模式 4：常规 Python 环境，内置 pip
            py_exe = sys.executable
            detach_cmd = f'cmd.exe /c "ping 127.0.0.1 -n 2 >nul & \"{py_exe}\" -m pip uninstall -y excel-relay >nul 2>&1"'
        else:
            # 模式 5：极端情况兜底（既无 pip 也无 uv）：解析 Python PEP 376 RECORD 元数据直接物理清理已安装文件
            import importlib.metadata
            file_del_cmds = []
            dir_rm_cmds = []
            try:
                dist = importlib.metadata.distribution("excel-relay")
                name = dist.metadata.get("Name", "excel-relay")
                ver = dist.version
                for f in (dist.files or []):
                    p = f.locate().resolve()
                    if p.is_file():
                        file_del_cmds.append(f'del /f /q "{p}"')
                    if p.parent.name in (name, f"{name}-{ver}.dist-info"):
                        dir_rm_cmds.append(f'rmdir /s /q "{p.parent}"')
            except Exception:
                pass

            all_cmds = " & ".join(list(set(file_del_cmds)) + list(set(dir_rm_cmds)))
            if all_cmds:
                detach_cmd = f'cmd.exe /c "ping 127.0.0.1 -n 2 >nul & {all_cmds}"'
            else:
                detach_cmd = 'cmd.exe /c "exit 0"'

        subprocess.Popen(
            detach_cmd,
            shell=True,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS,
            close_fds=True,
        )
    else:
        # 非 Windows (POSIX) 平台
        if is_uv_tool and uv_exe:
            subprocess.Popen([str(uv_exe), "tool", "uninstall", "excel-relay"])
        elif uv_exe:
            subprocess.Popen([str(uv_exe), "pip", "uninstall", "--python", sys.executable, "excel-relay"])
        elif has_pip:
            subprocess.Popen([sys.executable, "-m", "pip", "uninstall", "-y", "excel-relay"])

    console.print("      [bold green][OK] 核心引擎自毁指令已成功下发！[/bold green]")
    console.print("\n[bold cyan]卸载完成。当前窗口关闭后，相关沙箱文件将在后台自动清理完毕。感谢使用Excel Relay！[/bold cyan]\n")
    raise typer.Exit(0)


if __name__ == "__main__":
    app()
