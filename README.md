# Excel Relay

Excel Relay 将 Excel 投研底稿登记到本地资产目录。可通过 CLI 或 MCP 查询底稿、按已登记的参数白名单重算，并获取刷新后的 Excel 文件路径。项目目录使用 `.excel-relay/catalog.json`，用户全局目录使用 `~/.excel-relay/catalog.json`。

## 安装

使用 macOS 原生 `uv` 和 Python 3.10 或更高版本。在 wheel 所在目录执行以下命令。

```bash
uv tool install './excel_relay-0.1.0-py3-none-any.whl'
excel-relay --version
excel-relay doctor --json
```

单项目试用可安装到项目内的隔离环境。

```bash
uv venv .excel-relay-venv
uv pip install --python .excel-relay-venv/bin/python '/绝对路径/excel_relay-0.1.0-py3-none-any.whl'
.excel-relay-venv/bin/excel-relay doctor --json
```

## 刷新 Wind 底稿

首次刷新前，在本机图形会话打开 Excel，确认 Wind 加载项已启用且 Wind 账号已登录。首次收到 macOS 自动化提示时，允许发起方控制 Microsoft Excel。运行期间需独占 Excel 工作簿会话，因为刷新会短暂清理该会话的 Wind 缓存。请保存并关闭其他工作簿。

在底稿所在的项目目录运行以下命令。使用隔离环境时，将命令名替换成 `.excel-relay-venv/bin/excel-relay`。

```bash
excel-relay init
excel-relay add './底稿.xlsx' --project
excel-relay ls --project-only
excel-relay run '<source_id>' --json
```

输出的 `result_file_path` 指向刷新后的副本，项目资产的产物默认保存在 `.excel-relay/artifacts/`。macOS 版只写副本，并保留取数公式。Wind 请求及结果通过校验后才交付文件。遇到权限异常、Wind 未登录、等待超时或取数证据缺失时，按报错处理后重试。

首版的动态刷新范围是已装 Wind 加载项的桌面版 Excel 和 `.xlsx` 底稿。`.xlsm`、`.xls` 以及 iFinD 的动态刷新需要其他运行环境。Windows 资源管理器右键菜单在 Windows 上使用。

## 底稿参数与迁移

`excel-relay pack init './底稿.xlsx'` 会在指定文件内创建 `_excel_relay` 工作表，供填写底稿简介和可调参数位置。迁移旧底稿时，先在 Excel 中把已有的 `_cjxl` 工作表改名为 `_excel_relay` 并保存，再运行 `excel-relay add` 登记到新目录。

## 接入 MCP

将宿主的本地 stdio MCP 命令设置为 `excel-relay` 的绝对路径，参数为 `serve`。项目级配置可指向隔离环境中的 `.excel-relay-venv/bin/excel-relay`。

```json
{
  "mcpServers": {
    "excel-relay": {
      "command": "/项目绝对路径/.excel-relay-venv/bin/excel-relay",
      "args": ["serve"]
    }
  }
}
```

调用 `refresh_data_source` 刷新已注册资产，项目底稿的 `workspace_path` 传入项目目录的绝对路径。使用响应中的 `result_file_path` 读取产物。宿主需以能访问同一 Excel 图形会话的用户身份运行。

## 从源码构建

在源码目录使用 `uv 0.11.32` 执行以下命令。

```bash
uv build --offline --wheel
```

## 许可证

Excel Relay 采用 [MIT 许可证](LICENSE)。
