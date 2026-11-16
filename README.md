# Excel Relay

Excel Relay 将 Excel 投研底稿登记到本地资产目录。可通过 CLI 或 MCP 查询底稿、按已登记的参数白名单重算，并获取刷新后的 Excel 文件和已登记的指标值。项目目录使用 `.excel-relay/catalog.json`，用户全局目录使用 `~/.excel-relay/catalog.json`。

## 安装

使用 macOS 原生 `uv` 和 Python 3.10 或更高版本。在 wheel 所在目录执行以下命令。

```bash
uv tool install './excel_relay-0.2.0-py3-none-any.whl'
excel-relay --version
excel-relay doctor --json
```

单项目试用可安装到项目内的隔离环境。

```bash
uv venv .excel-relay-venv
uv pip install --python .excel-relay-venv/bin/python '/绝对路径/excel_relay-0.2.0-py3-none-any.whl'
.excel-relay-venv/bin/excel-relay doctor --json
```

以上命令安装正式版 v0.2.0。开发时可在源码目录创建独立的可编辑环境。

```bash
uv venv .venv
uv pip install --python .venv/bin/python --editable .
```

## 刷新 Wind 底稿

首次刷新前，在本机图形会话打开 Excel，确认 Wind 加载项已启用且 Wind 账号已登录。首次收到 macOS 自动化提示时，允许发起方控制 Microsoft Excel。

v0.2.0 在 macOS 上使用独立的隐藏 Excel 实例，同机并行刷新已用桌面 Excel 和 Wind 加载项完成实机验收。每个用户同时运行一个后台任务，其他刷新请求返回 `status=busy`。

系统会等待能归属到后台进程的 Wind 新响应，再按后台进程 ID 触发完整重算，让底稿读取已回到该实例的数据。确认公式结果就绪后才保存副本。若无法确认任务实例已退出，错误信息会给出保留的副本和恢复记录位置。

在底稿所在的项目目录运行以下命令。使用开发环境时，命令指向源码目录的 `.venv/bin/excel-relay`。

```bash
excel-relay init
excel-relay add './底稿.xlsx' --project
excel-relay ls --project-only
excel-relay run '<source_id>' --json
```

输出的 `result_file_path` 指向刷新后的副本，项目资产的产物默认保存在 `.excel-relay/artifacts/`。macOS 版只写副本，并保留取数公式。Wind 请求及结果通过校验后才交付文件。遇到权限异常、Wind 未登录、等待超时或取数证据缺失时，按报错处理后重试。

macOS 的动态刷新范围是已装 Wind 加载项的桌面版 Excel 和 `.xlsx` 底稿。`.xlsm`、`.xls` 以及 iFinD 的动态刷新需要其他运行环境。Windows 资源管理器右键菜单在 Windows 上使用。

## 底稿参数与迁移

`excel-relay pack init './底稿.xlsx'` 会在指定文件内创建 `_excel_relay` 工作表，供填写底稿简介和可调参数位置。迁移旧底稿时，先在 Excel 中把已有的 `_cjxl` 工作表改名为 `_excel_relay` 并保存，再运行 `excel-relay add` 登记到新目录。

### 输入日期和固定指标

在 v0.2.0 中，先登记可调参数的实际单元格，再单独登记日期类型和结果指标。以下示例假设底稿只需一个可调参数，截止日在 `日期!A2`，指标位于 `海外!C2`。

```bash
excel-relay set '<source_id>' -i '截止日=日期!A2'
```

将以下定义保存为项目中的 `指标.json`。

```json
{
  "date_inputs": ["截止日"],
  "outputs": [
    {"name": "恒指周涨跌幅", "cell": "海外!C2", "unit": "%"}
  ]
}
```

```bash
excel-relay set '<source_id>' --contract './指标.json'
excel-relay run '<source_id>' -p '截止日=2026-10-30' --json
```

指标登记存放在项目的 `.excel-relay/contracts.json`。已声明的日期参数接受 `YYYY-MM-DD`，也继续接受 Excel 序列值。日期按工作簿的日期系统转换。

登记指标后，刷新结果包含 `result_file_path`、`artifact_id` 和 `refreshed_at`。`outputs` 中逐项返回指标的名称、带工作表名的单元格位置、单位和状态。`status=ok` 时，`value` 是本次保存的副本中的原始值。空白、错误值或缺失位置返回 `status=error` 和 `error`。同一组参数保留最新一份副本。

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

开发验收时，可另设 `excel-relay-dev` 入口，命令指向源码目录中的 `.venv/bin/excel-relay`。在 v0.2.0 调用 `configure_source_contract` 登记上述 `date_inputs` 和 `outputs`，再用 `refresh_data_source` 提交参数。两个工具的 `workspace_path` 均传入项目目录的绝对路径。指标从响应的 `outputs` 读取，工作簿副本使用 `result_file_path`。宿主需以能访问同一 Excel 图形会话的用户身份运行。

## 从源码构建

在源码目录使用 `uv 0.11.32` 执行以下命令。

```bash
uv build --offline --wheel
```

## 许可证

Excel Relay 采用 [MIT 许可证](LICENSE)。
