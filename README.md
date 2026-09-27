# Paper Format Modification Tool

项目当前包含可独立运行的格式处理核心，以及 Spring Boot 文件处理服务。它使用 Java 17 目标字节码和 docx4j，根据学校模板检查并修正 DOCX 论文格式。

详细接口见 `docs/API.md`，后续计划见 `docs/ROADMAP.md`。

## 目录

```text
paper-format-modification-tool/
├─ backend/
│  ├─ core/                 # Word 解析与修改核心
│  └─ server/               # Spring Boot 上传、任务和下载接口
├─ frontend/                # Vue 3 + Element Plus 操作页面
├─ build-app.ps1            # 构建完整应用
├─ run-app.ps1              # 一键启动完整应用
├─ run-demo.ps1             # 命令行 Demo
└─ run-server.ps1           # 启动后端服务
```

## 当前能力

- 读取模板和待处理 `.docx`
- 同时提取模板页面设置、自定义样式、文字格式要求和示例段落，并在报告中记录每条规则的来源证据
- 先划分封面、中英文摘要、目录、正文、参考文献、致谢和附录，再识别各区域内部结构
- 上传后只生成带识别依据和置信度的格式计划；用户在二次确认页面勾选规则并最终确认后，系统才会修改文档
- 识别正文一级至四级标题、正文、图题、表题、续表、参考文献条目、致谢正文和章节题名
- 图题即使位于无边框表格中也会按模板的“论文图号图名”样式处理，并尽量与对应图片保持同页
- 将缺失章标题、重复标题编号等结构问题列入二次确认，只有用户勾选后才修改文字
- 按模板对应组件修复完整封面，包括说明书标题、课题名称、信息表、段落和单元格格式
- 中文/英文摘要、关键词分别处理，支持同一段中“摘要/Abstract/关键词/Keywords”标签与正文采用不同字符格式
- 修复目录标题与一/二级目录样式；目录保留为 Word 的真实 `TOC` 域，只显示一级和二级标题，并把目录标题本身排除在目录条目之外
- 分别处理参考文献标题与条目、致谢标题与正文、附录标题，不再把这些区域当成普通正文
- 按名称同步模板中的受管样式定义，同时保留原文、图片、表格、页眉页脚和关系部件
- 输出新的 `.docx` 和 JSON 修改报告
- 校验原文件哈希未变化，并重新加载输出 DOCX

## 运行

在项目根目录执行下面的命令即可使用默认文件名运行：

```powershell
.\run-demo.ps1
```

默认读取项目上一级目录中的模板和 `说明书.docx`，并写入上一级目录的 `outputs`。也可以显式传入四个路径：

```powershell
.\run-demo.ps1 -Template "D:\Paper Format Modification Tool\6-0毕业设计说明书_模板_20251016.docx" -Source "D:\Paper Format Modification Tool\说明书.docx" -Output "D:\Paper Format Modification Tool\outputs\formatted-manual.docx" -Report "D:\Paper Format Modification Tool\outputs\format-report.json"
```

底层 Maven 方式如下：

在 `backend/core` 目录执行：

```powershell
mvn -q compile exec:java '-Dexec.args=<模板.docx> <原说明书.docx> <输出.docx> <报告.json>'
```

路径含空格时，请把每个路径放在双引号中。

## 当前边界

当前版本会保留目录为 Word 自动目录域并设置打开文档时刷新。服务端没有安装 Microsoft Word 时不会预先计算新的目录页码，但用户在 Word 中打开文件后即可更新；当前验证文件已在 Word 中完成刷新。公式、复杂域、交叉引用和未经用户确认的论文内容不会自动改写；复杂对象会原样保留并在报告中列为人工复核项。

## 启动后端服务

在项目根目录执行：

```powershell
.\run-server.ps1
```

默认端口为 `8080`，任务文件保存在项目上一级的 `storage` 目录。接口如下：

- `GET /api/v1/health`：服务状态
- `POST /api/v1/tasks`：上传 `template` 和 `document` 两个 DOCX 文件
- `GET /api/v1/tasks/{taskId}`：查询任务状态
- `GET /api/v1/tasks/{taskId}/analysis`：读取待确认的模板格式计划
- `POST /api/v1/tasks/{taskId}/confirm`：确认启用的规则并开始修改
- `GET /api/v1/tasks/{taskId}/result`：下载修改后的 DOCX
- `GET /api/v1/tasks/{taskId}/report`：下载修改报告

## 一键启动完整应用

在项目根目录执行：

```powershell
.\run-app.ps1
```

首次运行会安装前端依赖并构建应用。看到启动成功提示后，在浏览器访问 `http://127.0.0.1:8080`。后续如果文件没有变化，可以使用 `./run-app.ps1 -SkipBuild` 快速启动。启动脚本会自动选择最新构建的 `paper-format-server-*.jar`。前后端作为同一应用一起构建和升级，不在运行时进行版本校验。

## 启动前端

先启动后端，然后在 `frontend` 目录执行：

```powershell
npm install
npm run dev
```

浏览器访问 `http://127.0.0.1:5173`。开发服务器会把 `/api` 请求转发到本机 `8080` 端口。
