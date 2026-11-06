# 发布与维护说明

给 fork / 二次发布的人看的注意事项。日常使用不需要读这个文件。

当前仓库：<https://github.com/tttkkk123/wordpdf>

---

## 一、提交哪些文件

**25 个文件、约 300 KB**，全部在仓库根目录下：

| 路径 | 作用 |
|------|------|
| `wordpdf/`（16 个） | 转换核心：`core/` 下的 5 个引擎与编排层、`gui.py`、`cli.py`、`addin/WordPdfAddin.bas` |
| `启动转换器.bat` + `run_gui.pyw` | 一键启动（bat 找环境 → pyw 拉起界面） |
| `requirements.txt` + `pyproject.toml` | 依赖与安装声明 |
| `README.md` `LICENSE` `.gitignore` `.gitattributes` | 文档与仓库规则 |
| `docs/screenshot.png` + `docs/PUBLISHING.md` | README 截图与本文件 |
| `tests/selftest.py` + `tests/gui_smoke.py` | 端到端自检、界面冒烟 |

## 二、绝对不要提交

| 排除项 | 原因 |
|--------|------|
| `.venv/` | 上百 MB 二进制，且 `pyvenv.cfg` 里写着本机绝对路径，换机器就废 |
| `__pycache__/` `*.pyc` | 编译缓存 |
| `logs/` `tests/_out/` | 运行产物、测试中间文件 |
| `*.egg-info/` `build/` `dist/` | 打包中间产物 |
| `.workbuddy/` | 本地工具状态 |
| 任何写死个人路径的文件 | 泄露目录结构 |

`.gitignore` 已经全部覆盖，`git status --ignored` 可以核对。

## 三、两个容易踩的坑

### 1. `.bat` 必须是 CRLF，且禁止 Git 转换换行

之前就因为这个翻过车：bat 用 LF 保存，`cmd.exe` 解析时会把 `python` 截成 `ython`，
双击后窗口一闪而过（报「不是内部或外部命令」）。

`.gitattributes` 里因此写了：

```gitattributes
*.bat  -text      # 完全禁用换行转换，clone / raw 下载 / zip 下载都是 CRLF
*.cmd  -text
```

**验证方法**（克隆后执行，输出里必须能看到 `\r \n`）：

```bash
git clone https://github.com/tttkkk123/wordpdf.git
cd wordpdf
od -c 启动转换器.bat | head -2
```

另外：**不要用 GitHub 网页编辑器改这个 bat**，网页保存会改成 LF + UTF-8，改完就不能双击了，只在本机改完再 push。

### 2. bat 里不要写中文

批处理文件如果用 GBK 编码，GitHub 网页上会显示成乱码；如果用 UTF-8，中文 Windows 的
cmd 又可能乱码。所以 `启动转换器.bat` 的提示语一律用 ASCII 英文，
中文只出现在 Python 写的界面里（GUI 与日志都是 UTF-8，不受控制台代码页影响）。

## 四、Word 加载项（VBA）的路径

`wordpdf/addin/WordPdfAddin.bas` 顶部两行是占位路径，发布前确认没写自己的本机路径：

```vb
Private Const PROJECT_DIR As String = "C:\path\to\wordpdf"
Private Const PY_EXE As String = "C:\path\to\wordpdf\.venv\Scripts\python.exe"
```

使用者导入后需要自己改成实际路径；改错时宏会弹窗提示，而不是静默失败。

## 五、日常更新流程

```bash
cd wordpdf
# 改完代码后
git add -A
git status --short          # 确认没有误加 .venv / 产物
git commit -m "fix: 描述改动"
git push
```

首次在本机配置过身份即可（本仓库使用的是）：

```bash
git config user.name  "tttkkk123"
git config user.email "267032926+tttkkk123@users.noreply.github.com"
```

## 六、发布前验收清单

1. `git clone` 到空目录，双击 `启动转换器.bat`：能自动建环境 → 装依赖（1–3 分钟）→ 弹出界面；
2. 在该克隆里跑 `tests/selftest.py`：Word→PDF、PDF→Word、两个兜底引擎全部通过；
3. `od -c 启动转换器.bat | head -2` 能看到 `\r \n`；
4. 仓库首页 README 与截图正常显示（`docs/screenshot.png` 返回 200）。

## 七、重新创建仓库（如果需要换账号或改名）

```bash
gh repo create <名字> --public --source=. --remote=origin --push \
  --description "Windows 桌面工具：Word ⇄ PDF 双向转换。混合引擎 + 自动降级"
gh repo edit --add-topic word-to-pdf --add-topic pdf-to-word --add-topic docx \
  --add-topic pdf --add-topic python --add-topic tkinter
```
