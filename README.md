# SYCM Keyword Tool - 生意参谋关键词获取工具

一款基于 Python + Selenium 的淘宝生意参谋搜索词自动提取工具，支持按类目层级批量采集关键词数据并导出为 Excel 文件。

## 功能特性

- **自动化采集** — 通过 Selenium 自动操控 Chrome 浏览器，从生意参谋页面批量提取搜索关键词数据
- **多级类目支持** — 支持一级、二级、三级类目的层级选择与遍历
- **多浏览器会话** — 支持多个浏览器实例并行采集，提升效率
- **智能筛选** — 可配置搜索人气阈值（默认 ≥ 150），自动过滤低价值关键词
- **分页采集** — 自动翻页采集，最多支持 6 页数据
- **Excel 导出** — 采集数据自动保存为 Excel 文件，按类目分类整理
- **暂停/继续** — 采集过程中可随时暂停和继续
- **进度显示** — 实时显示采集进度（一级类目 / 二级类目进度）
- **图形界面** — 基于 tkinter 的可视化操作界面，简单易用

## 截图预览

> 运行程序后会打开"搜索词自动提取工具"窗口界面

## 环境要求

- **操作系统**：Windows
- **Python**：3.10+
- **Chrome 浏览器**：需安装 Google Chrome

## 安装与使用

### 方式一：直接使用 EXE（推荐）

从 [Releases](https://github.com/Assute/sycm-keyword-tool/releases) 页面下载最新的 `.exe` 文件，双击即可运行，无需安装 Python 环境。

### 方式二：从源码运行

1. **克隆仓库**

```bash
git clone https://github.com/Assute/sycm-keyword-tool.git
cd sycm-keyword-tool
```

2. **安装依赖**

```bash
pip install selenium openpyxl pywin32 webdriver-manager
```

3. **运行程序**

```bash
python 生意参谋关键词获取工具.py
```

## 使用步骤

1. 启动程序，在界面中配置 Chrome 驱动路径（程序会通过 `webdriver-manager` 自动下载匹配版本）
2. 登录生意参谋账号（通过程序打开的 Chrome 浏览器窗口）
3. 选择目标类目层级
4. 点击开始采集，程序将自动遍历类目并提取关键词数据
5. 采集完成后，数据自动导出为 Excel 文件

## 依赖列表

| 依赖包 | 用途 |
|--------|------|
| `selenium` | 浏览器自动化控制 |
| `openpyxl` | Excel 文件读写 |
| `pywin32` | Windows COM 组件交互 |
| `webdriver-manager` | 自动管理 ChromeDriver 版本 |

## 项目结构

```
keyword-tool/
├── 生意参谋关键词获取工具.py   # 主程序源码
├── README.md                   # 项目说明文档
└── LICENSE                     # 开源协议
```

## 配置说明

以下参数可在源码中调整：

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `min_popularity_threshold` | 150 | 搜索人气筛选阈值 |
| `max_pages` | 6 | 每个类目最大采集页数 |
| `base_debug_port` | 9000 | Chrome 调试端口起始值 |
| `exclude_level1_serials` | `[4,34,52,53,54,58,59,60]` | 排除的一级类目序号 |

## 免责声明

本工具仅供学习和研究使用。请遵守相关平台的使用条款和规定，合理使用本工具。因使用本工具产生的任何问题，作者不承担任何责任。

## 开源协议

本项目基于 [CC BY-NC 4.0](LICENSE)（创意共享-署名-非商业性使用 4.0 国际）协议发布。

- 你可以自由查看、使用、修改和分享本项目代码
- **禁止用于商业用途**
- 使用时需注明原作者和出处

## 作者

**Assute**

- GitHub：[https://github.com/Assute](https://github.com/Assute)
- Gitee：[https://gitee.com/Assute](https://gitee.com/Assute)
