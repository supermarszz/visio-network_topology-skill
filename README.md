# visio-network_topology-skill

**v0.1** — 从 Excel / CSV 链路表生成可编辑的 Microsoft Visio 网络拓扑图。

这个 Skill 将链路解析、分层排版、专业模具加载、接口标签避让和保存后校验整理成可复用流程。默认布局为防火墙在上、双核心居中、接入交换机在下，同层设备水平对齐。

## 功能

- 40G 红色直线、10G 黄色直线、网线黑色直线，保留不同接口的平行链路。
- 左上角纵向图例，顺序为 40G、10G、网线，不标注数量。
- 每条线路两端分别标注本端接口。标签中心在线上，与线路同向旋转，白底和同色细边框帮助识别归属。
- 自动扩大接口密集的核心图标、分散端点；标签避开设备、名称、其他标签以及其他线路。
- 防火墙成对紧凑排列。画布宽度随接入设备数量增长，支持配置设备位置。
- 保存后检查实际 VSDX 中的直线几何、端点 Glue、标签文本和位置、线路颜色、图例位置。

## 安装为 Skill

仓库名保留用户指定的下划线；可安装的 Skill 名称使用 `visio-network-topology`。

```powershell
git clone https://github.com/supermarszz/visio-network_topology-skill.git
Copy-Item -Recurse .\visio-network_topology-skill\skills\visio-network-topology "$env:USERPROFILE\.codex\skills\visio-network-topology"
python -m pip install -r .\visio-network_topology-skill\requirements.txt
```

随后在支持 Skills 的工具中使用：

> 使用 $visio-network-topology 根据链路表绘图。防火墙放上面，同层设备水平对齐。所有连接使用直线，接口名称放在对应线路上并靠近设备。

安装到其他 Agent 时，将 `skills/visio-network-topology` 复制到它的 Skills 目录即可；脚本不依赖 Codex 专有运行时。

也可从 `v0.1` Release 下载安装 ZIP，将其中的 `visio-network-topology` 文件夹解压到 Skills 目录。该文件夹自带 `requirements.txt`。维护者运行 `python tools/package_skill.py` 可重新构建安装包。

## 独立运行

规划需要 Python 3.10+；读取 XLSX 需要 `openpyxl`。渲染还需要 **Windows、本机 Microsoft Visio、pywin32**。Linux/macOS 可以运行解析、布局和测试，但不能通过 COM 渲染 Visio。

```powershell
python skills/visio-network-topology/scripts/topology.py --version
python skills/visio-network-topology/scripts/topology.py masters "网络模具.vssx"
python skills/visio-network-topology/scripts/topology.py plan "链路.xlsx" --sheet "Sheet1" --stencil "网络模具.vssx" --output outputs/topology.json
python skills/visio-network-topology/scripts/topology.py render outputs/topology.json --output outputs/topology.vsdx --preview outputs/topology.png --report outputs/validation.json
python skills/visio-network-topology/scripts/topology.py validate outputs/topology.json --vsdx outputs/topology.vsdx
```

没有自定义模具时省略 `--stencil`，将使用基本矩形。现有输出默认不会覆盖，明确需要重绘时传入 `render --overwrite`。

使用仓库中人工编写的演示清单：

```powershell
python skills/visio-network-topology/scripts/topology.py plan examples/links.csv --output outputs/example.json
python -m unittest discover -s tests -v
```

## 输入格式

七列必需，顺序可以不同：

| 本端设备名 | 本端设备类型 | 本端接口 | 对端设备名 | 对端设备类型 | 对端接口 | 线路类型 |
|---|---|---|---|---|---|---|
| Core-A | Core-switch | E1/1 | Access-A | Switch | E1/47 | 10G光纤 |

默认设备类型为 `Core-switch`、`Switch`、`Firewall` 及对应中文名称。通过 JSON 配置可扩展类型到 `core / access / firewall` 三种层角色。重复、反向记录会合并为同一物理链路并报告源行号；不同接口的并行连接不会合并。

详细列约定、配置格式、重复处理和边界见 [输入与配置](skills/visio-network-topology/references/input-and-config.md)。Agent 的工作流程见 [SKILL.md](skills/visio-network-topology/SKILL.md)。

## 设计过程沉淀

1. 从链路表提取设备、类型、接口、速率，先核对数量与端口冲突。
2. 按层级确定行，再计算画布与设备间距；避免设备重叠及直线穿过无关设备。
3. 通过真实 Master 名称创建专业图标，使用原生单线段并绑定端点。
4. 将组合的接口说明拆成两个端点标签；标签靠近设备但不能漂离所属线路。
5. 高密度区扩大图标、分散端点，沿线路搜索不遮挡其他线的标签位置。
6. 保存、重新读取 VSDX 校验，并查看全图与核心接口局部。仅凭“文件生成成功”不能判定排版完成。

## v0.1 范围与限制

面向常见的防火墙、核心和接入网络，支持数据驱动的新建图表。高度网状拓扑可能需要配置坐标；当无法安全放置标签时脚本报错，不隐藏链路或缩小文字。普通的线路交叉仍可能存在，不表示设备连接。

不包含图片反推、完整通用图布局、RTL 解析或任意现有 VSDX 的增量编辑。移动生成图的设备后，直线和标签位置公式会跟随，但需要重新检查新的拥挤和交叉。

本仓库不包含用户原始 Excel、真实输出图、第三方模具、Visio 或二进制依赖。请自行提供有权使用的 `.vssx`。代码独立封装，绘图思路参考了 [supermarszz/visio-skill](https://github.com/supermarszz/visio-skill) 的 JSON + Visio COM 工作流，没有打包该仓库代码或图标资产。
