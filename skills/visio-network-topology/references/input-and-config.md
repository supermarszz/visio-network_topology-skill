# 输入与配置

## XLSX / CSV

每行代表一条物理链路，列名须包含以下七列，列顺序可以不同：

| 本端设备名 | 本端设备类型 | 本端接口 | 对端设备名 | 对端设备类型 | 对端接口 | 线路类型 |
|---|---|---|---|---|---|---|
| Core-A | Core-switch | E1/1 | Access-A | Switch | E1/47 | 10G光纤 |

默认选择第一个非空工作表；建议明确传入 `--sheet`。空行忽略。公式单元格拒绝读取，先导出为值，避免无缓存或陈旧缓存造成错误连接。

类型默认映射：`Core-switch / 核心交换机` → core / Core-Switch；`Switch / 交换机 / 接入交换机` → access / Switch；`Firewall / 防火墙` → firewall / Firewall。忽略类型大小写。模具内 Master 匹配大小写后使用真实名称。

线路接受 `40G / 40G光纤 / 40Gbps`，`10G / 10G光纤 / 10Gbps`，`网线 / 铜缆 / copper / ethernet`。未知类型必须显式映射，不能擅自认作 10G。

同一物理链路的重复或反向记录只画一次，全部源行号记录在 `source_rows` 中，重复行会列在报告里。不同接口的并行链路始终单独保留。单个物理端口连接不同对端、设备类型冲突、自连接、缺字段均报错。

## JSON 输入

输入为数组，或带 `rows` / `links` 数组的对象。字段为 `from, from_type, from_port, to, to_type, to_port, line_type`；可带 `source_row`。

## 可选配置

`plan --config config.json`，长度单位为 Visio 英寸。坐标 `cx/cy` 为设备中心，Y 轴向上。

```json
{
  "title": "生产网络拓扑图",
  "page_width": 36,
  "page_height": 21,
  "device_types": {
    "核心": {"role": "core", "master": "Core-Switch"},
    "接入": {"role": "access", "master": "Switch"}
  },
  "line_types": {"40G DAC": "40G"},
  "positions": {
    "Core-A": {"cx": 11, "cy": 11.9, "w": 3.4, "h": 2.7862}
  }
}
```

支持的层角色为 `firewall / core / access`。新设备类型可以映射到现有角色及相应模具；其他层级或特殊布局需要扩展脚本，不要宣称已经完整支持。

`positions` 用于局部修正。保持同层 `cy` 一致，增大高密度节点尺寸时保留图标比例。使用超长设备名时应增宽名称框并进行视觉检查；v0.1 不自动推断机房、业务域或设备分组语义。

## 边界

- 无自动 OCR、图片反推、任意现有 VSDX 增量编辑或 RTL 解析。
- 默认布局针对防火墙、核心与接入交换机的常见网络。密集网状网络可能需要手动调整坐标。
- 不随仓库分发 Visio、pywin32 二进制、用户 Excel、第三方图标或模具。用户自行提供有权使用的 `.vssx`。
