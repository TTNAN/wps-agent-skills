# WPS 演示版式 ID

PowerPoint 系尺寸单位是**磅**（72 磅 = 1 英寸）。

| ID | 含义 | 备注 |
|---|---|---|
| 1 | ppLayoutTitle（标题幻灯片） | 标题页用这个 |
| 2 | ppLayoutTitleAndContent | 模板相关，可能不一致 |
| 5 | ppLayoutTitleOnly | 模板相关，可能不一致 |
| 12 | ppLayoutBlank（空白） | **最稳**：WPS 模板版式编号可能和 PowerPoint 不一致，拿不准就用空白 + 手写形状 |

版式探测失败时的回退策略：空白版式 + `AddTextbox(1, left, top, width, height)`（1 = msoTextBox），永远可用。
