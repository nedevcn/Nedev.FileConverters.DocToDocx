# Nedev.FileConverters.DocToDocx

一个面向 .NET 的 DOC 到 DOCX 转换器，当前代码基线以“可生成有效 DOCX、优先恢复可读内容”为目标，而不是追求对旧版 Word 二进制格式的完全等价复刻。

## 当前完成度

基于现有源码和测试，项目可以判断为“可用的早期版本”，不是纯概念验证，但也还没有到“复杂旧文档高保真转换”的阶段。

### 已完成且可直接使用

- 提供文件路径和流两套转换 API。
- 提供同步、异步、进度回调、取消、告警收集等调用方式。
- 可以识别输入是 DOC 还是 DOCX，DOCX 会直接透传复制。
- 可以输出结构合法的 DOCX 包，并提供包校验辅助方法。
- 提供 CLI，可处理单文件，也可批量处理目录。
- 已接入 `Nedev.FileConverters.Core` 的 `IFileConverter` 接口。

### 已有实现，但属于“最佳努力恢复”

- 旧版 DOC 二进制解析链路已经搭起来，包括 CFB、FIB、CLX/Piece Table、FKP、SPRM、样式、段落与 run 解析。
- 基础文本、常见 run/段落格式、超链接字段、书签、批注、脚注、页眉页脚、列表和部分主题色映射已有代码与测试支撑。
- 表格、嵌套表格、图片、OfficeArt 形状、文本框、OLE 对象、图表、公式对象都已进入解析或写出链路，部分能力已有真实样本与回归测试支撑，但复杂场景下仍有不少地方依赖启发式恢复。
- OfficeArt/FSPA 已能恢复更多浮动对象锚点信息、环绕模式，并在部分自定义几何场景下把 wrap polygon 保留到输出模型。
- 文本框读取链路已开始把 textbox story 与 textbox shape 元数据合并，能保留更多位置、尺寸、环绕与基础对齐信息。
- 文本框匹配已不再只依赖简单顺序，开始结合主文档中的 textbox 锚点字段位置和段落提示来关联 textbox story 与 OfficeArt textbox shape。
- textbox 锚点读取已进一步从单点 CP 提升为 field begin/separate/end 边界重建，后续可继续在此基础上细化 textbox 归属判定。
- `samples/sample1.doc` 已纳入真实回归样本，当前覆盖加载、带 warning 的完整转换、标题/内联格式/表格/drawing 输出，以及 legacy field code 中非法 XML 控制字符的清洗。
- BIFF 图表扫描除基础数据恢复外，已开始补充部分布局偏好，例如条形图轴位置和可推断的类目顺序方向。
- 支持 XOR/RC4 相关加密读取路径，但还缺真实样本驱动的端到端回归验证。
- 一批历史上容易静默降级的二进制结构现在已补上显式边界校验和 warning，包括 bookmark/annotation/textbox PLC、STTBF、字体表、样式表和 section PLC；对应的 malformed synthetic regression 测试也已补充。

### 目前不能承诺的能力

- 不保证复杂版式、复杂图形、复杂图表在 Word 中高保真还原。
- 不保证损坏文档、截断流、异常 OfficeArt/OLE 数据都能稳定恢复。
- 不保证与源 DOC 在结构上 round-trip 等价。
- 不保证所有主题、域、修订、控件、SmartArt、宏相关对象都被完整建模。

## 代码里已经明确存在的能力

下面这些能力是可以从当前入口代码和测试中直接确认的。

- `DocToDocxConverter` 提供文件/流转换、异步转换、进度回调、带告警结果的转换，以及保存/加载/校验 DOCX 包的辅助方法。
- `DocToDocxFileConverter` 可作为 `IFileConverter` 被宿主框架发现和调用。
- CLI 支持以下行为：
  - 单文件转换。
  - 目录批量转换。
  - `-r` 递归处理子目录。
  - `-p` 指定密码。
  - `--no-hyperlinks` 禁止生成超链接关系。
  - `-h`/`--help` 与 `-v`/`--version`。
- 测试已经覆盖的一些关键点：
  - DOCX 透传复制。
  - 流式转换与非可 seek 输入。
  - 进度事件与诊断采集。
  - 包结构校验。
  - 文档写出中的基础段落、格式、批注、页眉页脚、脚注、字段、部分图表 XML。
  - 图表扫描器对基础 BIFF 数值和标签恢复。
  - 图表扫描器对部分轴布局提示和类目顺序方向恢复。
  - OfficeArt、SPRM、主题读取、RC4 流等核心辅助模块。
  - 已增加一批 parser 到 writer 的回归测试，覆盖 textbox shape 合并、wrap polygon 输出和图表布局提示传递。

当前测试项目下有大量针对 writer 和底层解析器的单元测试，也已经补上部分“解析模型 -> 写出 OOXML”的集成回归；但端到端真实样本文档覆盖仍然偏少，尤其是加密、损坏、复杂排版、复杂 OLE/图表这几类。

## 适合当前版本的使用场景

- 批量把普通历史 DOC 文档迁移到可继续编辑的 DOCX。
- 服务端或工具链中做“尽量恢复内容”的自动化转换。
- 对文本、基础格式、常见批注脚注、基础表格结构的保留要求高，但接受部分版式近似。

## 暂不建议过度承诺的场景

- 法务、审计、档案类场景下要求像素级版式一致。
- 大量依赖复杂图形、复杂图表、嵌套对象、控件、宏的 DOC 文档。
- 对损坏文档修复能力有严格 SLA 的场景。

## 使用方式

### 类库调用

```csharp
using Nedev.FileConverters.DocToDocx;

DocToDocxConverter.Convert("input.doc", "output.docx");

DocToDocxConverter.Convert(
    "input.doc",
    "output.docx",
    password: null,
    enableHyperlinks: true);

using var input = File.OpenRead("input.doc");
using var output = File.Create("output.docx");

DocToDocxConverter.Convert(input, output);
```

### 异步与进度

```csharp
using Nedev.FileConverters.DocToDocx;

var progress = new Progress<ConversionProgress>(update =>
{
    Console.WriteLine($"[{update.PercentComplete,3}%] {update.Stage}: {update.Message}");
});

await DocToDocxConverter.ConvertAsync(
    "input.doc",
    "output.docx",
    progress,
    password: null,
    enableHyperlinks: true,
    cancellationToken: CancellationToken.None);
```

### 获取非致命告警

```csharp
using Nedev.FileConverters.DocToDocx;

var result = DocToDocxConverter.ConvertWithWarnings("input.doc", "output.docx");

foreach (var diagnostic in result.Diagnostics)
{
    Console.WriteLine($"[{diagnostic.Level}] {diagnostic.Message}");
}
```

### CLI

```bash
Nedev.FileConverters.DocToDocx.Cli <input.doc|inputDir> <output.docx|outputDir> [-p <password>] [-r] [--no-hyperlinks]
```

参数说明：

- `<input.doc|inputDir>` 输入 DOC 文件或目录。
- `<output.docx|outputDir>` 输出 DOCX 文件或目录。
- `-p`, `--password` 加密 DOC 的密码。
- `-r`, `--recursive` 目录模式下递归处理。
- `--no-hyperlinks` 输出普通文本，不生成超链接关系。
- `-h`, `--help` 显示帮助。
- `-v`, `--version` 显示版本。

CLI 退出码：

- `0` 成功。
- `1` 转换失败或运行时异常。
- `2` 参数错误。

## 目标框架

- 类库：`net8.0`、`netstandard2.1`
- 测试：`net10.0`

## 测试

在仓库根目录执行：

```bash
dotnet test
```

当前测试更偏向“模块正确性”和“生成包有效性”，而不是“真实世界 DOC 样本大规模兼容性”。因此，测试通过不代表所有旧文档都能高质量还原。

## 下一阶段开发计划

### 第一优先级

1. 补齐真实加密 DOC 样本回归，重点验证 RC4 边界、错误密码、表流和数据流解密路径。
2. 给最佳努力恢复路径补结构化诊断，减少静默降级，尤其是 OLE、图表、形状、图片和截断流解析。
3. 建立一批真实 DOC 回归样本，覆盖中文文档、复杂表格、页眉页脚、批注、脚注、图文混排。

### 第二优先级

1. 继续深挖 OfficeArt 与 textbox 对齐关系，减少 textbox story 与 shape 元数据只能按顺序合并的启发式处理。
2. 扩大图表恢复范围，把标题、图例、轴标签、更多 BIFF 记录恢复从“可编辑占位图”推进到“保留更多源信息”。
3. 补齐恶意或损坏输入的防御性测试，覆盖 FKP、SPRM、CLX、OLE 和 OfficeArt。

### 第三优先级

1. 强化复杂表格和嵌套表格恢复，补异常 cell 边界和 merge 场景测试。
2. 评估是否需要公开更明确的“支持矩阵”与“已知不兼容清单”。
3. 在有足够样本前，不建议把项目版本含义提升为高保真生产级转换器。
