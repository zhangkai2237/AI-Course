# P0—P3 Coze课堂实操资产整体Review

## 一、完整性

| 项目 | 配置清单 | 可复制Prompt | 测试与答案 | JSON蓝图 | 补充材料 | 结果 |
|---|---:|---:|---:|---:|---:|---|
| P0 | 有 | 1份 | 有 | 有 | 统一输入 | 通过 |
| P1 | 有 | 4份 | 有 | 有 | 图片备用路线 | 通过 |
| P2 | 有 | 4份Prompt | 有 | 有 | 2份课程知识材料＋空召回排查 | 通过 |
| P3 | 有 | Agent＋2份模拟工具Prompt | 有 | 有 | 工具定义＋模拟材料 | 通过 |

## 二、JSON检查

以下文件已通过JSON语法校验：

- `共享材料/01_统一正常输入.json`；
- `P0_Prompt/P0_blueprint.json`；
- `P1_Workflow/P1_blueprint.json`；
- `P2_内容RAG/P2_blueprint.json`；
- `P3_Agent/02_工具定义.json`；
- `P3_Agent/P3_blueprint.json`。

所有blueprint均明确标记 `importable:false`，避免被误认为官方一键导入格式。

## 三、四版递进关系

```text
P0：一次Prompt调用
P1：Prompt节点组成固定Workflow
P2：按topic检索课程知识，整理证据并带来源生成
P3：Agent按需调用P2、搜索和模拟保存
```

没有出现P1替代Prompt、P2替代Workflow或P3必须调用所有工具的错误表达。

## 四、参数一致性

P0使用：topic、audience、goal、source_material、tone。

P1使用：

```text
topic
audience
goal
source_material
tone
image_style
page_count
```

P2以及P3调用P2时统一使用：

```text
topic
audience
goal
source_material
target_format
tone_hint
image_style
page_count
```

P3中的 `run_xhs_rag_workflow` 参数与新版P2 Start输入一致。

## 五、边界检查

- P0没有知识、搜索、保存和发布；
- P1增加固定流程，但没有知识、搜索、保存和发布；
- P2使用Start.topic直接检索课程知识，不增加LLM查询改写节点；
- P2正文事实只能来自本次source_material和实际命中的知识片段；
- P2展示真实文档名、证据卡和冲突；
- P2知识未命中时返回knowledge_not_found，不使用模型记忆假装命中；
- P2没有外部搜索、保存和发布；
- P3只有课堂模拟保存，没有真实发布；
- 搜索不可用时使用明确标注为虚构的模拟材料；
- 课堂知识材料不包含当前任职公司、当前业务、客户信息和敏感数字；
- 未使用真实客户、商家和经营数据。

## 六、正式授课前必须人工完成

1. 按配置清单在当前Coze版本中搭建P0—P3；
2. 选择并固定课堂模型；
3. 检查变量绑定和结构化输出能力；
4. 导入K001和K002课程知识材料，并用三个标准问题抽查真实命中；
5. 将P2配置成P3可调用工具；
6. 选择真实搜索工具或搭建模拟搜索Workflow；
7. 搭建课堂模拟保存Workflow；
8. 逐条运行全部测试并保存截图；
9. 基于老师完成版复制学员半成品。

## 七、结论

P0—P3的本地配置资产已完整生成，可以进入Coze实际搭建与平台适配阶段。P2已经改回标准内容RAG：topic负责检索，命中片段提供知识依据，生成结果展示来源，无命中时停止。只有四个项目在当前授课工作区中实际跑通后，才能标记为“老师完成版”。
