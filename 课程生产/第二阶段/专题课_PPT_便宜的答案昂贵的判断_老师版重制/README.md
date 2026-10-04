# 专题课老师版 v2.0

本工作区从空白 16:9 演示文稿生成《便宜的答案，昂贵的判断》老师版 v2.0。

## 内容与安全边界

- 唯一内容源：`../07_专题课_便宜的答案昂贵的判断_前五节老师讲义_v1.0.md`
- 构建不读取任何旧 `.ppt` / `.pptx`，也不使用旧专题课目录作为输入。
- 公共客服案例统一标注“公共模拟案例”；无真实用户、生产连接或经营成效。
- 本人项目部分只提供条件式模板，须由学员替换为本人可举证事实。
- 不修改讲义母稿；讲义中的真实上线冲突在演示文稿中改为条件式、安全口径。

## 结构

- `src/content.mjs`：44页内容层与120分钟时长
- `src/theme.mjs`：浅色“产品评审工作台”主题
- `src/components.mjs`：通用原生形状组件
- `src/wireframes.mjs`：P21—P23可编辑低保真UI
- `src/notes.mjs`：老师备注结构
- `src/source-policy.mjs`：唯一来源与口径策略
- `src/build.mjs`：PptxGenJS 4.0.1 构建入口
- `src/qa.mjs`：静态QA
- `dist/`：正式PPTX与可选PDF
- `qa/`：构建清单与静态QA报告
- `rendered/`：PDF逐页PNG与联系表

## 构建

仓库根目录已安装 `pptxgenjs@4.0.1`。本工作区依赖声明固定为4.0.1。

```bash
node src/build.mjs
node src/qa.mjs
```

或：

```bash
npm run all
```

## 可选渲染

本次构建中，PowerPoint 可以打开并识别44页，但 AppleScript 的 PDF 导出返回 `-9074` 或未产生文件，因此**没有伪称为 PowerPoint 导出成功**。为完成视觉QA，改用 Keynote 从最终PPTX导出44页PDF；该PDF仅作兼容性预览与渲染检查。

生成路径：`dist/专题课_便宜的答案昂贵的判断_老师版_v2.0.pdf`

随后使用共享工具渲染PNG：

```bash
swift ../../_tools/render_pdf_pages.swift dist/专题课_便宜的答案昂贵的判断_老师版_v2.0.pdf rendered/pages
```

并生成联系表：

```bash
python3 ../../_tools/make_ppt_contact_sheets.py rendered/pages rendered/contact-sheet
```

## 人工检查

静态QA不能替代 PowerPoint 中的逐页视觉检查。请重点检查：字体回退、P21—P23低保真UI、表格密度、演讲者备注显示、动画/切换（本版本未依赖动画）以及PDF导出后的分页与裁切。
