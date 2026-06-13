> [English](./TOOLS.md) | [中文](./TOOLS.zh-CN.md)

# Word MCP Server — 工具参考

> 108 个工具，涵盖 12 个模块。

---

### 📂 文档生命周期

| 工具 | 描述 |
|------|------|
| `word_get_status` | 查询当前 Word/文档状态（NO_WORD/NO_DOC/DOC_ACTIVE/DIALOG） |
| `word_document` | 通用文档入口：按路径打开或使用当前文档 |
| `word_open` | 打开已有 .docx 文件 |
| `word_save` | 保存当前文档（自动 .bak 备份） |
| `word_save_as` | 另存为新文件（自动适配格式） |
| `word_close` | 关闭当前文档 |
| `word_get_info` | 获取文档统计信息（字数/页数/段落数） |
| `word_get_structure` | 获取文档标题大纲（含段落索引） |
| `word_quit` | 退出 Microsoft Word |

### ✏️ 内容输入

| 工具 | 描述 |
|------|------|
| `word_insert_paragraph` | 插入段落换行 |
| `word_insert_page_break` | 插入分页符 |
| `word_insert_horizontal_line` | 插入水平分割线 |
| `word_insert_list` | 插入项目符号或编号列表 |
| `word_set_columns` | 设置当前节的分栏数 |
| `word_insert_section_break` | 插入分节符 |
| `word_insert_file` | 在光标处插入另一个 .docx 文件内容 |
| `word_backspace` | 删除光标前字符 |
| `word_get_text` | 读取全文文本 |
| `word_get_paragraph` | 读取指定段落文本 |
| `word_get_table_data` | 提取表格内容为结构化数据 |
| `word_get_comments` | 列出所有批注及作者信息 |
| `word_get_bookmarks` | 列出所有书签 |
| `word_get_lists` | 列出所有列表（含层级） |
| `word_get_sections` | 列出所有节及页面设置 |
| `word_export_to_pdf` | 导出为 PDF |

### 📤 流式文档（推荐内容写入方式）

| 工具 | 描述 |
|------|------|
| `word_stream_start` | 启动流式会话：创建文档，配置页面设置、基础样式（`baseStyleProfile` 支持按样式名预设字体/段落格式）和模板 |
| `word_stream_block` | 写入 Markdown 内容块（Word 中即时预览），支持标题/粗体/斜体/列表/表格/代码块/引用/链接 |
| `word_stream_end` | 结束流式会话：保存文档，可选导出 PDF，返回块数/字符数/耗时 |

### 🔗 链接、引用与批注

| 工具 | 描述 |
|------|------|
| `word_add_hyperlink` | 添加超链接 |
| `word_add_footnote` | 添加脚注（自动编号） |
| `word_add_comment` | 添加批注 |

### 🔍 查找、选区与导航

| 工具 | 描述 |
|------|------|
| `word_find_text` | 查找文本并定位光标 |
| `word_find_replace` | 查找替换文本 |
| `word_go_to` | 跳转到指定位置（页/节/书签/末尾） |
| `word_go_to_paragraph` | 按段落索引跳转 |
| `word_select_all` | 全选 |
| `word_select_text` | 按字符范围选择 |
| `word_select_current_word` | 选择光标所在词 |
| `word_select_current_paragraph` | 选择光标所在段落 |
| `word_delete` | 删除选中内容 |

### 📋 剪贴板与撤销

| 工具 | 描述 |
|------|------|
| `word_copy` | 复制到剪贴板 |
| `word_cut` | 剪切到剪贴板 |
| `word_paste` | 从剪贴板粘贴 |
| `word_undo` | 撤销上一步操作 |
| `word_redo` | 重做已撤销的操作 |
| `word_get_cursor_info` | 获取光标位置/选区信息 |

### 🎨 格式编排

| 工具 | 描述 |
|------|------|
| `word_set_font` | 设置字体（名称/大小/加粗/斜体/颜色/删除线/上标/下标） |
| `word_set_paragraph` | 设置段落格式（对齐/缩进/间距/行距） |
| `word_apply_style` | 应用命名样式 |
| `word_set_page_setup` | 页面设置（边距/方向/纸张） |
| `word_set_properties` | 设置文档属性（标题/作者/主题/关键词） |
| `word_list_styles` | 列出可用样式 |
| `word_set_body_indent` | 为所有"正文"段落设置首行缩进（中文学术论文） |
| `word_set_track_changes` | 开启/关闭修订模式 |
| `word_accept_changes` | 接受文档中所有修订 |
| `word_reject_changes` | 拒绝文档中所有修订 |

### 📊 表格

| 工具 | 描述 |
|------|------|
| `word_insert_table` | 插入表格（可选填充数据） |
| `word_edit_cell` | 修改单元格内容 |
| `word_edit_cells` | 批量填充多个单元格 |
| `word_add_table_row` | 添加行 |
| `word_delete_table_row` | 删除行 |
| `word_add_table_column` | 添加列 |
| `word_delete_table_column` | 删除列 |
| `word_set_table_borders` | 设置表格边框样式/颜色/宽度 |
| `word_set_table_shading` | 设置表格/首行背景色 |
| `word_merge_table_cells` | 合并单元格 |
| `word_set_column_width` | 设置列宽 |
| `word_set_row_height` | 设置行高 |
| `word_set_cell_font` | 设置单元格字体 |
| `word_set_cell_vertical_alignment` | 设置单元格垂直对齐 |
| `word_apply_table_style` | 应用内置 Word 表格样式 |
| `word_table_to_text` | 表格转纯文本 |
| `word_text_to_table` | 分隔文本转表格 |

### 🖼️ 媒体

| 工具 | 描述 |
|------|------|
| `word_insert_image` | 插入图片文件 |
| `word_insert_chart` | 插入图表（column/bar/line/pie/area） |
| `word_insert_textbox` | 插入文本框 |

### 📄 文档结构

| 工具 | 描述 |
|------|------|
| `word_set_header` | 设置页眉 |
| `word_set_footer` | 设置页脚 |
| `word_set_page_numbers` | 插入页码 |
| `word_insert_toc` | 插入目录 |
| `word_add_bookmark` | 添加书签 |
| `word_set_watermark` | 设置/移除水印 |

### 🏗️ Manager 层（高层 API）

| 工具 | 描述 |
|------|------|
| `word_mgr_set_header` / `word_mgr_set_footer` | 设置页眉/页脚（自动返回正文） |
| `word_mgr_set_page_numbers` | 插入页码（自动返回正文） |
| `word_mgr_insert_table` / `word_mgr_insert_chart` / `word_mgr_insert_image` | 插入元素（自动段落分隔） |
| `word_mgr_insert_textbox` | 插入文本框（自动复位光标） |
| `word_mgr_add_bookmark` / `word_mgr_add_comment` / `word_mgr_add_footnote` / `word_mgr_add_hyperlink` | 添加标记（自动深度复位光标） |
| `word_mgr_insert_section_break` | 分节符（自动返回正文） |
| `word_mgr_set_watermark` | 水印（自动返回正文） |

### 🧭 语义定位

| 工具 | 描述 |
|------|------|
| `word_locate` | 解析标题/段落/表格/书签位置（只读） |
| `word_select_at` | 按语义位置移动光标 |
| `word_insert_at` | 在语义位置写入 Markdown |
| `word_edit_cell_at` | 在语义位置编辑表格单元格 |

### 🔄 变量替换

| 工具 | 描述 |
|------|------|
| `word_replace_variables` | 替换文档中的 {{placeholder}} 变量 |

---

返回 [README](./README.zh-CN.md)
