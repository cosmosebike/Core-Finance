# 工具配置 — 业务系统虾

## 已启用工具

### 飞书内置工具（通过 feishu-openclaw-plugin）

| 工具 | 用途 | 权限级别 |
|------|------|----------|
| `bitable_read` | 读取多维表格记录 | 全部用户可触发 |
| `bitable_search` | 按条件筛选多维表格记录 | 全部用户可触发 |
| `bitable_create_record` | 新增记录 | 仅 Mia 可触发 |
| `bitable_update_record` | 修改记录 | 仅 Mia 可触发，修改前必须确认 |
| `message_send` | 发送飞书消息 | 系统自动 + Mia 触发 |
| `message_send_card` | 发送带格式的卡片消息 | 定时任务使用 |
| `calendar_query` | 查询日历事件 | 仅 Mia |

### 计算工具

| 工具 | 用途 |
|------|------|
| `code_exec` | 执行 Python 计算（汇总、统计、AR分析） |

---

## 未启用工具（原因说明）

| 工具 | 未启用原因 |
|------|-----------|
| `web_search` | 改用 Cowork 做网络搜索，避免重复 |
| `file_write` | 文件生成交给 Cowork，格式更专业 |
| `bitable_delete_record` | 危险操作，禁止删除 |
| `email_send` | 暂无需求 |

---

## 飞书多维表格 App Token 映射

```
宇宙销售统计中心主表：bitable_token = YOUR_APP_TOKEN_HERE
```

> 填写方式：进入飞书多维表格 → 复制链接中的 token 字段
> 格式示例：`https://xxxx.feishu.cn/base/[APP_TOKEN]?table=...`

---

## 工具调用优先级

1. 先读表格，再回答
2. 计算类任务用 `code_exec` 确保精准
3. 写入操作必须先发消息确认，等用户回"确认"才执行
4. 无法完成的任务 → 提示"建议用 Cowork 完成"
