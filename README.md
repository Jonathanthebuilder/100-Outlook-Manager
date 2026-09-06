# Outlook Mailbox Manager

把 Outlook 库存台账与 OAuth 收信合并到一个 Railway 应用中。

## 功能

- 导入和导出 `邮箱----密码----client_id----refresh_token`
- 管理未售/已售状态与订单备注
- 使用 OAuth Token 读取 Outlook 收件箱和垃圾箱
- 默认显示收件箱与垃圾箱的全部邮件，可按文件夹和关键词筛选
- 自动识别并复制常见验证码
- 自动轮转到期的未售库存 Token，并保留单个/每批 25 个手动刷新入口
- 记录 Token 检查时间、刷新时间、健康状态与错误原因
- 已售账号在服务端强制禁止收信和刷新 Token
- 复制发货说明与 CDK 声明

## 本地运行

```bash
npm ci
OUTLOOK_MANAGER_PASSWORD='change-me' npm run dev
```

默认地址为 `http://localhost:4173`。

`npm run dev` 会先构建前端，再启动包含 API 的完整本地服务。`npm run dev:ui` 只启动 Vite 界面，不支持导入和保存台账。

## Railway 配置

需要配置：

```text
OUTLOOK_MANAGER_USERNAME=outlook
OUTLOOK_MANAGER_PASSWORD=使用随机强密码
OUTLOOK_MANAGER_DATA_DIR=/data
AUTO_TOKEN_REFRESH_ENABLED=true
AUTO_TOKEN_REFRESH_BATCH_SIZE=30
AUTO_TOKEN_REFRESH_MAX_AGE_DAYS=30
AUTO_TOKEN_REFRESH_PAUSE_MS=1000
```

自动任务在服务启动一分钟后检查，此后每小时检查一次是否到了每日运行时间。每23小时最多执行一批，默认只处理30天以上未成功检查的 `available`/`prepared` 健康 Token，每批最多30个，逐个刷新并立即写回。已售、未验证和已隔离邮箱不会进入自动任务。运行状态保存在 `/data/token-maintenance.json`，因此服务重启不会重复执行当天批次。

同时给 Railway 服务挂载 Volume 到 `/data`。共享台账保存在：

```text
/data/ledger.json
```

不要把 `ledger.json`、TXT 凭据或环境变量提交到 Git。

## 数据迁移

1. 从旧 Outlook Manager 导出当前有效邮箱 TXT。
2. 从旧 Outlook Token Mailbox 导出 Token TXT。
3. 按邮箱去重，以较新的 Token 为准合并文件。
4. 在本项目中导入合并后的 TXT。
5. 先用少量未售账号验证收信，再切换线上服务。

导入会按邮箱地址合并：新邮箱进入库存，已有邮箱更新凭据但保留已售状态和订单备注。部署前仍应保留加密备份。

## 安全边界

- 收信和刷新接口只允许操作未售的 `available`/`prepared` 记录。
- 标记为 `used` 后，服务端会拒绝对应邮件读取和 Token 刷新请求。
- 读取邮件时 Microsoft 可能返回新的 Refresh Token；系统会同步更新 `refreshToken` 与 `rawCredential`。
- 应用必须放在强密码保护后，不应公开给客户直接访问。
