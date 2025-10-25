# 邮件 HTML 渲染修复（已集成）

本次在原项目中加入 `render_fix/` 模块，并（若检测到 Express 入口）自动挂载了以下路由：

- `GET /api/message/:id` → 返回 `{ html, from }`（优先 HTML，其次文本，最后原始片段）
- `GET /api/message/:id/html` → 直接输出渲染后的 HTML 页面

关键能力：
- 自动优先 `text/html`，正确处理 `quoted-printable/base64` 与多字符集
- 当 Graph `body` 不可用或为 `text` 时，自动回退到 `/$value` 原始 MIME，用 `mailparser` 解析
- 尝试把小体积 `cid:` 内联资源内联为 `data:` URL，减少图片丢失
- 输出 HTML 时不做二次转义，避免“源码样式”

## 使用方法
1. `npm i` 安装依赖（已把 `mailparser、node-fetch、express` 写入 package.json）
2. 保持你原有启动方式不变（`npm start`/`pm2`）。如果自动挂载成功，接口会直接可用；
   若未自动挂载（某些项目结构较特殊），请手工在你的 Express 入口中加入：

   ```js
   // ESM:
   import { mountMailHtmlRoutes } from './render_fix/routes.js';
   // 或 CJS:
   // const { mountMailHtmlRoutes } = require('./render_fix/routes.js');

   // 在 `const app = express()` 之后：
   mountMailHtmlRoutes(app, '/api');
   ```

3. 环境变量沿用你现有配置；若用 Graph 刷新令牌方式，需要：
   - `MS_CLIENT_ID` / `MS_CLIENT_SECRET` / `MS_REFRESH_TOKEN`
   - 可选：`MS_TENANT_ID`（默认 common），`MS_USER_ID`（不填则使用 /me）

4. 在你的前端/模板里，渲染从 `/api/message/:id` 返回的 `html` 时，**不要 HTML 转义**（Blade 用 `{!! $html !!}`）。

## 兼容旧接口
- 新增：`GET /api/mail-new?message_id=...&response_type=html|json`
  - `response_type=html` 时直接输出 HTML 页面（最兼容你现有的“HTML 展开”）。
