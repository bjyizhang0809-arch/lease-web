# 租赁合同计算器 Web 版

将商业租赁合同 Excel 数据上传，自动计算指定时间段内的**应收总额**、**收入总额**，并提供月度明细报表下载。

> 基于 924 行 Python 计算引擎，替代复杂的 Excel LET 嵌套公式。支持免租期、多年度租金变化、银行/发票对账。

---

## 功能概览

| 功能 | 说明 |
|------|------|
| 应收总额计算 | 按天精确计算，支持免租期、跨年度租金变化 |
| 收入总额计算 | 日均摊法，平滑到合同全周期（会计报表用） |
| 银行/发票对账 | 按客户名称和日期自动匹配 |
| 数据校验 | 自动检测租期与租金年数冲突，输出告警 |
| 批量处理 | 一次上传处理全部合同，生成 3 个 Excel 报表 |
| 在线下载 | lease / single / income 三份报表独立下载 |

---

## 快速开始（Vercel 一键部署）

1. Fork 本仓库
2. 访问 [vercel.com](https://vercel.com) → Add New Project → 选择 fork 的仓库
3. 保持默认设置 → Deploy
4. ~2 分钟后获得公网链接，例如 `https://lease-web-xxx.vercel.app`

---

## 本地开发

```bash
# 安装前端依赖
npm install

# 安装 Python 依赖
pip install pandas openpyxl python-dateutil

# 启动开发服务器
npm run dev
```

访问 http://localhost:3000

---

## 使用方式

1. **下载模板**：点击页面上的「下载数据填写模板」按钮，获取 `template.xlsx`
2. **填写数据**：按模板格式填写合同数据（删除示例数据行，填入真实数据）
3. **上传文件**：拖放或点击上传填好的 xlsx 文件
4. **选择区间**：选择计算的起始月和结束月
5. **下载报表**：计算完成后下载 3 个 Excel 报表

---

## 数据模板说明

模板包含 3 个 Sheet：

### Sheet 1 — 合同原始数据（必填）

| 列名 | 说明 | 示例 |
|------|------|------|
| 客户名称 | 客户公司全名 | 北京lbcy餐饮管理有限公司 |
| 商户编号 | 商铺编号 | B1-01c |
| 交付日 | 合同起始日，格式 YYYY-MM-DD | 2025-05-12 |
| 租期届满日 | 合同到期日（含），格式 YYYY-MM-DD | 2027-05-11 |
| 免租期 | 免租天数，无则填 0 | 30 |
| 保底租金第1年（必须） | 第1年年租金 | 26496.00 |
| 保底租金第2年 | 第2年年租金（可选） | 27820.80 |
| 保底租金第N年 | 支持最多7年，多年合同可扩展列 | — |

### Sheet 2 — 银行对账单（可选）

| 列名 | 说明 |
|------|------|
| 交易时间 | 格式 YYYY-MM-DD |
| 贷方发生额（收入） | 收到的金额，正数 |
| 对方户名 | 付款方名称（需与合同客户名称一致） |

### Sheet 3 — 发票信息汇总表（可选）

| 列名 | 说明 |
|------|------|
| 购买方名称 | 需与合同客户名称一致 |
| 开票日期 | 格式 YYYY-MM-DD |
| 价税合计 | 发票金额（含税） |

---

## 输出文件说明

| 文件名 | 内容 |
|--------|------|
| `lease.xlsx` | 各合同汇总：应收总额、收入总额、银行匹配、发票匹配、数据告警 |
| `single.xlsx` | 月度应收明细：每个合同每个月的应收金额 |
| `income.xlsx` | 月度收入明细：每个合同每个月的收入确认金额 |

---

## 项目结构

```
lease-web/
├── api/
│   └── calculate.py          # Vercel Python Serverless 函数
├── lib/
│   └── lease_calculator.py   # 计算引擎（924 行）
├── src/app/
│   ├── page.tsx              # 主页面（React）
│   ├── layout.tsx            # HTML 外壳
│   └── globals.css           # Tailwind 样式
├── public/
│   └── template.xlsx         # 数据填写模板
├── scripts/
│   └── generate_template.py  # 重新生成模板的脚本
├── requirements.txt           # Python 依赖
├── vercel.json                # Vercel 配置（Python runtime + 60s 超时）
└── package.json               # Node.js 依赖
```

---

## 技术栈

- **前端**: Next.js 15 + React 19 + Tailwind CSS
- **后端**: Vercel Python Serverless Function（python3.12）
- **计算引擎**: Python + pandas + openpyxl
- **部署**: Vercel（免费，自动 HTTPS，全球 CDN）
