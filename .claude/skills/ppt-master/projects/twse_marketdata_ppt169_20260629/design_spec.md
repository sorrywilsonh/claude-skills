# twse_marketdata - Design Spec

> 證券交易所行情傳輸宣導單頁投影片。現況流量(6/25 09:05 尖峰)＋頻寬不足改善建議。

## I. Project Information

| Item | Value |
| ---- | ----- |
| **Project Name** | twse_marketdata |
| **Canvas Format** | PPT 16:9 (1280×720) |
| **Page Count** | 1 |
| **Design Style** | B) 一般顧問(數據優先) ＋ 證交所官方、穩重專業 |
| **Target Audience** | 資訊商／市場參與者(行情線路接收端) |
| **Use Case** | TWSE 行情傳輸頻寬宣導 |
| **Created Date** | 2026-06-29 |

---

## II. Canvas Specification

| Property | Value |
| -------- | ----- |
| **Format** | PPT 16:9 |
| **Dimensions** | 1280×720 |
| **viewBox** | `0 0 1280 720` |
| **Margins** | 左右 40px、上下 安全區 40px |
| **Content Area** | 標題帶 0–96；內容 120–690 |

---

## III. Visual Theme

### Theme Style

- **Style**: 證交所官方數據宣導,穩重專業
- **Theme**: Light theme
- **Tone**: 權威、清晰、可信賴

### Color Scheme

| Role | HEX | Purpose |
| ---- | --- | ------- |
| **Background** | `#FFFFFF` | 頁面背景 |
| **Secondary bg** | `#F7F9FC` | 卡片／區塊底 |
| **Primary** | `#003366` | 標題帶、區段標題 |
| **Accent (即時)** | `#1565C0` | 即時訊息 IP 長條、重點數字 |
| **Snapshot (快照)** | `#00897B` | 快照訊息 IP 長條 |
| **Tertiary (其他)** | `#90A4AE` | 小流量 IP 長條、輔助 |
| **Body text** | `#1F2937` | 內文 |
| **Secondary text** | `#5B6B7B` | 註解、軸標 |
| **Border/divider** | `#D9E1EC` | 卡片框、格線 |
| **Warning** | `#C62828` | >50 Mbps 超載警示 |
| **Success** | `#2E7D32` | 正常／可行標示 |

### Gradient Scheme

無漸層(平面填色,符合官方穩重調性)。

---

## IV. Typography System

### Font Plan

**Typography direction**: 現代中文黑體,數字加粗強調。

| Role | Chinese | English | Fallback tail |
| ---- | ------- | ------- | ------------- |
| **Title** | `"Microsoft YaHei"` | `Arial` | `sans-serif` |
| **Body** | `"Microsoft YaHei"` | `Arial` | `sans-serif` |
| **Emphasis** | `"Microsoft YaHei"` | `Arial` | `sans-serif` |

**Per-role font stacks**:

- Title: `"Microsoft YaHei", Arial, sans-serif`
- Body: `"Microsoft YaHei", Arial, sans-serif`
- Emphasis: same as Body (bold weight for emphasis)

### Font Size Hierarchy

**Baseline**: Body font size = 18px(資料密集單頁)

| Purpose | Ratio | Size |
| ------- | ----- | ---- |
| Page title | 1.7x | 31px |
| Section title | 1.3x | 24px |
| Hero number (Mbps) | 1.8x | 32px |
| Subtitle | 1.2x | 22px |
| **Body content** | **1x** | **18px** |
| Annotation / axis | 0.8x | 14px |
| Footnote | 0.6x | 11px |

Formula policy: text-only(無公式)。

---

## V. Layout Principles

### Page Structure

- **Header area**: y 0–96,深藍標題帶,頁標題＋宣導標籤
- **Content area**: 左半 現況流量(長條圖＋累計流量升階);右半 三點改善建議＋連結
- **Footer area**: y 695,資料來源註

### Layout Pattern

對稱 5:5 雙欄。左欄=資料(長條圖＋升階流量條);右欄=三點建議卡片(縱向堆疊)＋連結。

### Spacing Specification

- 安全邊距 40px;區塊間距 24–28px;卡片內距 18–22px;卡片圓角 12px。

---

## VI. Icon Usage Specification

### Source

- **Built-in icon library**: `tabler-outline`(線性、專業簡潔),stroke-width 2

### Recommended Icon List

| Purpose | Icon Path | Page |
| ------- | --------- | ---- |
| 即時/快照擇一 | `tabler-outline/arrows-split` | Slide 01 |
| 依需求分流 | `tabler-outline/route` | Slide 01 |
| 提升頻寬 | `tabler-outline/gauge` | Slide 01 |
| 超載警示 | `tabler-outline/alert-triangle` | Slide 01 |
| 申請連結 | `tabler-outline/external-link` | Slide 01 |
| 行情流量 | `tabler-outline/activity-heartbeat` | Slide 01 |

---

## VII. Visualization Reference List

Catalog read: 71 templates

| Page | Template | Path | Summary-quote (verbatim from `charts_index.json`) | Usage |
| ---- | -------- | ---- | ------------------------------------------------- | ----- |
| P01 | bar_chart | `templates/charts/bar_chart.svg` | "Pick for single-series category value comparison, 3-8 categories. Skip for >12 long-label items (use horizontal_bar_chart) or multi-series (use grouped_bar_chart)." | 6/25 09:05 五個 IP 尖峰流量(Mbps)比較 |

**Runners-up considered**:

- `horizontal_bar_chart` | rejected for P01: 僅 5 個短標籤類別(IP1–IP5),且需縱向長條凸顯尖峰高度,bar_chart 更合適。
- `grouped_bar_chart` | rejected for P01: 單一時點(09:05)單一序列流量,非多序列比較。
- `kpi_cards` | rejected for P01: 五個 IP 需相互比較高度差異,非各自獨立指標卡。

---

## VIII. Image Resource List

無圖片(以原生 SVG 長條圖與卡片呈現)。

---

## IX. Content Outline

### Slide 01 - 行情線路頻寬現況與改善建議

- **Layout**: 對稱 5:5 雙欄
- **Title**: 證券交易所行情傳輸宣導 — 行情線路頻寬現況與改善建議
- **Subtitle**: 以 2026/6/25 尖峰時段 09:05 為例
- **Visualization**: bar_chart(左欄 五 IP 尖峰流量)
- **Content（左欄 現況）**:
  - 第一 IP 約 8 Mbps、第二 IP 約 12 Mbps(即時訊息)
  - 第三 IP 約 3 Mbps(快照訊息)
  - 第四 IP 約 0.3 Mbps、第五 IP 約 0.4 Mbps
  - 升階:五 IP 合計約 23.7 Mbps → 含 OTC 約 30 Mbps → 同收兩條線路 >50 Mbps(超過上限)
- **Content（右欄 改善建議）**:
  - 即時/快照擇一接收:IP1·IP2 即時、IP3 快照,依線路頻寬與需求擇一
  - 依業務需求分流接收,或僅接收一份行情資訊
  - 提升硬體線路頻寬:電作部去年 12 月底發文,全市場可申請 70M/100M(附連結)

---

## X. Speaker Notes Requirements

- **Filename**: `01_行情頻寬現況與建議.md`
- **Content**: 宣導口吻,說明尖峰流量現況、超載風險與三項建議。

---

## XI. Technical Constraints Reminder

### SVG Generation Must Follow:

1. viewBox: `0 0 1280 720`
2. 背景用 `<rect>`
3. 換行用 `<tspan>`;禁 `<foreignObject>`
4. 透明用 `fill-opacity`;禁 `rgba()`
5. 禁 `mask`/`<style>`/`class`/`textPath`/`@font-face`/`<animate*>`/`<script>`
6. 文字用原生 Unicode;XML 保留字 `&` → `&amp;`

### PPT Compatibility Rules:

- 禁 `<g opacity>`;逐元素設 opacity
- 內聯樣式;禁外部 CSS
