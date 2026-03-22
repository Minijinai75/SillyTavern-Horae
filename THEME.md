# Horae 美化指南

## 快速開始

Horae 的所有視覺樣式由 **CSS 變數** 控制。只需覆蓋這些變數，即可改變整個插件外觀。

### 方式一：修改 CSS 變數（推薦）

在插件設定 → 外觀設定 → 客製化CSS 中輸入：

```css
#horae_drawer,
.horae-message-panel,
.horae-modal,
.horae-context-menu,
.horae-progress-overlay {
    --horae-primary: #ec4899;      /* 主色改為粉色 */
    --horae-primary-light: #f472b6;
    --horae-bg: #1a1020;           /* 背景改為深紫 */
    --horae-bg-secondary: #2d1f3c;
}
```

### 方式二：導入美化資料

1. 獲取他人分享的 `.json` 美化資料
2. 插件設定 → 外觀設定 → 點擊導入按鈕（📥）
3. 在主題下拉列表中選擇導入的美化

### 方式三：導出並分享

1. 調好你喜歡的樣式後，點擊導出按鈕（📤）
2. 會下載一個 `horae-theme.json` 資料
3. 分享給其他用戶即可

---

## CSS 變數一覽

### 配色

| 變數 | 預設值（暗色） | 說明 |
|------|---------------|------|
| `--horae-primary` | `#7c3aed` | 主色（按鈕、發光、漸變） |
| `--horae-primary-light` | `#a78bfa` | 主色亮版（文字發光） |
| `--horae-primary-dark` | `#5b21b6` | 主色暗版（漸變起點） |
| `--horae-accent` | `#f59e0b` | 強調色（金色標記、NPC名） |
| `--horae-success` | `#10b981` | 成功色（好感度正值） |
| `--horae-warning` | `#f59e0b` | 警告色 |
| `--horae-danger` | `#ef4444` | 危險色（刪除、負好感度） |
| `--horae-info` | `#3b82f6` | 資訊色（NPC外貌標籤） |

### 背景與邊框

| 變數 | 預設值（暗色） | 說明 |
|------|---------------|------|
| `--horae-bg` | `#1e1e28` | 主背景（區塊、卡片） |
| `--horae-bg-secondary` | `#2d2d3c` | 次級背景（容器、表頭） |
| `--horae-bg-hover` | `#3c3c50` | 懸停背景 |
| `--horae-border` | `rgba(255,255,255,0.1)` | 邊框色 |

### 文字

| 變數 | 預設值（暗色） | 說明 |
|------|---------------|------|
| `--horae-text` | `#e5e5e5` | 主文字色 |
| `--horae-text-muted` | `#a0a0a0` | 次級文字色（標籤、提示） |

### 雷達圖

| 變數 | 預設值 | 說明 |
|------|--------|------|
| `--horae-radar-color` | 跟隨 `--horae-primary` | 雷達圖數據區顏色（填充/描邊/頂點） |
| `--horae-radar-label` | 跟隨 `--horae-text` | 雷達圖示籤文字顏色 |

### 其他

| 變數 | 預設值 | 說明 |
|------|--------|------|
| `--horae-shadow` | `0 4px 20px rgba(0,0,0,0.3)` | 陰影 |
| `--horae-radius` | `8px` | 大圓角 |
| `--horae-radius-sm` | `4px` | 小圓角 |

---

## 主要容器類名

想針對特定區域微調樣式時，使用以下選擇器：

### 頂級容器

| 選擇器 | 說明 |
|--------|------|
| `#horae_drawer` | 主抽屜面板（設定、狀態、時間線等） |
| `.horae-message-panel` | 訊息底部的元數據面板 |
| `.horae-modal` | 所有模態彈窗 |
| `.horae-context-menu` | 右鍵選單 |
| `.horae-progress-overlay` | 進度覆蓋層 |

### 抽屜內部

| 選擇器 | 說明 |
|--------|------|
| `.horae-tabs` | 標籤頁導航欄 |
| `.horae-tab` | 單個標籤頁按鈕 |
| `.horae-tab-contents` | 標籤頁內容容器 |
| `.horae-state-section` | 狀態區塊（儀表板內的各個卡片） |
| `.horae-settings-section` | 設定區塊 |

### 數據展示

| 選擇器 | 說明 |
|--------|------|
| `.horae-timeline-item` | 時間線事件卡片 |
| `.horae-timeline-list` | 時間線列表容器 |
| `.horae-affection-item` | 好感度條目 |
| `.horae-npc-item` | NPC 卡片 |
| `.horae-full-item` | 物品條目 |
| `.horae-item-tag` | 物品標籤（小圓角膠囊） |
| `.horae-agenda-item` | 待辦事項條目 |
| `.horae-relationship-item` | 關係網通行證目 |
| `.horae-relationship-list` | 關係網路列表容器 |
| `.horae-location-card` | 場景記憶卡片 |
| `.horae-mood-tag` | 情緒標籤（圓角膠囊） |
| `.horae-panel-rel-row` | 底部面板關係行 |
| `.horae-empty-hint` | 空數據提示文字 |

### 摘要與壓縮

| 選擇器 | 說明 |
|--------|------|
| `.horae-timeline-item.summary` | 摘要事件卡片（active 狀態） |
| `.horae-timeline-item.horae-summary-collapsed` | 已展開為原始事件時的摺疊指示條 |
| `.horae-summary-actions` | 摘要卡片上的切換/刪除按鈕容器 |
| `.horae-summary-toggle-btn` | 摘要/時間線切換按鈕 |
| `.horae-summary-delete-btn` | 刪除摘要按鈕 |
| `.horae-compressed-restored` | 被摘要覆蓋但目前已恢復顯示的事件（虛線框） |

### 客製化表格

| 選擇器 | 說明 |
|--------|------|
| `.horae-excel-table-container` | 表格外層容器 |
| `.horae-excel-table` | 表格主體 `<table>` |
| `.horae-excel-table th` | 表頭單元格 |
| `.horae-excel-table td` | 數據單元格 |
| `.horae-table-prompt-row` | 表格底部提示詞區域 |

### 按鈕

| 選擇器 | 說明 |
|--------|------|
| `.horae-btn` | 通用按鈕 |
| `.horae-btn.primary` | 主色按鈕（紫色漸變） |
| `.horae-btn.danger` | 危險按鈕（紅色） |
| `.horae-icon-btn` | 小圖示按鈕（28×28） |
| `.horae-data-btn` | 數據管理大按鈕（帶圖示+文字） |
| `.horae-data-btn.primary` | 主功能按鈕（跨兩列） |

---

## 美化資料格式

導出的 `.json` 資料結構如下：

```json
{
    "name": "我的美化",
    "author": "你的名字",
    "version": "1.0",
    "variables": {
        "--horae-primary": "#ec4899",
        "--horae-primary-light": "#f472b6",
        "--horae-primary-dark": "#be185d",
        "--horae-accent": "#f59e0b",
        "--horae-bg": "#1a1020",
        "--horae-bg-secondary": "#2d1f3c",
        "--horae-bg-hover": "#3c2f50",
        "--horae-border": "rgba(255, 255, 255, 0.08)",
        "--horae-text": "#e5e5e5",
        "--horae-text-muted": "#a0a0a0"
    },
    "css": "/* 可選：額外CSS覆蓋 */\n.horae-timeline-item { border-radius: 12px; }"
}
```

**資料欄說明：**
- `name`：美化名稱（顯示在主題選擇器中）
- `author`：作者名（可選）
- `version`：版本號（可選）
- `variables`：CSS 變數鍵值對，會覆蓋預設變數
- `css`：額外的 CSS 代碼（可選），用於無法透過變數實現的樣式調整

---

## 示例美化

### 櫻花粉

```json
{
    "name": "櫻花粉",
    "variables": {
        "--horae-primary": "#ec4899",
        "--horae-primary-light": "#f472b6",
        "--horae-primary-dark": "#be185d",
        "--horae-accent": "#fb923c",
        "--horae-bg": "#1f1018",
        "--horae-bg-secondary": "#2d1825",
        "--horae-bg-hover": "#3d2535",
        "--horae-text": "#fce7f3",
        "--horae-text-muted": "#d4a0b9"
    }
}
```

### 森林綠

```json
{
    "name": "森林綠",
    "variables": {
        "--horae-primary": "#059669",
        "--horae-primary-light": "#34d399",
        "--horae-primary-dark": "#047857",
        "--horae-accent": "#fbbf24",
        "--horae-bg": "#0f1a14",
        "--horae-bg-secondary": "#1a2e22",
        "--horae-bg-hover": "#2a3e32",
        "--horae-text": "#d1fae5",
        "--horae-text-muted": "#6ee7b7"
    }
}
```

### 海洋藍

```json
{
    "name": "海洋藍",
    "variables": {
        "--horae-primary": "#3b82f6",
        "--horae-primary-light": "#60a5fa",
        "--horae-primary-dark": "#1d4ed8",
        "--horae-accent": "#f59e0b",
        "--horae-bg": "#0c1929",
        "--horae-bg-secondary": "#162a45",
        "--horae-bg-hover": "#1e3a5f",
        "--horae-text": "#dbeafe",
        "--horae-text-muted": "#93c5fd"
    }
}
```

---

## 常見問題 & 美化技巧

### 底部面板被其他元素遮擋（無法互動）

部分酒館美化或預設的 z-index 較高，導致 Horae 底部面板被蓋住。在客製化 CSS 中添加：

```css
.horae-message-panel {
    margin-bottom: 10px;
    z-index: 9999;
    position: relative;
}
```

### 客製化頂部抽屜圖示

將頂部導航欄的 Horae 圖示替換為客製化圖片：

```css
#horae_drawer .drawer-icon::before {
    background-image: url('你的圖片URL') !important;
}
```

---

## 注意事項

1. **變數作用域**：CSS 變數定義在 `#horae_drawer`、`.horae-modal` 等頂級容器上，不要在 `body` 或 `:root` 上定義，否則不會生效。

2. **`!important` 防護**：部分按鈕樣式帶有 `!important` 以抵抗酒館全域主題干擾。如需覆蓋這些樣式，你的客製化 CSS 也需要使用 `!important`。

3. **深淺模式**：客製化美化選擇後，會覆蓋預設的暗色/淺色變數。如果你的美化是淺色系，記得調整 `--horae-text` 為深色。

4. **不影響酒館**：Horae 的所有樣式都限定在插件容器內，不會影響酒館主界面。
