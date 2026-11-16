# Word-to-MDX 轉換器個人使用手冊

本文記錄個人網站 DeluxeAvidSite 的 Word-to-MDX 轉換器日常使用方法。

- 程式位置：`scripts/word-to-mdx.py`
- 輸入格式：Word `.docx`
- 輸出格式：Markdown `.md` 或 MDX `.mdx`
- 輸出位置：`src/content/blog/`

## 1. 使用前準備

### 1.1 專案位置

```text
E:\Website\DeluxeAvidSite
```

主要目錄：

```text
DeluxeAvidSite/
├── scripts/
│   └── word-to-mdx.py
├── src/
│   └── content/
│       ├── blog/
│       └── essays/
├── docs/
├── requirements.txt
└── .venv/
```

**注意：目前轉換器只會輸出到 `src/content/blog/`。如果要撰寫隨筆，轉換後需要自行調整文章位置及相關配置。**

### 1.2 啟動 Python 環境

在 VS Code 開啟專案，確認終端機目前位於專案根目錄。

PowerShell 執行：

```powershell
.\.venv\Scripts\Activate.ps1
```

成功後，命令列前方通常會出現 `(.venv)`。

目前依賴套件記錄在 `requirements.txt`：

- `python-docx`
- `lxml`
- `pypinyin`

日常使用不需要重新安裝套件。只有環境缺少依賴時，才需要重新安裝。

## 2. 準備 Word 文件

### 2.1 文件位置與命名

- 只支援 `.docx` 格式。
- 建議將待轉換文件放在專案根目錄。
- Word 檔名會直接作為文章標題。
- 檔名中的底線會替換為空格。

例如：

```text
我的遊戲收藏.docx
```

生成標題：

```text
我的遊戲收藏
```

文章目錄名稱則會根據檔名自動生成拼音 slug。

### 2.2 支援的 Word 格式

目前支援：

- 一般正文及 Word 標題樣式
- Word 編號清單及巢狀清單
- 粗體、斜體、底線及刪除線
- 超連結
- 軟換行及 Tab
- 引用段落
- 表格及儲存格內換行
- 圖片及圖片擷取

**標題注意事項：**

轉換器會識別 Word 內建的 Heading 樣式。使用 Word 編號清單的標題段落也有相應處理，但單純在正文中手動輸入 `1. 標題`，不會自動保證轉換為 Markdown 標題。

建議使用 Word 內建標題樣式，並在需要編號時搭配 Word 的編號功能。

## 3. 執行轉換

### 3.1 命令格式

```powershell
python scripts/word-to-mdx.py "Word文件名稱.docx" --date YYYY-MM-DD --description "文章描述"
```

每次執行都必須提供：

| 參數 | 說明 |
|---|---|
| DOCX 檔案名稱 | 待轉換的 Word 文件 |
| `--date` | 發布日期，格式為 `YYYY-MM-DD` |
| `--description` | 文章描述，不可為空白 |

**`--date` 和 `--description` 都是必要參數，不可省略。**

### 3.2 執行範例

```powershell
python scripts/word-to-mdx.py "我的遊戲收藏.docx" --date 2026-11-13 --description "整理個人遊戲收藏及相關資料。"
```

日期及描述應按照實際文章內容修改。

### 3.3 成功訊息

轉換成功後，終端機會顯示類似內容：

```text
已成功讀取：我的遊戲收藏.docx
文章標題：我的遊戲收藏
========== 開始轉換 ==========

========== 轉換完成 ==========
文章：...\index.md
圖片：3 張
```

## 4. 輸出結果

### 4.1 文章目錄

轉換器會根據檔名生成 slug，並建立文章目錄。

例如：

```text
src/content/blog/
└── wo-de-you-xi-shou-cang/
    ├── index.md
    └── images/
        ├── image-01.jpg
        ├── image-02.png
        └── image-03.jpeg
```

如果文章中包含自訂元件，則會生成 `index.mdx`。

### 4.2 Markdown 與 MDX 判斷

轉換器會檢查以下五個元件名稱：

1. `Figure`
2. `ImageGallery`
3. `ImageCompare`
4. `ImageLightbox`
5. `HeaderLink`

只要轉換後的內容中出現符合識別條件的元件標籤，就會選擇 `.mdx`；否則輸出 `.md`。

**注意：這項功能只負責選擇副檔名，不會自動生成元件的 `import`、參數或完整 JSX 語法。**

### 4.3 Frontmatter

轉換器會自動生成：

```yaml
---
title: '文章標題'
description: '文章描述'
pubDate: 2026-11-13
---
```

如果文章包含圖片，還會自動增加 `heroImage`。

## 5. 圖片處理

### 5.1 圖片擷取

圖片按照文件中的出現順序命名：

```text
image-01.jpg
image-02.png
image-03.jpeg
```

統一放入文章目錄的 `images/` 資料夾。

同一張圖片重複使用時，轉換器會透過圖片關聯識別，避免重複擷取。

### 5.2 頭圖設定

目前轉換器會自動將第一張擷取到的圖片設定為頭圖：

```yaml
heroImage: ./images/image-01.jpg
```

如果第一張圖片不是預期封面，需要手動修改。

例如：

```yaml
heroImage: ./images/埼玉縣.png
```

修改前確認圖片確實存在於文章的 `images/` 資料夾。

### 5.3 圖片排版限制

轉換器主要負責擷取圖片及保留圖片在文章中的位置。

Word 的浮動圖片、文字環繞及錨定位置不會完整保留，部分複雜排版需要在 MDX 中手動整理。

## 6. MDX 自訂元件

需要使用自訂元件時，參考：

```text
src/content/blog/using-mdx.mdx
```

### 6.1 Figure

單張圖片及圖說。

| 參數 | 說明 |
|---|---|
| `src` | 圖片變數 |
| `alt` | 替代文字 |
| `caption` | 圖片說明，可選 |
| `layout` | `normal` 或 `wide` |

### 6.2 ImageGallery

多張圖片畫廊。

| 參數 | 說明 |
|---|---|
| `images` | 圖片項目陣列 |
| `columns` | 欄數，可選 2、3 或 4 |
| `layout` | `normal` 或 `wide` |

每個圖片項目包含 `src`、`alt`，以及可選的 `caption`。

### 6.3 ImageCompare

兩張圖片並排比較。

| 參數 | 說明 |
|---|---|
| `left`、`right` | 左右圖片變數 |
| `leftAlt`、`rightAlt` | 左右圖片替代文字 |
| `leftCaption`、`rightCaption` | 左右圖片說明，可選 |
| `layout` | `normal` 或 `wide` |

### 6.4 ImageLightbox

文章圖片燈箱。

這個元件已經由 `BlogPost.astro` 自動載入，會處理正文 `.prose` 內的圖片。

因此：

- 不需要在文章 MDX 中自行匯入。
- 不需要手動插入 `<ImageLightbox />`。
- 一般正文圖片即可使用燈箱功能。

### 6.5 HeaderLink

帶有目前頁面狀態的連結。

接受一般 HTML `<a>` 元素屬性，主要使用 `href`，並透過標籤內的文字作為連結內容。

### 6.6 MDX 使用範例

以下是 Figure 的基本用法：

```mdx
import Figure from '../../components/Figure.astro';
import exampleImage from './images/example.jpg';

<Figure
    src={exampleImage}
    alt="圖片描述"
    caption="圖片說明"
/>
```

其他元件的完整語法及圖片匯入方式，以 `using-mdx.mdx` 和元件實際程式碼為準。

## 7. 錯誤處理與資料保護

### 7.1 常見錯誤

| 情況 | 處理方式 |
|---|---|
| 找不到 DOCX | 確認檔案名稱及位置 |
| 不支援的格式 | 使用 `.docx` |
| 日期無效 | 使用 `YYYY-MM-DD` |
| 描述為空白 | 提供有效的 `--description` |
| 文章目錄已存在 | 檢查既有目錄，避免覆寫 |

### 7.2 已存在的文章目錄

如果相同 slug 的文章目錄已存在，轉換器會中止，不會覆寫舊文章或混入舊圖片。

需要重新轉換時：

1. 先檢查舊文章目錄。
2. 必要時備份文章及圖片。
3. 確認舊目錄不再需要後，再自行處理。
4. 重新執行轉換命令。

不要直接刪除尚未備份的文章資料。

### 7.3 轉換失敗

如果轉換過程發生例外，程式會嘗試清理本次建立的文章目錄。

若出現未預期錯誤，保留完整終端機訊息，以便進一步檢查。

## 8. 轉換後檢查

每篇文章轉換完成後，依序檢查：

1. 終端機是否顯示「轉換完成」。
2. 文章目錄及 slug 是否正確。
3. 副檔名是否符合預期。
4. Frontmatter 的標題、描述及日期是否正確。
5. `heroImage` 是否指向正確圖片。
6. `images/` 是否包含所需圖片。
7. 正文、標題、清單及表格是否正常。
8. MDX 元件的 `import`、參數及圖片路徑是否正確。

### 8.1 本地預覽

```powershell
npm run dev
```

瀏覽器開啟：

```text
http://localhost:4321/
```

### 8.2 建置檢查

```powershell
npm run build
```

確認 Astro 建置成功後，再進行 Git 提交及推送。

## 9. 完整日常流程

```text
Word 撰寫文章
      ↓
將 DOCX 放入專案根目錄
      ↓
啟動 .venv 環境
      ↓
執行轉換命令
      ↓
檢查文章目錄及 Frontmatter
      ↓
檢查圖片及 heroImage
      ↓
必要時整理 MDX 元件
      ↓
npm run dev
      ↓
本地檢查文章
      ↓
npm run build
      ↓
Git commit
      ↓
推送至 GitHub
```

**使用原則：**

Word 負責撰寫，轉換器負責生成文章結構，Astro 負責最終呈現。轉換器不會取代轉換後的人工檢查。