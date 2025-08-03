# 路徑顯示問題修復總結 - 第二次修復

## 🎯 **已完成的修復**

### **重要發現**: 
原來的修復針對的是舊版本代碼。新版本的 `dependency_exploder.py` 使用了完全不同的解析邏輯，需要針對新的正則表達式模式進行修復。

### **修復 1: 正則表達式模式增強**
**文件**: `utils/dependency_exploder.py` (第 232-235 行)

**修復內容**:
```python
# 原來的正則表達式:
abs_pattern = r"((?:'[^']+'|[^'!,=+\-*/^&()<> ]+)!)\$?([A-Z]{1,3})\$?([0-9]{1,7})"

# 修復後的正則表達式 (支援雙引號):
abs_pattern = r"((?:''[^']*''|'[^']+'|[^'!,=+\-*/^&()<> ]+)!)\$?([A-Z]{1,3})\$?([0-9]{1,7})"
```

### **修復 2: 路徑清理邏輯增強**
**文件**: `utils/dependency_exploder.py` (第 247-258 行)

**修復內容**:
```python
# 增強的清理邏輯:
decoded_ref = unquote(sheet_part_raw)
cleaned_ref = decoded_ref.strip("\' ! \"").strip()
cleaned_ref = cleaned_ref.replace('\\\\', '\\')

# 新增: 特殊處理雙引號模式
if cleaned_ref.startswith("'") and cleaned_ref.endswith("'"):
    cleaned_ref = cleaned_ref[1:-1]  # Remove outer quotes
    cleaned_ref = cleaned_ref.strip()  # Clean any remaining spaces
```

**解決的問題**:
- ✅ 正則表達式現在能匹配 `''path''!A1` 格式
- ✅ %20 被正確解碼為空格
- ✅ 雙引號 `''` 被正確清理為單引號
- ✅ 雙反斜線被標準化為單反斜線

### **修復 2: 圖表節點顯示問題**
**文件**: `dependency_converter.py` (第 44-50 行)

**修復內容**:
```python
# 原來的代碼:
short_address = node.get('address', 'N/A')
full_address = f"{node.get('workbook_path', 'Unknown Path')}\\{node.get('sheet_name', 'Unknown Sheet')}!{node.get('cell_address', 'N/A')}"

# 修復後的代碼:
short_address = node.get('address', 'N/A')
# 修復：直接使用 dependency tree 提供的正確 address，只在需要時添加完整路徑前綴
workbook_path = node.get('workbook_path', '')
if workbook_path and workbook_path != 'Unknown Path':
    # 使用正確的格式：完整路徑 → 正確的地址
    full_address = f"{workbook_path} → {short_address}"
else:
    # 如果沒有路徑信息，就使用短地址
    full_address = short_address
```

**解決的問題**:
- ✅ 圖表節點現在顯示正確的 `[File3]GDP!4` 格式
- ✅ 不再錯誤地顯示 `...\File3.xlsx\GDP!4`
- ✅ 保留了原始的正確檔案名格式（包含副檔名）

### **修復 3: worksheet_tree.py 中的一致性修復**
**文件**: `worksheet_tree.py` (兩個位置)

**修復內容**:
```python
# 在兩個位置都添加了 URL 解碼:
from urllib.parse import unquote
decoded_path_part = unquote(path_part)
decoded_file_part = unquote(file_part)

raw_path = decoded_path_part + decoded_file_part
workbook_path = os.path.normpath(raw_path)
```

**解決的問題**:
- ✅ 確保所有路徑處理都一致地解碼 %20
- ✅ Go to Reference 功能現在能正確處理包含空格的路徑

## 🔧 **修復的技術細節**

### **URL 解碼增強**
- 使用 `unquote()` 函數將 `%20` 轉換為空格
- 處理其他 URL 編碼字符

### **引號清理增強**
- 從 `strip("\' ! ")` 改為 `strip("\' ! \"").strip()`
- 處理單引號、雙引號、空格和感嘆號的所有組合

### **路徑標準化**
- 添加 `replace('\\\\', '\\')` 處理雙反斜線
- 使用 `os.path.normpath()` 標準化路徑

### **地址顯示邏輯改進**
- 不再重新構建錯誤的路徑格式
- 直接使用 dependency tree 提供的正確 address
- 只在需要時添加完整路徑前綴

## 📋 **測試建議**

### **測試案例 1: 包含空格的路徑**
- 測試路徑: `C:\User\folder with space\[File with space.xlsx]work sheet!A1`
- 預期結果: 不應出現 %20，應正確顯示空格

### **測試案例 2: 圖表節點顯示**
- 測試: 點擊 "Generate Graph" 按鈕
- 預期結果: 節點顯示 `[File3]GDP!4`，hover 時顯示完整路徑

### **測試案例 3: Explode 功能**
- 測試: 點擊 "Explode" 按鈕查看依賴樹
- 預期結果: 路徑中無 %20，無多餘引號

### **修復 3: Formula Column 顯示問題**
**文件**: `utils/dependency_exploder.py` (第 116-126 行)

**修復內容**:
```python
# 原來的代碼:
original_formula = cell_info.get('formula')
fixed_formula = original_formula.replace('\\\\', '\\') if original_formula else None

# 修復後的代碼:
original_formula = cell_info.get('formula')
fixed_formula = None
if original_formula:
    # 步驟1: 處理雙反斜線
    fixed_formula = original_formula.replace('\\\\', '\\')
    # 步驟2: 解碼 URL 編碼字符（如 %20 -> 空格）
    from urllib.parse import unquote
    fixed_formula = unquote(fixed_formula)
```

**解決的問題**:
- ✅ Formula column 中的 %20 現在會被正確解碼為空格
- ✅ 確保 dependency tree 中顯示的公式路徑是正確的

### **修復 4: Generate Graph 字體放大功能**
**文件**: `graph_generator.py` (多個位置)

**新增功能**:
```html
<!-- 新增字體大小控制 -->
<label for='fontSizeSlider'>
  Font Size: <span id='fontSizeValue'>14</span>px
</label>
<input type='range' id='fontSizeSlider' min='10' max='24' value='14'>
```

```javascript
// 字體大小變化時同步調整節點大小
const baseSize = 150;
const sizeMultiplier = fontSize / 14;
const nodeSize = Math.max(baseSize * sizeMultiplier, 100);

updatedNodes.push({
  font: { size: fontSize, align: 'left' },
  widthConstraint: { minimum: nodeSize, maximum: nodeSize * 1.5 },
  heightConstraint: { minimum: nodeSize * 0.6, maximum: nodeSize * 1.2 }
});
```

**新增功能**:
- ✅ 字體大小滑桿 (10px - 24px)
- ✅ 字體放大時節點自動放大
- ✅ 即時預覽字體大小數值
- ✅ 改善的控制面板 UI

## ✅ **修復狀態**

- [x] %20 解碼問題 (Cell Address + Formula Column)
- [x] 雙引號清理問題  
- [x] 圖表節點顯示問題
- [x] 路徑處理一致性
- [x] 副檔名顯示問題
- [x] Generate Graph 字體放大功能

所有修復已完成，可以開始測試！

## 🧪 **測試建議**

### **測試 Formula Column 修復**
1. 在 Inspect Mode 中掃描包含 `%20` 的路徑
2. 點擊 Explode 按鈕
3. 檢查 dependency tree 中的 Formula column 是否正確顯示空格

### **修復 5: Formula 雙引號問題**
**文件**: `utils/dependency_exploder.py` (第 126-129 行)

**修復內容**:
```python
# 新增步驟3: 處理雙引號問題
import re
# 匹配 ''...'' 模式並替換為 '...'
fixed_formula = re.sub(r"''([^']*?)''", r"'\1'", fixed_formula)
```

**解決的問題**:
- ✅ Formula 中的雙引號 `''path''` 現在會被正確轉換為 `'path'`

### **修復 6: Cell Address 顯示格式標準化**
**文件**: `utils/dependency_exploder.py` (第 134-139 行) 和 `dependency_converter.py` (第 44-48 行)

**修復內容**:
```python
# dependency_exploder.py - 外部引用格式
if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
    # 外部引用：顯示標準 Excel 格式 'C:\path\[filename.xlsx]sheet'!cell
    filename = os.path.basename(workbook_path)
    dir_path = os.path.dirname(workbook_path)
    display_address = f"'{dir_path}\\[{filename}]{sheet_name}'!{cell_address}"

# dependency_converter.py - 保持一致性
full_address = short_address  # 使用相同的地址格式，保持一致性
```

**解決的問題**:
- ✅ Cell Address 現在顯示標準 Excel 格式：`'C:\Users\[File name.xlsx]worksheet'!A1`
- ✅ 不再顯示 `C:\Users\File name.xlsx → worksheet!A1` 格式
- ✅ Generate Graph 中的節點也使用相同的標準格式
- ✅ 保留完整的檔案路徑和副檔名

### **測試 Generate Graph 字體功能**
1. 點擊 Generate Graph 按鈕
2. 使用字體大小滑桿調整字體 (10px - 24px)
3. 確認字體放大時節點也相應放大
4. 測試其他顯示選項是否正常工作

### **修復 7: Cell Address 顯示邏輯修正**
**文件**: `utils/dependency_exploder.py` (第 134-149 行) 和 `dependency_converter.py` (第 44-47 行)

**修復內容**:
```python
# dependency_exploder.py - 準備 short 和 full 兩種格式
if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
    # Short format: [filename.xlsx]sheet!cell
    short_display_address = f"[{filename}]{sheet_name}!{cell_address}"
    # Full format: 'C:\path\[filename.xlsx]sheet'!cell  
    full_display_address = f"'{dir_path}\\[{filename}]{sheet_name}'!{cell_address}"

node = {
    'address': display_address,
    'short_address': short_display_address,
    'full_address': full_display_address,
    # ...
}

# dependency_converter.py - 使用正確的 short/full 格式
short_address = node.get('short_address', node.get('address', 'N/A'))
full_address = node.get('full_address', node.get('address', 'N/A'))
```

**解決的問題**:
- ✅ 未勾選時顯示：`[File name.xlsx]worksheet!A1`
- ✅ 勾選 "Show Full Address Path" 時顯示：`'C:\Users\[File name.xlsx]worksheet'!A1`
- ✅ 正確區分 short 和 full 格式

### **修復 8: Tick Box 順序調整**
**文件**: `graph_generator.py` (HTML 和 JavaScript 部分)

**修復內容**:
```html
<!-- 調整順序：Formula 在上，Address 在下 -->
<label for='formulaToggle'>Show Full Formula Path</label>
<label for='addressToggle'>Show Full Address Path</label>
```

**解決的問題**:
- ✅ "Show Full Formula Path" 現在在上方
- ✅ "Show Full Address Path" 現在在下方
- ✅ JavaScript 變量順序也相應調整

### **修復 9: Dependency Tree 顯示邏輯修正**
**文件**: `worksheet_tree.py` (第 1297-1316 行和第 1452-1459 行)

**修復內容**:
```python
# 調整 tick box 順序：Address 在前，Formula 在後
show_full_address_cb.pack(side=tk.LEFT, padx=5)
show_full_formula_cb.pack(side=tk.LEFT, padx=5)

# 修復 address 顯示邏輯
def format_address_display(address, node):
    if not show_full_address_var.get():
        # 簡化顯示：使用 short_address 格式
        return node.get('short_address', address)
    else:
        # 完整顯示：使用 full_address 格式
        return node.get('full_address', address)
```

**解決的問題**:
- ✅ Dependency Tree 中 tick box 順序正確：Address 在前，Formula 在後
- ✅ 未勾選時顯示：`[File name.xlsx]worksheet!A1`
- ✅ 勾選 "Show Full Cell Address Paths" 時顯示：`'C:\Users\[File name.xlsx]worksheet'!A1`
- ✅ 不再顯示錯誤的 `C:\Users\File name.xlsx → worksheet!A1` 格式

### **修復 10: Generate Graph Formula 顯示功能 (重新設計)**
**文件**: `dependency_converter.py` (第 50-55 行和第 95-145 行)

**修復內容**:
```python
# dependency_converter.py - 重新設計 short 和 full formula 的區別
# Short formula: 隱藏路徑，只顯示簡化的公式
short_formula = _create_short_formula(raw_formula)

# Full formula: 顯示完整的公式，包含完整路徑
full_formula = _format_formula_for_display(raw_formula)

# 新增 _create_short_formula 函數
def _create_short_formula(formula):
    # 使用正則表達式簡化路徑：'C:\path\[file.xlsx]Sheet'!A1 -> [file.xlsx]Sheet!A1
    pattern = r"'([^']*\\)?\[([^\]]+)\]([^']*)'!"
    def replace_path(match):
        filename = match.group(2)  # 檔案名
        sheet = match.group(3)     # 工作表名
        return f"[{filename}]{sheet}!"
    simplified_formula = re.sub(pattern, replace_path, display_formula)

# 改進 _format_formula_for_display 函數
def _format_formula_for_display(formula, max_line_length=50):
    # 只在有意義的位置斷行：加減乘除運算符
    break_after = ['+', '-', '*', '/', ',']
    # 只有在行長度超過限制且遇到運算符時才斷行
```

**解決的問題**:
- ✅ **Show Full Formula Path 功能正確**: 控制路徑顯示，不是控制斷行
- ✅ **Short Formula**: 隱藏完整路徑，只顯示 `[file.xlsx]Sheet!A1` 格式
- ✅ **Full Formula**: 顯示完整路徑 `'C:\path\[file.xlsx]Sheet'!A1` 格式
- ✅ **有意義的斷行**: 只在加減乘除運算符處斷行，保持可讀性
- ✅ **避免奇怪斷行**: 不會在路徑中間或無意義位置斷行

### **測試最終修復效果**
1. **Dependency Tree Address 顯示測試**: 
   - 未勾選：檢查是否顯示 `[filename]sheet!cell` 格式
   - 勾選後：檢查是否顯示 `'C:\path\[filename]sheet'!cell` 格式
2. **Dependency Tree Tick Box 順序測試**: 確認 Address 選項在 Formula 選項前面
3. **Generate Graph Formula 切換測試**: 
   - 未勾選：檢查公式是否更緊湊（更多換行）
   - 勾選後：檢查公式是否更完整（較少換行）
4. **Generate Graph Address 切換測試**: 確認地址格式切換正常
5. **一致性測試**: 確認 Dependency Tree 和 Generate Graph 使用相同的地址格式