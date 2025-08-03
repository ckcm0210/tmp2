# INDIRECT解析工具 - 完整技術文檔

## 概述

呢個工具提供兩種模式嚟解析Excel嘅INDIRECT函數：
1. **Excel模式** - 需要打開Excel，速度快，準確度高
2. **Pure模式** - 唔需要Excel，純粹用openpyxl，適合服務器環境

## Pure模式技術原理

### 核心挑戰：點樣喺唔打開Excel嘅情況下計算函數

Excel函數（如ROW(), COLUMN(), VLOOKUP等）嘅計算結果會根據佢哋所在嘅儲存格位置而改變。Pure模式需要模擬呢個行為。

### 解決方案：Context-Aware函數計算

#### 1. 儲存格上下文傳遞
```python
def resolve_position_aware_function(self, func_expr, func_type, context_cell=None):
    """
    統一處理位置相關函數（ROW, COLUMN等）
    context_cell: 公式實際所在嘅儲存格地址
    func_type: 函數類型 ('ROW', 'COLUMN', 'ADDRESS'等)
    
    例如：B11儲存格入面有 ="B"&(ROW()+18)
    咁ROW()就應該返回11，而唔係當前選中儲存格嘅行號
    """
    target_cell = context_cell if context_cell else self.cell_var.get()
    
    if func_type == 'ROW':
        row_num = int(re.search(r'\d+', target_cell).group())
        # ROW()+18 = 11+18 = 29
        return row_num
    elif func_type == 'COLUMN':
        col_letters = re.search(r'[A-Z]+', target_cell).group()
        # B列 = 2
        return column_letters_to_number(col_letters)
```

#### 2. 智能組件分析
```python
# INDIRECT內容：="'"&B8&"\"&"["&B9&"]"&B10&"'"&"!"&B11
# 分析每個組件：
components = [
    ('string', "'"),           # 字串常數
    ('cell', 'B8'),           # 儲存格引用
    ('string', '\'),          # 字串常數  
    ('cell', 'B11')           # 儲存格引用（可能包含公式）
]
```

#### 3. 嵌套公式處理
當遇到B11儲存格包含公式 `="B"&(ROW()+18)` 時：
```python
# 步驟1：識別係公式
if raw_value.startswith('='):
    formula = raw_value[1:]  # 移除=號
    
# 步驟2：按&分割
parts = ["B", "(ROW()+18)"]

# 步驟3：逐個處理
for part in parts:
    if 'ROW()' in part:
        # 傳入B11作為context
        result = resolve_position_aware_function(part, 'ROW', 'B11')
        # ROW() = 11, ROW()+18 = 29
```

### 關鍵技術點

#### 1. 括弧配對算法
```python
def extract_indirect_content(self, formula):
    """處理嵌套括弧同引號"""
    bracket_count = 1
    in_quotes = False
    quote_char = None
    
    for char in formula:
        if char in ['"', "'"] and not in_quotes:
            in_quotes = True
            quote_char = char
        elif char == quote_char and in_quotes:
            in_quotes = False
        elif not in_quotes:  # 只在引號外計算括弧
            if char == '(':
                bracket_count += 1
            elif char == ')':
                bracket_count -= 1
```

#### 2. 智能字串分割
```python
def smart_split_by_ampersand(self, content):
    """按&分割，但唔會分割引號入面嘅&"""
    # 例子："'"&B8&"\"&"["  ->  ["'", "B8", "\", "["]
```

#### 3. 函數類型識別
```python
def identify_component_type(self, component):
    """自動識別組件類型"""
    if component.startswith('"'):
        return ('string', component[1:-1])
    elif re.match(r'^\$?[A-Z]+\$?\d+$', component):
        return ('cell', component)
    elif re.match(r'^[A-Z]+\s*\(', component):
        return ('function', component)
```

## 工具使用指南

### Excel模式

**適用場景**：
- 開發環境
- 有Excel安裝嘅機器
- 需要100%準確度
- 複雜嘅跨文件引用

**使用方法**：
1. 選擇"Excel Mode"
2. 載入Excel文件
3. 選擇工作表同儲存格
4. 點擊"Resolve INDIRECT"

**工作原理**：
```python
# 1. 連接Excel COM
xl = win32.GetActiveObject("Excel.Application")

# 2. 提取INDIRECT內容
indirect_content = extract_indirect_content(formula)

# 3. 直接喺Excel計算
excel_cell.Formula = f"={indirect_content}"
result = excel_cell.Value
```

### Pure模式

**適用場景**：
- 服務器環境（冇Excel）
- 批量處理
- 自動化腳本
- 基本嘅INDIRECT解析

**使用方法**：
1. 選擇"Pure Mode"
2. 載入Excel文件（用openpyxl）
3. 選擇工作表同儲存格
4. 點擊"Resolve INDIRECT"

**支援嘅函數**：
- 位置相關函數：ROW(), ROW()+數字, COLUMN()
- 查找函數：VLOOKUP()
- 字串操作：字串連接（&）
- 基本引用：儲存格引用

## 整合到v70工具嘅建議

### 1. 模組化設計
```python
class IndirectResolver:
    def __init__(self, mode='auto'):
        self.mode = mode  # 'excel', 'pure', 'auto'
    
    def resolve(self, file_path, sheet_name, cell_address):
        if self.mode == 'auto':
            # 嘗試Excel模式，失敗就用Pure模式
            try:
                return self.resolve_excel_mode(...)
            except:
                return self.resolve_pure_mode(...)
    
    def resolve_position_aware_function(self, func_expr, func_type, context_cell):
        """統一處理位置相關函數"""
        # 支援 ROW, COLUMN, ADDRESS, CELL 等
        pass
```

### 2. 配置選項
```python
INDIRECT_CONFIG = {
    'prefer_excel_mode': True,
    'fallback_to_pure': True,
    'enable_debugging': False,
    'supported_functions': ['ROW', 'COLUMN', 'VLOOKUP']
}
```

### 3. 錯誤處理
```python
class IndirectResolutionError(Exception):
    pass

def safe_resolve_indirect(self, formula):
    try:
        return self.resolve_indirect(formula)
    except IndirectResolutionError as e:
        self.log_error(f"INDIRECT resolution failed: {e}")
        return formula  # 返回原始公式
```

### 4. 性能優化
```python
# 緩存機制
self.function_cache = {}
self.cell_value_cache = {}

def get_cell_value_cached(self, cell_ref):
    if cell_ref not in self.cell_value_cache:
        self.cell_value_cache[cell_ref] = self.calculate_cell_value(cell_ref)
    return self.cell_value_cache[cell_ref]
```

## 測試案例

### 基本測試
```python
test_cases = [
    {
        'formula': '=INDIRECT("A1")',
        'expected': 'A1',
        'description': '簡單儲存格引用'
    },
    {
        'formula': '=INDIRECT("B"&ROW())',
        'expected': 'B5',  # 如果喺第5行
        'description': '動態行引用'
    },
    {
        'formula': '=INDIRECT(VLOOKUP("key",A1:B10,2,FALSE))',
        'expected': 'C3',  # 根據VLOOKUP結果
        'description': '嵌套函數'
    }
]
```

### 跨文件測試
```python
cross_file_tests = [
    {
        'formula': '=INDIRECT("[File2.xlsx]Sheet1!A1")',
        'description': '外部文件引用'
    },
    {
        'formula': '=INDIRECT("'"&B8&"\"&"["&B9&"]"&B10&"'"&"!"&B11)',
        'description': '動態外部文件引用'
    }
]
```

## 限制同注意事項

### Pure模式限制
1. **函數支援有限**：只支援常用函數
2. **跨文件引用**：需要文件存在同可讀取
3. **複雜公式**：可能無法完全模擬Excel行為
4. **循環引用**：無法處理

### Excel模式限制
1. **需要Excel安裝**：服務器環境可能冇
2. **COM依賴**：Windows限定
3. **性能**：大量計算時較慢
4. **穩定性**：Excel崩潰會影響工具

### 未來擴展方向

### 1. 更多函數支援
- 位置函數：ADDRESS(), CELL(), OFFSET()
- 查找函數：INDEX/MATCH, XLOOKUP
- 格式支援：INDIRECT嘅第二個參數（R1C1格式）

### 2. 性能優化
- 並行處理
- 智能緩存
- 增量計算

### 3. 錯誤處理增強
- 詳細錯誤報告
- 部分解析結果
- 建議修復方案

### 4. 用戶界面改進
- 實時預覽
- 步驟式除錯
- 批量處理界面

## 結論

呢個INDIRECT解析工具提供咗一個強大嘅解決方案，可以喺有Excel同冇Excel嘅環境下都能工作。Pure模式通過模擬Excel嘅計算邏輯，實現咗基本嘅INDIRECT解析功能，而Excel模式則提供咗最高嘅準確度。

整合到v70工具時，建議採用自動模式，優先使用Excel模式，失敗時自動切換到Pure模式，咁就可以喺唔同環境下都提供最佳嘅用戶體驗。