# INDIRECT 函數解析文檔

## 📋 解析流程概述

### 基本原理
INDIRECT 函數會動態建構儲存格引用，需要在運行時解析其參數來確定真實的引用目標。

### 解析步驟

#### 1. 識別階段
- 掃描公式中的 INDIRECT 函數
- 提取 INDIRECT 的參數表達式
- 例：`=INDIRECT(D29&"!"&D26)` → 參數：`D29&"!"&D26`

#### 2. 參數分析階段
- 解析參數中的儲存格引用
- 例：`D29&"!"&D26` → 引用：`[D29, D26]`

#### 3. 值讀取階段
- 讀取引用儲存格的當前值
- 例：`D29 = "工作表2"`, `D26 = "A2"`

#### 4. 重構階段
- 將值代入參數表達式
- 評估字串連接操作
- 例：`"工作表2" & "!" & "A2"` → `"工作表2!A2"`

#### 5. 驗證階段
- 檢查重構的引用是否有效
- 確認目標儲存格是否存在

## 🔧 技術實作細節

### 正則表達式模式
```python
# INDIRECT 函數識別
INDIRECT_PATTERN = r'INDIRECT\s*\(\s*([^)]+)\s*\)'

# 儲存格引用識別
CELL_REF_PATTERNS = [
    r'\b[A-Z]+\d+\b',           # A1, B2
    r'\$[A-Z]+\$\d+',           # $A$1
    r'[^!]+![A-Z]+\d+',         # Sheet1!A1
]
```

### 字串連接評估
```python
def evaluate_concatenation(expression, values):
    """評估包含 & 運算符的字串連接"""
    parts = expression.split('&')
    result = ""
    for part in parts:
        part = part.strip()
        if part in values:
            result += str(values[part])
        elif part.startswith('"') and part.endswith('"'):
            result += part[1:-1]  # 移除引號
        else:
            result += part
    return result
```

## 📊 解析範例

### 範例 1：基本工作表引用
```excel
原始公式: =INDIRECT(D29&"!"&D26)
參數: D29&"!"&D26
引用: [D29, D26]
值: D29="工作表2", D26="A2"
解析結果: 工作表2!A2
```

### 範例 2：動態列引用
```excel
原始公式: =SUM(INDIRECT("A1:A"&B5))
參數: "A1:A"&B5
引用: [B5]
值: B5=100
解析結果: A1:A100
```

### 範例 3：複雜巢狀引用
```excel
原始公式: =INDIRECT("'"&C10&"'!B"&ROW())
參數: "'"&C10&"'!B"&ROW()
引用: [C10]
值: C10="銷售數據", ROW()=15
解析結果: '銷售數據'!B15
```

## ⚠️ 錯誤處理情況

### 常見錯誤類型
1. **引用儲存格為空**
   - 原因：參數中引用的儲存格沒有值
   - 處理：顯示警告，保留原始 INDIRECT 公式

2. **無效的引用格式**
   - 原因：重構後的引用格式不正確
   - 處理：標記為解析錯誤

3. **循環引用**
   - 原因：INDIRECT 指向包含它的儲存格
   - 處理：檢測並中斷循環

4. **跨檔案引用**
   - 原因：INDIRECT 指向其他檔案
   - 處理：嘗試解析，失敗則保留原始公式

### 錯誤處理策略
```python
def safe_resolve_indirect(formula, worksheet):
    try:
        return resolve_indirect(formula, worksheet)
    except IndirectError as e:
        return {
            'status': 'error',
            'original': formula,
            'error': str(e),
            'fallback': f"INDIRECT(...) - 無法解析"
        }
```

## 🎯 主程式整合要點

### 顯示格式
在主程式中，INDIRECT 解析結果應該簡潔顯示：

```
原始: =INDIRECT(D29&"!"&D26)
解析: 工作表2!A2
```

### 依賴關係處理
- 將 INDIRECT 視為特殊的依賴類型
- 在依賴樹中同時顯示原始公式和解析結果
- 繼續分析解析後的引用的依賴關係

### 性能考量
- 緩存解析結果避免重複計算
- 限制解析深度防止無限遞迴
- 批量處理多個 INDIRECT 函數

## 📈 未來擴展

### 支援更多動態函數
- OFFSET：動態偏移引用
- INDEX/MATCH：動態查找引用
- CHOOSE：條件選擇引用

### 智能提示
- 檢測常見的 INDIRECT 模式
- 提供優化建議
- 警告潛在的性能問題

## 🔍 測試案例

### 基本測試
```excel
A1: Sheet2
A2: B5
A3: =INDIRECT(A1&"!"&A2)  → 應解析為: Sheet2!B5
```

### 複雜測試
```excel
B1: Sales
B2: 2024
B3: =INDIRECT("'"&B1&B2&"'!C1:C10")  → 應解析為: 'Sales2024'!C1:C10
```

### 錯誤測試
```excel
C1: (空白)
C2: =INDIRECT(C1&"!A1")  → 應顯示解析錯誤
```

這份文檔記錄了 INDIRECT 解析的完整流程和技術細節，可作為實作和維護的參考。