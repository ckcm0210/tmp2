# 互動式依賴關係圖 - 開發日誌與最終技術方案

這份文件記錄了為「Excel 公式追蹤工具」開發互動式視覺化依賴關係圖的全過程，包括遇到的問題、解決方案的迭代，以及最終確定的技術實現細節，以供未來開發時參考。

---

## 1. 核心目標

為使用者提供一個比樹狀列表更直觀的、網狀的、可互動的依賴關係圖，以清晰地展示 Excel 儲存格之間的複雜依賴關係。

---

## 2. 技術選型

- **核心函式庫**: `pyvis` (基於 JavaScript 的 `vis.js`)
- **優點**: 能輕易將網路圖資料轉換為可互動的 HTML 檔案，支援豐富的自訂選項。
- **整合方式**: 在工具的「Explode」視窗中增加一個「產生關係圖」按鈕，點擊後執行 Python 腳本產生 HTML 檔案，並使用 `webbrowser` 套件在使用者預設的瀏覽器中打開。

---

## 3. 功能迭代與問題修正日誌

### V1: 初始原型 (力導向佈局)
- **實現**: 使用 `pyvis` 的預設設定產生圖表。
- **問題**: 節點佈局不穩定。拖動一個節點時，所有其他節點都會跟著「浮動」，無法精準定位（此為「力導向佈局」的特性）。

### V2: 佈局穩定化 (禁用物理引擎)
- **目標**: 讓使用者可以隨意拖曳節點，並將其「釘選」在畫布上。
- **實現**: 透過 `net.toggle_physics(False)` 關閉物理引擎。
- **問題**: 雖然節點不再浮動，但初始佈局是完全隨機的，缺乏可讀性。

### V3: 自動化佈局 (階層式佈局)
- **目標**: 讓圖表能自動排列成一個清晰的、由上至下的樹狀結構。
- **實現**: 啟用階層式佈局 `layout: { hierarchical: { enabled: true, direction: "UD" } }`。
- **問題**: 
    1.  **拖曳受限**: 使用者只能在節點所在的「層級」內左右拖動，無法上下移動，不夠自由。
    2.  **方向錯誤 (已修正)**: 初期曾誤設為 `"DU"` (由下至上)，導致圖表上下顛倒。

### V4: 互動性增強與問題浮現
- **目標**: 增加「顯示完整路徑」的互動式核取方塊，並解決拖曳限制。
- **實現**: 
    1.  注入 JavaScript，在圖表繪製完成後 (`afterDrawing` 事件) 禁用階層式佈局，以解除拖曳限制。
    2.  注入 JavaScript，監聽核取方塊事件，以更新節點標籤。
- **問題**: 
    1.  **位置重設**: 點擊核取方塊更新標籤後，階層佈局引擎會被重新觸發，導致使用者精心排列的佈局被重設回預設位置。
    2.  **渲染衝突 (最新 Bug)**: 在拖曳節點的同時，`afterDrawing` 事件可能觸發，導致渲染引擎衝突，圖表閃爍甚至崩潰變白。

---

## 4. 最終方案 (V5): 手動座標佈局 + 穩定互動

經過多次迭代，最終確定採用**完全由 Python 控制佈局**的方案，以獲得完美的穩定性和使用者體驗。此方案徹底解決了先前版本中遇到的所有佈局和互動問題。

### 4.1 核心理念

放棄依賴 `vis.js` 的動態階層佈局引擎 (`hierarchical` layout)，因為它雖然能自動排列，但會帶來拖曳限制和渲染衝突等副作用。取而代之的核心理念是：

1.  **佈局計算在 Python 完成**: 在 Python 腳本中，根據節點的層級 (`level`)，手動計算出每一個節點在初始時應有的、清晰的樹狀 `x` 和 `y` 座標。
2.  **Pyvis 僅做渲染**: `pyvis` 不再負責佈局，只負責根據我們提供的精確座標和樣式，將圖表「畫」出來。
3.  **JavaScript 負責互動**: 注入的 JavaScript 只處理使用者的互動事件（如點擊核取方塊），並在更新標籤時，**強制節點保持其當前位置**，從而實現完美的穩定性。

### 4.2 最終版 Python 腳本 (`generate_stable_graph.py`)

以下是產生最終版完美互動圖的完整 Python 程式碼。

```python
import os
from pyvis.network import Network
import json

# --- 1. 定義節點和邊的數據 ---
nodes_data = [
    # Level 0
    {"id": 0, "short_label": "Final_Report!A1\nFormula: =SUM(Data!B1:Data!B3)\nValue: 7825.0", "full_label": "C:\\My Reports\\[Master.xlsx]Final_Report!A1\nFormula: =SUM(Data!B1:Data!B3)\nValue: 7825.0", "color": "#e04141", "level": 0},
    # Level 1
    {"id": 1, "short_label": "Data!B1\nFormula: =VLOOKUP(C1, RefData!A:B, 2, 0)\nValue: 500.0", "full_label": "C:\\My Reports\\[Master.xlsx]Data!B1\nFormula: =VLOOKUP(C1, RefData!A:B, 2, 0)\nValue: 500.0", "color": "#007bff", "level": 1},
    {"id": 2, "short_label": "Data!B2\nFormula: =[External.xlsx]Sheet1!D5 * 1.05\nValue: 2100.0", "full_label": "C:\\My Reports\\[Master.xlsx]Sheet1!D5 * 1.05\nValue: 2100.0", "color": "#007bff", "level": 1},
    {"id": 3, "short_label": "Data!B3\nType: Value\nValue: 5225.0", "full_label": "C:\\My Reports\\[Master.xlsx]Data!B3\nType: Value\nValue: 5225.0", "color": "#28a745", "level": 1},
    # Level 2
    {"id": 4, "short_label": "Data!C1\nType: Value\nValue: SKU-001", "full_label": "C:\\My Reports\\[Master.xlsx]Data!C1\nType: Value\nValue: SKU-001", "color": "#28a745", "level": 2},
    {"id": 5, "short_label": "RefData!B1\nType: Value\nValue: 500.0", "full_label": "C:\\My Reports\\[Master.xlsx]RefData!B1\nType: Value\nValue: 500.0", "color": "#28a745", "level": 2},
    {"id": 6, "short_label": "[External.xlsx]Sheet1!D5\nType: External Value\nValue: 2000.0", "full_label": "C:\\Linked Files\\[External.xlsx]Sheet1!D5\nType: External Value\nValue: 2000.0", "color": "#ff8c00", "level": 2}
]
edges_data = [(0, 1), (0, 2), (0, 3), (1, 4), (1, 5), (2, 6)]

# --- 2. 手動計算階層式佈局的初始座標 ---
level_counts = {}
for node in nodes_data:
    level = node['level']
    if level not in level_counts:
        level_counts[level] = 0
    level_counts[level] += 1

node_positions = {}
level_y_step = 250
level_x_step = 350

current_level_counts = {level: 0 for level in level_counts}

for node in nodes_data:
    level = node['level']
    total_in_level = level_counts[level]
    current_index_in_level = current_level_counts[level]
    
    # 計算座標
    y = level * level_y_step
    x = (current_index_in_level - (total_in_level - 1) / 2.0) * level_x_step
    
    node['x'] = x
    node['y'] = y
    current_level_counts[level] += 1

# --- 3. 使用 Pyvis 產生圖表 ---
net = Network(height="90vh", width="100%", bgcolor="#ffffff", font_color="black", directed=True)

# 關鍵：不使用階層式佈局，直接設定全域選項
options_str = '''
{
  "interaction": {
    "dragNodes": true,
    "dragView": false,
    "zoomView": true
  },
  "physics": {
    "enabled": false
  },
  "nodes": {
    "font": {
      "align": "left"
    }
  },
  "edges": {
    "smooth": {
      "type": "cubicBezier",
      "forceDirection": "vertical",
      "roundness": 0.4
    }
  }
}
'''
net.set_options(options_str)

# 將節點資料（包含計算好的 x, y 座標）加入網路圖
for node_info in nodes_data:
    net.add_node(
        node_info["id"],
        label=node_info["short_label"],
        shape='box',
        color=node_info["color"],
        x=node_info['x'],
        y=node_info['y'],
        fixed=False, # 允許使用者拖曳
        full_label=node_info["full_label"],
        short_label=node_info["short_label"]
    )

for edge in edges_data:
    net.add_edge(edge[0], edge[1])

# --- 4. 注入 HTML 和 JavaScript ---
temp_file = "temp_stable_graph.html"
net.save_graph(temp_file)

with open(temp_file, 'r', encoding='utf-8') as f:
    html_content = f.read()

checkbox_html = """
<div style='position: absolute; top: 10px; left: 10px; background: #f8f9fa; padding: 10px; border: 1px solid #dee2e6; border-radius: 5px; z-index: 1000;'>
  <label for='pathToggle' style='font-family: sans-serif; font-size: 14px;'>
    <input type='checkbox' id='pathToggle'>
    顯示完整檔案路徑
  </label>
</div>
"""

# 關鍵：使用最穩定的 JavaScript 版本
javascript_injection = """
<script type='text/javascript'>
  document.addEventListener('DOMContentLoaded', function() {
    var network = window.network;
    var nodes = window.nodes;
    if (!network || !nodes) { return; }

    const pathToggle = document.getElementById('pathToggle');
    pathToggle.addEventListener('change', function() {
      const showFullPath = this.checked;
      const currentPositions = network.getPositions();
      let updatedNodes = [];

      nodes.forEach(node => {
        let newLabel = showFullPath ? node.full_label : node.short_label;
        const position = currentPositions[node.id];
        
        updatedNodes.push({
          id: node.id,
          label: newLabel,
          x: position.x,
          y: position.y,
          fixed: true // 點擊後固定，防止微小移動
        });
      });
      
      nodes.update(updatedNodes);

      // 短暫延遲後，解除固定，恢復自由拖曳
      setTimeout(function() {
          let releaseNodes = [];
          nodes.forEach(node => {
              releaseNodes.push({id: node.id, fixed: false});
          });
          nodes.update(releaseNodes);
      }, 100);
    });
  });
</script>
"""

html_content = html_content.replace('<body>', '<body>\n' + checkbox_html)
html_content = html_content.replace('</body>', javascript_injection + '\n</body>')

final_file = "stable_interactive_graph.html"
with open(final_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

os.remove(temp_file)

print(f"Successfully generated stable interactive graph at: {os.path.join(os.getcwd(), final_file)}")
```

### 4.3 關鍵設定與 JavaScript 詳解

- **`options_str` (Pyvis 設定)**:
    - `interaction`: 明確設定 `dragNodes: true` (允許拖動節點) 和 `dragView: false` (禁止拖動整個畫布)，避免使用者誤觸導致整個圖移動。
    - `physics`: `enabled: false`，完全禁用物理引擎，這是穩定性的基礎。
    - `nodes`: `font: { align: 'left' }`，實現節點內文字靠左對齊。
    - `edges`: `smooth: { type: 'cubicBezier', ... }`，產生平滑的曲線箭頭。

- **`javascript_injection` (互動邏輯)**:
    - `const currentPositions = network.getPositions();`: 在更新標籤**之前**，先獲取並儲存所有節點**當前**的、由使用者拖曳後決定的最新位置。
    - `updatedNodes.push({ ... x: position.x, y: position.y, fixed: true });`: 在 `nodes.update()` 指令中，不僅提供新的 `label`，還**強制**將節點的 `x` 和 `y` 座標設定為它之前的位置，並臨時設定 `fixed: true`，這是防止位置重設的核心步驟。
    - `setTimeout(...)`: 在更新完成後，透過一個微小的延遲 (100毫秒)，再將所有節點的 `fixed` 屬性改回 `false`。這一步是為了確保在標籤更新的瞬間節點位置被鎖定，但之後使用者又能恢復完全自由的拖曳能力。

### 4.4 未來整合建議

1.  **建立一個 `GraphGenerator` 類別**: 將 `generate_stable_graph.py` 的邏輯封裝到一個類別中。
2.  **`__init__`**: 接收 `dependency_tree_data` 作為輸入。
3.  **`convert_to_pyvis_data()`**: 編寫一個方法，遍歷 `dependency_tree_data`，將其轉換為 `pyvis` 需要的 `nodes_data` 和 `edges_data` 格式，並在這個過程中計算好每個節點的 `level`。
4.  **`generate_html()`**: 執行產生圖表、注入 JS/HTML 並儲存檔案的邏輯。
5.  **在主工具中呼叫**: 在 `worksheet_tree.py` 的「產生關係圖」按鈕的事件處理函數中，實例化這個 `GraphGenerator` 類別，呼叫其方法，最後用 `webbrowser.open()` 打開產生的 HTML 檔案。

```