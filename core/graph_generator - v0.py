# graph_generator.py

import os
import webbrowser
import json

class GraphGenerator:
    def __init__(self, nodes_data, edges_data):
        self.nodes_data = nodes_data
        self.edges_data = edges_data
        self.output_filename = "dependency_graph.html"

    def generate_graph(self):
        """
        生成完全獨立的 HTML 文件，所有資源都內嵌，可在受限瀏覽器中使用
        """
        # 1. 手動計算階層式佈局的初始座標
        self._calculate_node_positions()

        # 2. 生成獨立 HTML 內容
        html_content = self._generate_standalone_html()

        # 3. 保存最終文件
        final_file_path = os.path.join(os.getcwd(), self.output_filename)
        
        try:
            with open(final_file_path, 'w', encoding='utf-8', errors='replace') as f:
                f.write(html_content)
            print(f"Successfully generated standalone graph at: {final_file_path}")
        except Exception as e:
            print(f"Error saving file: {e}")
            return

        # 4. 在瀏覽器中打開
        webbrowser.open(f"file://{final_file_path}")

    def _generate_standalone_html(self):
        """
        生成完全獨立的 HTML，包含所有內嵌資源
        """
        # 準備節點和邊數據
        processed_nodes = []
        for node in self.nodes_data:
            processed_node = {
                "id": self._safe_string(node["id"]),
                "label": self._safe_string(node["label"]),
                "title": self._safe_string(node["title"]),
                "color": node["color"],
                "shape": "box",
                "x": node.get('x', 0),
                "y": node.get('y', 0),
                "fixed": False,
                "font": {"color": "black"},
                "filename": self._safe_string(node.get('filename', 'Current File')),
                "short_address_label": self._safe_string(node["short_address_label"]),
                "full_address_label": self._safe_string(node["full_address_label"]),
                "short_formula_label": self._safe_string(node["short_formula_label"]),
                "full_formula_label": self._safe_string(node["full_formula_label"]),
                "value_label": self._safe_string(node["value_label"])
            }
            processed_nodes.append(processed_node)

        processed_edges = []
        for edge in self.edges_data:
            processed_edge = {
                "arrows": "to",
                "from": self._safe_string(edge[0]),
                "to": self._safe_string(edge[1])
            }
            processed_edges.append(processed_edge)

        # 安全編碼 JSON
        nodes_json = self._safe_json_encode(processed_nodes)
        edges_json = self._safe_json_encode(processed_edges)

        print(f"Processing {len(processed_nodes)} nodes and {len(processed_edges)} edges")

        # 內嵌的 vis.js（完整版本 - 修正節點尺寸問題）
        vis_js_content = """
        // Complete vis.js implementation for network visualization
        var vis = (function() {
            
            function DataSet(data) {
                this.data = data || [];
                this.length = this.data.length;
            }
            
            DataSet.prototype.get = function(options) {
                if (options && options.returnType === "Object") {
                    var result = {};
                    this.data.forEach(item => {
                        result[item.id] = item;
                    });
                    return result;
                } else if (options && options.returnType === "Array") {
                    return this.data.slice();
                }
                return this.data.slice();
            };
            
            DataSet.prototype.update = function(updates) {
                if (!Array.isArray(updates)) {
                    updates = [updates];
                }
                
                updates.forEach(update => {
                    var index = this.data.findIndex(item => item.id === update.id);
                    if (index !== -1) {
                        Object.assign(this.data[index], update);
                    }
                });
            };
            
            function Network(container, data, options) {
                this.container = container;
                this.nodes = data.nodes;
                this.edges = data.edges;
                this.options = options || {};
                this.canvas = null;
                this.ctx = null;
                this.nodePositions = {};
                this.nodeSizes = {}; // 新增：儲存每個節點的計算尺寸
                this.isDragging = false;
                this.dragNode = null;
                this.dragOffset = {x: 0, y: 0};
                this.viewOffset = {x: 0, y: 0};
                this.isDraggingView = false;
                this.lastMousePos = {x: 0, y: 0};
                this.scale = 1;
                
                this.init();
            }
            
            Network.prototype.init = function() {
                this.canvas = document.createElement('canvas');
                this.canvas.width = this.container.clientWidth;
                this.canvas.height = this.container.clientHeight;
                this.canvas.style.display = 'block';
                this.canvas.style.cursor = 'grab';
                this.container.appendChild(this.canvas);
                this.ctx = this.canvas.getContext('2d');
                
                // 監聽窗口大小變化
                window.addEventListener('resize', () => {
                    this.canvas.width = this.container.clientWidth;
                    this.canvas.height = this.container.clientHeight;
                    this.draw();
                });
                
                // 初始化節點位置和尺寸
                var nodes = this.nodes.get();
                nodes.forEach(node => {
                    this.nodePositions[node.id] = {
                        x: (node.x || 0) + this.canvas.width / 2,
                        y: (node.y || 0) + this.canvas.height / 2
                    };
                    // 計算每個節點的尺寸
                    this.calculateNodeSize(node);
                });
                
                this.draw();
                this.setupEvents();
            };
            
            Network.prototype.calculateNodeSize = function(node) {
                // 創建臨時canvas來測量文字尺寸
                var tempCanvas = document.createElement('canvas');
                var tempCtx = tempCanvas.getContext('2d');
                
                var fontSize = (node.font && node.font.size) || 14;
                tempCtx.font = fontSize + 'px Arial';
                
                // 處理標籤文字
                var label = node.label || node.id;
                var lines = label.split('\\n');
                
                var maxWidth = 0;
                var totalHeight = 0;
                var lineHeight = fontSize + 4; // 增加行間距
                var padding = 20; // 內邊距
                
                lines.forEach((line, index) => {
                    // 移除HTML標籤來測量實際文字寬度
                    var cleanLine = line.replace(/<[^>]*>/g, '');
                    
                    // 測量這行文字的寬度
                    var lineWidth = tempCtx.measureText(cleanLine).width;
                    maxWidth = Math.max(maxWidth, lineWidth);
                    totalHeight += lineHeight;
                });
                
                // 設定最小和最大尺寸
                var minWidth = 150;
                var maxWidthLimit = 400;
                var minHeight = 60;
                
                var calculatedWidth = Math.max(minWidth, Math.min(maxWidth + padding * 2, maxWidthLimit));
                var calculatedHeight = Math.max(minHeight, totalHeight + padding);
                
                // 儲存計算出的尺寸
                this.nodeSizes[node.id] = {
                    width: calculatedWidth,
                    height: calculatedHeight
                };
            };
            
            Network.prototype.draw = function() {
                var ctx = this.ctx;
                ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
                
                // 應用視圖變換
                ctx.save();
                ctx.scale(this.scale, this.scale);
                ctx.translate(this.viewOffset.x / this.scale, this.viewOffset.y / this.scale);
                
                // 繪製邊
                var edges = this.edges.get();
                edges.forEach(edge => {
                    var fromPos = this.nodePositions[edge.from];
                    var toPos = this.nodePositions[edge.to];
                    
                    if (fromPos && toPos) {
                        ctx.beginPath();
                        ctx.moveTo(fromPos.x, fromPos.y);
                        ctx.lineTo(toPos.x, toPos.y);
                        ctx.strokeStyle = '#848484';
                        ctx.lineWidth = 1;
                        ctx.stroke();
                        
                        // 繪製箭頭
                        this.drawArrow(ctx, fromPos.x, fromPos.y, toPos.x, toPos.y, edge.from, edge.to);
                    }
                });
                
                // 繪製節點
                var nodes = this.nodes.get();
                nodes.forEach(node => {
                    var pos = this.nodePositions[node.id];
                    if (pos) {
                        this.drawNode(ctx, node, pos.x, pos.y);
                    }
                });
                
                ctx.restore();
            };
            
            Network.prototype.drawNode = function(ctx, node, x, y) {
                // 使用計算出的尺寸
                var nodeSize = this.nodeSizes[node.id];
                if (!nodeSize) {
                    this.calculateNodeSize(node);
                    nodeSize = this.nodeSizes[node.id];
                }
                
                var width = nodeSize.width;
                var height = nodeSize.height;
                
                // 繪製節點背景
                ctx.fillStyle = node.color || '#97C2FC';
                ctx.fillRect(x - width/2, y - height/2, width, height);
                
                // 繪製邊框
                ctx.strokeStyle = '#2B7CE9';
                ctx.lineWidth = 1;
                ctx.strokeRect(x - width/2, y - height/2, width, height);
                
                // 繪製文字
                ctx.fillStyle = 'black';
                var fontSize = (node.font && node.font.size) || 14;
                ctx.textAlign = 'left';
                ctx.textBaseline = 'top';
                
                // 處理多行文字和HTML標籤
                var label = node.label || node.id;
                var lines = label.split('\\n');
                var lineHeight = fontSize + 4;
                var padding = 10;
                var startY = y - height/2 + padding;
                var startX = x - width/2 + padding;
                var maxLineWidth = width - padding * 2;
                
                var currentY = startY;
                
                lines.forEach((line, index) => {
                    // 移除HTML標籤但保留格式信息
                    var cleanLine = line.replace(/<b>(.*?)<\\/b>/g, '$1');
                    cleanLine = cleanLine.replace(/<i>(.*?)<\\/i>/g, '$1');
                    cleanLine = cleanLine.replace(/<[^>]*>/g, '');
                    
                    // 檢查格式
                    var isBold = line.includes('<b>');
                    var isItalic = line.includes('<i>');
                    
                    // 設定字體樣式
                    var fontStyle = '';
                    if (isBold && isItalic) {
                        fontStyle = 'bold italic ';
                    } else if (isBold) {
                        fontStyle = 'bold ';
                    } else if (isItalic) {
                        fontStyle = 'italic ';
                    }
                    ctx.font = fontStyle + fontSize + 'px Arial';
                    
                    // 文字換行處理
                    var words = cleanLine.split(' ');
                    var currentLine = '';
                    
                    for (var i = 0; i < words.length; i++) {
                        var testLine = currentLine + words[i] + ' ';
                        var metrics = ctx.measureText(testLine);
                        var testWidth = metrics.width;
                        
                        if (testWidth > maxLineWidth && i > 0) {
                            ctx.fillText(currentLine.trim(), startX, currentY);
                            currentLine = words[i] + ' ';
                            currentY += lineHeight;
                        } else {
                            currentLine = testLine;
                        }
                    }
                    
                    // 繪製最後一行
                    if (currentLine.trim()) {
                        ctx.fillText(currentLine.trim(), startX, currentY);
                        currentY += lineHeight;
                    }
                    
                    // 為不同段落增加額外間距
                    if (index < lines.length - 1) {
                        currentY += 4;
                    }
                });
            };
            
            Network.prototype.drawArrow = function(ctx, fromX, fromY, toX, toY, fromNodeId, toNodeId) {
                var angle = Math.atan2(toY - fromY, toX - fromX);
                var length = 10;
                
                // 使用目標節點的實際尺寸來計算箭頭位置
                var toNodeSize = this.nodeSizes[toNodeId];
                var nodeWidth = toNodeSize ? toNodeSize.width : 200;
                var nodeHeight = toNodeSize ? toNodeSize.height : 100;
                
                // 計算到節點邊緣的距離
                var dx = Math.abs(Math.cos(angle)) * nodeWidth / 2;
                var dy = Math.abs(Math.sin(angle)) * nodeHeight / 2;
                var edgeDistance = Math.max(dx, dy);
                
                var distance = Math.sqrt((toX - fromX) * (toX - fromX) + (toY - fromY) * (toY - fromY));
                var ratio = Math.max(0, (distance - edgeDistance) / distance);
                var arrowX = fromX + (toX - fromX) * ratio;
                var arrowY = fromY + (toY - fromY) * ratio;
                
                ctx.beginPath();
                ctx.moveTo(arrowX, arrowY);
                ctx.lineTo(arrowX - length * Math.cos(angle - Math.PI / 6), 
                          arrowY - length * Math.sin(angle - Math.PI / 6));
                ctx.moveTo(arrowX, arrowY);
                ctx.lineTo(arrowX - length * Math.cos(angle + Math.PI / 6), 
                          arrowY - length * Math.sin(angle + Math.PI / 6));
                ctx.strokeStyle = '#848484';
                ctx.lineWidth = 1;
                ctx.stroke();
            };
            
            Network.prototype.setupEvents = function() {
                var self = this;
                
                // 鼠標按下事件
                this.canvas.addEventListener('mousedown', function(e) {
                    var rect = self.canvas.getBoundingClientRect();
                    var mouseX = (e.clientX - rect.left - self.viewOffset.x) / self.scale;
                    var mouseY = (e.clientY - rect.top - self.viewOffset.y) / self.scale;
                    
                    self.lastMousePos = {x: e.clientX - rect.left, y: e.clientY - rect.top};
                    
                    // 檢查是否點擊在節點上
                    var nodes = self.nodes.get();
                    var nodeClicked = false;
                    
                    for (var i = 0; i < nodes.length; i++) {
                        var node = nodes[i];
                        var pos = self.nodePositions[node.id];
                        var nodeSize = self.nodeSizes[node.id];
                        
                        if (!nodeSize) {
                            self.calculateNodeSize(node);
                            nodeSize = self.nodeSizes[node.id];
                        }
                        
                        var width = nodeSize.width;
                        var height = nodeSize.height;
                        
                        if (pos && 
                            mouseX >= pos.x - width/2 && mouseX <= pos.x + width/2 &&
                            mouseY >= pos.y - height/2 && mouseY <= pos.y + height/2) {
                            self.isDragging = true;
                            self.dragNode = node.id;
                            self.dragOffset = {
                                x: mouseX - pos.x,
                                y: mouseY - pos.y
                            };
                            self.canvas.style.cursor = 'grabbing';
                            nodeClicked = true;
                            break;
                        }
                    }
                    
                    // 如果沒有點擊節點，則開始拖拽視圖
                    if (!nodeClicked) {
                        self.isDraggingView = true;
                        self.canvas.style.cursor = 'grabbing';
                    }
                });
                
                // 鼠標移動事件
                this.canvas.addEventListener('mousemove', function(e) {
                    var rect = self.canvas.getBoundingClientRect();
                    var mouseX = (e.clientX - rect.left - self.viewOffset.x) / self.scale;
                    var mouseY = (e.clientY - rect.top - self.viewOffset.y) / self.scale;
                    var currentMousePos = {x: e.clientX - rect.left, y: e.clientY - rect.top};
                    
                    if (self.isDragging && self.dragNode) {
                        // 拖拽節點
                        self.nodePositions[self.dragNode] = {
                            x: mouseX - self.dragOffset.x,
                            y: mouseY - self.dragOffset.y
                        };
                        self.draw();
                    } else if (self.isDraggingView) {
                        // 拖拽整個視圖
                        var deltaX = currentMousePos.x - self.lastMousePos.x;
                        var deltaY = currentMousePos.y - self.lastMousePos.y;
                        
                        self.viewOffset.x += deltaX;
                        self.viewOffset.y += deltaY;
                        
                        self.draw();
                    }
                    
                    self.lastMousePos = currentMousePos;
                });
                
                // 鼠標釋放事件
                this.canvas.addEventListener('mouseup', function(e) {
                    self.isDragging = false;
                    self.isDraggingView = false;
                    self.dragNode = null;
                    self.canvas.style.cursor = 'grab';
                });
                
                // 滾輪縮放事件
                this.canvas.addEventListener('wheel', function(e) {
                    e.preventDefault();
                    
                    var rect = self.canvas.getBoundingClientRect();
                    var mouseX = e.clientX - rect.left;
                    var mouseY = e.clientY - rect.top;
                    
                    var scaleFactor = e.deltaY > 0 ? 0.9 : 1.1;
                    var newScale = self.scale * scaleFactor;
                    
                    // 限制縮放範圍
                    newScale = Math.max(0.1, Math.min(5, newScale));
                    
                    if (newScale !== self.scale) {
                        // 計算縮放中心點
                        var scaleChange = newScale / self.scale;
                        
                        self.viewOffset.x = mouseX - (mouseX - self.viewOffset.x) * scaleChange;
                        self.viewOffset.y = mouseY - (mouseY - self.viewOffset.y) * scaleChange;
                        
                        self.scale = newScale;
                        self.draw();
                    }
                });
            };
            
            Network.prototype.getPositions = function() {
                var result = {};
                for (var nodeId in this.nodePositions) {
                    result[nodeId] = {
                        x: this.nodePositions[nodeId].x,
                        y: this.nodePositions[nodeId].y
                    };
                }
                return result;
            };
            
            return {
                DataSet: DataSet,
                Network: Network
            };
        })();
        """

        html_template = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Dependency Graph - Standalone</title>
    
    <style type="text/css">
        body {{
            margin: 0;
            padding: 0;
            font-family: Arial, sans-serif;
            background-color: #f5f5f5;
        }}
        
        #mynetwork {{
            width: 100%;
            height: 100vh;
            background-color: #ffffff;
            border: 1px solid lightgray;
            position: relative;
        }}
        
        .controls {{
            position: absolute;
            top: 10px;
            left: 10px;
            background: rgba(248, 249, 250, 0.95);
            padding: 12px;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            z-index: 1000;
            font-family: sans-serif;
            font-size: 14px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            min-width: 200px;
        }}
        
        .controls h4 {{
            margin: 0 0 10px 0;
            font-weight: bold;
            color: #333;
        }}
        
        .control-item {{
            margin-bottom: 8px;
        }}
        
        .control-item label {{
            cursor: pointer;
            display: flex;
            align-items: center;
        }}
        
        .control-item input[type="checkbox"] {{
            margin-right: 8px;
        }}
        
        .slider-container {{
            margin-bottom: 8px;
        }}
        
        .slider-container label {{
            display: block;
            margin-bottom: 4px;
            cursor: pointer;
        }}
        
        .slider-container input[type="range"] {{
            width: 100%;
        }}
        
        .legend {{
            border-top: 1px solid #ccc;
            margin-top: 12px;
            padding-top: 10px;
        }}
        
        .legend h5 {{
            margin: 0 0 6px 0;
            font-weight: bold;
            color: #333;
        }}
        
        .legend-item {{
            display: flex;
            align-items: center;
            margin-bottom: 4px;
        }}
        
        .legend-color {{
            width: 16px;
            height: 16px;
            margin-right: 8px;
            border-radius: 3px;
            border: 1px solid #ddd;
        }}
        
        .legend-text {{
            font-size: 12px;
            font-weight: 500;
            color: #333;
        }}
        
        .legend-help {{
            margin-top: 8px;
            font-size: 11px;
            color: #666;
        }}
    </style>
</head>

<body>
    <!-- 控制面板 -->
    <div class="controls">
        <h4>Display Options</h4>
        
        <div class="control-item">
            <label>
                <input type='checkbox' id='formulaToggle'> Show Full Formula Path
            </label>
        </div>
        
        <div class="control-item">
            <label>
                <input type='checkbox' id='addressToggle'> Show Full Address Path
            </label>
        </div>
        
        <div class="slider-container">
            <label for='fontSizeSlider'>
                Font Size: <span id='fontSizeValue'>14</span>px
            </label>
            <input type='range' id='fontSizeSlider' min='10' max='24' value='14'>
        </div>
        
        <div class="legend">
            <h5>File Legend</h5>
            <div id='fileLegend'>
                <!-- 動態生成的檔案圖例 -->
            </div>
            <div class="legend-help">
                相同顏色 = 同一檔案<br>
                不同顏色間的箭頭 = 跨檔案依賴
            </div>
        </div>
    </div>
    
    <!-- 圖表容器 -->
    <div id="mynetwork"></div>

    <!-- 內嵌的 vis.js -->
    <script type="text/javascript">
        {vis_js_content}
    </script>

    <!-- 主要應用邏輯 -->
    <script type="text/javascript">
        // 全局變量
        var nodes;
        var edges;
        var network;
        var nodeData = {nodes_json};
        var edgeData = {edges_json};

        // 初始化圖表
        function initGraph() {{
            console.log('Initializing graph with', nodeData.length, 'nodes and', edgeData.length, 'edges');
            
            var container = document.getElementById('mynetwork');
            
            nodes = new vis.DataSet(nodeData);
            edges = new vis.DataSet(edgeData);
            
            var data = {{
                nodes: nodes,
                edges: edges
            }};
            
            var options = {{
                interaction: {{
                    dragNodes: true,
                    dragView: true,
                    zoomView: true
                }},
                physics: {{
                    enabled: false
                }}
            }};
            
            network = new vis.Network(container, data, options);
            
            console.log('Graph initialized successfully');
            
            // 初始化控制項
            initControls();
        }}
        
        // 初始化控制項
        function initControls() {{
            var formulaToggle = document.getElementById('formulaToggle');
            var addressToggle = document.getElementById('addressToggle');
            var fontSizeSlider = document.getElementById('fontSizeSlider');
            var fontSizeValue = document.getElementById('fontSizeValue');
            
            function updateNodeLabels() {{
                if (!network || !nodes) return;
                
                var showFullAddress = addressToggle.checked;
                var showFullFormula = formulaToggle.checked;
                var fontSize = parseInt(fontSizeSlider.value);
                
                var allNodes = nodes.get();
                var updatedNodes = [];
                
                allNodes.forEach(function(node) {{
                    var addressLabel = showFullAddress ? node.full_address_label : node.short_address_label;
                    var formulaLabel = showFullFormula ? node.full_formula_label : node.short_formula_label;
                    
                    // 保持原始格式：Address : <b>...</b>
                    var newLabel = 'Address : <b>' + (addressLabel || node.short_address_label) + '</b>';
                    
                    if (formulaLabel && formulaLabel !== 'N/A' && formulaLabel !== null) {{
                        var displayFormula = formulaLabel.indexOf('=') === 0 ? formulaLabel : '=' + formulaLabel;
                        newLabel += '\\n\\nFormula : <i>' + displayFormula + '</i>';
                    }} else {{
                        newLabel += '\\n\\nFormula : <i>N/A</i>';
                    }}
                    
                    newLabel += '\\n\\nValue     : ' + (node.value_label || 'N/A');
                    
                    updatedNodes.push({{
                        id: node.id,
                        label: newLabel,
                        font: {{ size: fontSize }}
                    }});
                }});
                
                if (updatedNodes.length > 0) {{
                    nodes.update(updatedNodes);
                    // 重新計算所有節點尺寸
                    var allNodes = nodes.get();
                    allNodes.forEach(function(node) {{
                        network.calculateNodeSize(node);
                    }});
                    network.draw();
                }}
            }}
            
            function updateFontSize() {{
                var fontSize = parseInt(fontSizeSlider.value);
                fontSizeValue.textContent = fontSize;
                updateNodeLabels();
            }}
            
            function generateFileLegend() {{
                var fileLegendDiv = document.getElementById('fileLegend');
                if (!fileLegendDiv || !nodes) return;
                
                var fileColors = new Map();
                var allNodes = nodes.get();
                
                allNodes.forEach(function(node) {{
                    var color = node.color || '#808080';
                    var filename = node.filename || 'Unknown File';
                    
                    if (!fileColors.has(filename)) {{
                        fileColors.set(filename, color);
                    }}
                }});
                
                var sortedFiles = Array.from(fileColors.entries()).sort(function(a, b) {{
                    if (a[0] === 'Current File') return -1;
                    if (b[0] === 'Current File') return 1;
                    return a[0].localeCompare(b[0]);
                }});
                
                var legendHTML = '';
                sortedFiles.forEach(function(item) {{
                    var filename = item[0];
                    var color = item[1];
                    legendHTML += '<div class="legend-item" title="檔案: ' + filename + '">';
                    legendHTML += '<div class="legend-color" style="background-color: ' + color + ';"></div>';
                    legendHTML += '<span class="legend-text">' + filename + '</span>';
                    legendHTML += '</div>';
                }});
                
                fileLegendDiv.innerHTML = legendHTML;
            }}
            
            // 綁定事件
            if (addressToggle) addressToggle.addEventListener('change', updateNodeLabels);
            if (formulaToggle) formulaToggle.addEventListener('change', updateNodeLabels);
            if (fontSizeSlider) fontSizeSlider.addEventListener('input', updateFontSize);
            
            // 生成圖例
            generateFileLegend();
        }}
        
        // 頁面加載完成後初始化
        window.addEventListener('load', function() {{
            initGraph();
        }});
    </script>
</body>
</html>"""
        
        return html_template

    def _safe_string(self, value):
        """
        安全地處理字符串，移除可能導致問題的字符
        """
        if value is None:
            return ""
        
        str_value = str(value)
        # 移除可能導致編碼問題的字符
        try:
            return str_value.encode('utf-8', errors='ignore').decode('utf-8')
        except:
            return str_value.encode('ascii', errors='ignore').decode('ascii')

    def _safe_json_encode(self, data):
        """
        安全地將數據編碼為 JSON
        """
        try:
            return json.dumps(data, ensure_ascii=False, separators=(',', ':'))
        except Exception as e:
            print(f"JSON encoding error: {e}")
            # 備用方案
            return json.dumps(data, ensure_ascii=True, separators=(',', ':'))

    def _calculate_node_positions(self):
        """
        根據節點的層級計算初始座標
        """
        level_counts = {}
        for node in self.nodes_data:
            level = node.get('level', 0)
            if level not in level_counts:
                level_counts[level] = 0
            level_counts[level] += 1

        level_y_step = 250
        level_x_step = 400

        current_level_counts = {level: 0 for level in level_counts}

        for node in self.nodes_data:
            level = node.get('level', 0)
            total_in_level = level_counts.get(level, 1)
            current_index_in_level = current_level_counts.get(level, 0)
            
            y = level * level_y_step
            x = (current_index_in_level - (total_in_level - 1) / 2.0) * level_x_step
            
            node['x'] = x
            node['y'] = y
            current_level_counts[level] = current_level_counts.get(level, 0) + 1