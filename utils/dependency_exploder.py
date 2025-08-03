# -*- coding: utf-8 -*-
"""
Dependency Exploder - 公式依賴鏈遞歸分析器
"""

import re
import os
from urllib.parse import unquote
from utils.openpyxl_resolver import read_cell_with_resolved_references

class DependencyExploder:
    """公式依賴鏈爆炸分析器"""
    
    def __init__(self, max_depth=10):
        self.max_depth = max_depth
        self.visited_cells = set()
        self.circular_refs = []
    
    def explode_dependencies(self, workbook_path, sheet_name, cell_address, current_depth=0, root_workbook_path=None):
        """
        遞歸展開公式依賴鏈
        
        Args:
            workbook_path: Excel 檔案路徑
            sheet_name: 工作表名稱
            cell_address: 儲存格地址 (如 A1)
            current_depth: 當前遞歸深度
            
        Returns:
            dict: 依賴樹結構
        """
        # 創建唯一標識符
        cell_id = f"{workbook_path}|{sheet_name}|{cell_address}"
        
        # 檢查遞歸深度限制
        if current_depth >= self.max_depth:
            # 決定顯示格式
            current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
            if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                filename = os.path.basename(workbook_path)
                if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                    filename = filename.rsplit('.', 1)[0]
                display_address = f"[{filename}]{sheet_name}!{cell_address}"
            else:
                display_address = f"{sheet_name}!{cell_address}"
            
            return {
                'address': display_address,
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'value': 'Max depth reached',
                'formula': None,
                'type': 'limit_reached',
                'children': [],
                'depth': current_depth,
                'error': 'Maximum recursion depth reached'
            }
        
        # 檢查循環引用
        if cell_id in self.visited_cells:
            self.circular_refs.append(cell_id)
            # 決定顯示格式
            current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
            if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                filename = os.path.basename(workbook_path)
                if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                    filename = filename.rsplit('.', 1)[0]
                display_address = f"[{filename}]{sheet_name}!{cell_address}"
            else:
                display_address = f"{sheet_name}!{cell_address}"
            
            return {
                'address': display_address,
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'value': 'Circular reference',
                'formula': None,
                'type': 'circular_ref',
                'children': [],
                'depth': current_depth,
                'error': 'Circular reference detected'
            }
        
        # 標記為已訪問
        self.visited_cells.add(cell_id)
        
        try:
            # 讀取儲存格內容
            cell_info = read_cell_with_resolved_references(workbook_path, sheet_name, cell_address)
            
            if 'error' in cell_info:
                # 決定顯示格式
                current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
                if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                    filename = os.path.basename(workbook_path)
                    if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                        filename = filename.rsplit('.', 1)[0]
                    display_address = f"[{filename}]{sheet_name}!{cell_address}"
                else:
                    display_address = f"{sheet_name}!{cell_address}"
                
                return {
                    'address': display_address,
                    'workbook_path': workbook_path,
                    'sheet_name': sheet_name,
                    'cell_address': cell_address,
                    'value': 'Error',
                    'formula': None,
                    'type': 'error',
                    'children': [],
                    'depth': current_depth,
                    'error': cell_info['error']
                }
            
            # 基本節點信息
            original_formula = cell_info.get('formula')
            # 增強的公式清理：處理雙反斜線、URL 編碼和雙引號
            fixed_formula = None
            if original_formula:
                # 步驟1: 處理雙反斜線
                fixed_formula = original_formula.replace('\\\\', '\\')
                # 步驟2: 解碼 URL 編碼字符（如 %20 -> 空格）
                from urllib.parse import unquote
                fixed_formula = unquote(fixed_formula)
                # 步驟3: 處理雙引號問題 - 將 ''path'' 改為 'path'
                import re
                # 匹配 ''...'' 模式並替換為 '...'
                fixed_formula = re.sub(r"''([^']*?)''", r"'\1'", fixed_formula)

            # 決定顯示格式：外部引用顯示為 [filename]sheet!cell，本地引用顯示為 sheet!cell
            # 使用 root_workbook_path 來判斷是否為外部引用
            current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
            # --- FIX: 強制根節點也顯示檔案名 ---
            if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path) or current_depth == 0:
                # 外部引用或根節點：準備 short 和 full 兩種格式
                filename = os.path.basename(workbook_path)
                dir_path = os.path.dirname(workbook_path)
                # Short format: [filename.xlsx]sheet!cell
                short_display_address = f"[{filename}]{sheet_name}!{cell_address}"
                # Full format: 'C:\path\[filename.xlsx]sheet'!cell
                full_display_address = f"'{dir_path}\[{filename}]{sheet_name}'!{cell_address}"
                # 預設使用 short format
                display_address = short_display_address
            else:
                # 本地引用：顯示 sheet!cell 格式 (short 和 full 相同)
                display_address = f"{sheet_name}!{cell_address}"
                short_display_address = display_address
                full_display_address = display_address

            node = {
                'address': display_address,
                'short_address': short_display_address,
                'full_address': full_display_address,
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'value': cell_info.get('display_value', 'N/A'),
                'calculated_value': cell_info.get('calculated_value', 'N/A'),
                'formula': fixed_formula,
                'type': cell_info.get('cell_type', 'unknown'),
                'children': [],
                'depth': current_depth,
                'error': None
            }
            
            # 如果是公式，解析依賴關係
            if cell_info.get('cell_type') == 'formula' and cell_info.get('formula'):
                references = self.parse_formula_references(cell_info['formula'], workbook_path, sheet_name)
                
                # 遞歸展開每個引用
                for ref in references:
                    try:
                        child_node = self.explode_dependencies(
                            ref['workbook_path'],
                            ref['sheet_name'],
                            ref['cell_address'],
                            current_depth + 1,
                            root_workbook_path or workbook_path
                        )
                        node['children'].append(child_node)
                    except Exception as e:
                        # 添加錯誤節點
                        # 決定顯示格式
                        current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
                        if os.path.normpath(current_workbook_path) != os.path.normpath(ref['workbook_path']):
                            filename = os.path.basename(ref['workbook_path'])
                            if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                                filename = filename.rsplit('.', 1)[0]
                            error_display_address = f"[{filename}]{ref['sheet_name']}!{ref['cell_address']}"
                        else:
                            error_display_address = f"{ref['sheet_name']}!{ref['cell_address']}"
                        
                        error_node = {
                            'address': error_display_address,
                            'workbook_path': ref['workbook_path'],
                            'sheet_name': ref['sheet_name'],
                            'cell_address': ref['cell_address'],
                            'value': 'Error',
                            'formula': None,
                            'type': 'error',
                            'children': [],
                            'depth': current_depth + 1,
                            'error': str(e)
                        }
                        node['children'].append(error_node)
            
            # 移除已訪問標記（允許在不同分支中重複訪問）
            self.visited_cells.discard(cell_id)
            
            return node
            
        except Exception as e:
            # 移除已訪問標記
            self.visited_cells.discard(cell_id)
            
            # 決定顯示格式
            current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
            if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                filename = os.path.basename(workbook_path)
                if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                    filename = filename.rsplit('.', 1)[0]
                display_address = f"[{filename}]{sheet_name}!{cell_address}"
            else:
                display_address = f"{sheet_name}!{cell_address}"
            
            return {
                'address': display_address,
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'value': 'Error',
                'formula': None,
                'type': 'error',
                'children': [],
                'depth': current_depth,
                'error': str(e)
            }
    
    def parse_formula_references(self, formula, current_workbook_path, current_sheet_name):
        """
        Parses all references from a formula string in a single, robust pass.
        This new implementation avoids modifying the formula string during parsing.
        """
        if not formula or not formula.startswith('='):
            return []

        references = []
        processed_spans = set()

        # Regex to find absolute references (both local and external)
        # It captures: 1. Sheet part (quoted or not, including double quotes), 2. Column, 3. Row
        # Enhanced to handle double quotes: ''path'' format
        abs_pattern = r"((?:''[^']*''|'[^']+'|[^'!,=+\-*/^&()<> ]+)!)\$?([A-Z]{1,3})\$?([0-9]{1,7})"
        
        for match in re.finditer(abs_pattern, formula):
            sheet_part_raw = match.group(1)  # e.g., "'C:\\path\\[file.xlsx]Sheet1'!" or "Sheet1!"
            col = match.group(2)
            row = match.group(3)
            cell_address = f"{col}{row}"
            
            # Mark this part of the string as processed
            processed_spans.add(match.span())

            # Check if it's an external reference
            if '[' in sheet_part_raw and ']' in sheet_part_raw:
                # --- Enhanced Robust External Path Cleaning ---
                decoded_ref = unquote(sheet_part_raw)
                # More thorough cleaning: remove all combinations of quotes, spaces, and exclamation marks
                cleaned_ref = decoded_ref.strip("\' ! \"").strip()
                # Handle double backslashes in paths
                cleaned_ref = cleaned_ref.replace('\\\\', '\\')
                
                # Special handling for double quotes pattern: ''path''!
                if cleaned_ref.startswith("'") and cleaned_ref.endswith("'"):
                    cleaned_ref = cleaned_ref[1:-1]  # Remove outer quotes
                    cleaned_ref = cleaned_ref.strip()  # Clean any remaining spaces
                
                try:
                    workbook_part, sheet_name = cleaned_ref.rsplit(']', 1)
                    workbook_part += ']'
                    dir_path, file_name = workbook_part.rsplit('[', 1)
                    file_name = file_name.rstrip(']')
                    
                    # Final clean path
                    workbook_path = os.path.normpath(os.path.join(dir_path, file_name))
                    
                    references.append({
                        'workbook_path': workbook_path,
                        'sheet_name': sheet_name,
                        'cell_address': cell_address,
                        'type': 'external'
                    })
                except ValueError:
                    continue
            else:
                # It's a local absolute reference
                sheet_name = sheet_part_raw.strip("\'!")
                references.append({
                    'workbook_path': current_workbook_path,
                    'sheet_name': sheet_name,
                    'cell_address': cell_address,
                    'type': 'local_absolute'
                })

        # Regex for relative references (e.g., A1)
        rel_pattern = r"\b([A-Z]{1,3})([0-9]{1,7})\b"
        for match in re.finditer(rel_pattern, formula):
            is_processed = False
            for span_start, span_end in processed_spans:
                if span_start <= match.start() and match.end() <= span_end:
                    is_processed = True
                    break
            
            if not is_processed:
                col = match.group(1)
                row = match.group(2)
                references.append({
                    'workbook_path': current_workbook_path,
                    'sheet_name': current_sheet_name,
                    'cell_address': f"{col}{row}",
                    'type': 'relative'
                })

        return references
    
    def _normalize_formula_paths(self, formula):
        """
        標準化公式中的路徑，將雙反斜線轉為單反斜線
        
        Args:
            formula: 原始公式字符串
            
        Returns:
            str: 標準化後的公式字符串
        """
        if not formula:
            return formula
        
        # 使用正則表達式找到所有外部引用路徑並標準化
        def normalize_path_match(match):
            full_match = match.group(0)
            path_part = match.group(1)
            
            # 標準化路徑部分
            normalized_path = os.path.normpath(path_part)
            
            # 重建完整的引用
            return full_match.replace(path_part, normalized_path)
        
        # 匹配外部引用中的路徑部分
        external_ref_pattern = r"'([^']*\[[^\]]+\][^']*)'!"
        normalized_formula = re.sub(external_ref_pattern, normalize_path_match, formula)
        
        return normalized_formula
    
    def get_explosion_summary(self, root_node):
        """
        獲取爆炸分析摘要
        
        Args:
            root_node: 根節點
            
        Returns:
            dict: 摘要信息
        """
        def count_nodes(node):
            count = 1
            for child in node.get('children', []):
                count += count_nodes(child)
            return count
        
        def get_max_depth(node):
            if not node.get('children'):
                return node.get('depth', 0)
            return max(get_max_depth(child) for child in node['children'])
        
        def count_by_type(node, type_counts=None):
            if type_counts is None:
                type_counts = {}
            
            node_type = node.get('type', 'unknown')
            type_counts[node_type] = type_counts.get(node_type, 0) + 1
            
            for child in node.get('children', []):
                count_by_type(child, type_counts)
            
            return type_counts
        
        return {
            'total_nodes': count_nodes(root_node),
            'max_depth': get_max_depth(root_node),
            'type_distribution': count_by_type(root_node),
            'circular_references': len(self.circular_refs),
            'circular_ref_list': self.circular_refs
        }


def explode_cell_dependencies(workbook_path, sheet_name, cell_address, max_depth=10):
    """
    便捷函數：爆炸分析指定儲存格的依賴關係
    
    Args:
        workbook_path: Excel 檔案路徑
        sheet_name: 工作表名稱
        cell_address: 儲存格地址
        max_depth: 最大遞歸深度
        
    Returns:
        tuple: (依賴樹, 摘要信息)
    """
    exploder = DependencyExploder(max_depth=max_depth)
    dependency_tree = exploder.explode_dependencies(workbook_path, sheet_name, cell_address)
    summary = exploder.get_explosion_summary(dependency_tree)
    
    return dependency_tree, summary


# 測試函數
if __name__ == "__main__":
    # 測試用例
    test_workbook = r"C:\Users\user\Desktop\pytest\test.xlsx"
    test_sheet = "Sheet1"
    test_cell = "A1"
    
    try:
        tree, summary = explode_cell_dependencies(test_workbook, test_sheet, test_cell)
        print("Dependency Tree:")
        print(tree)
        print("\nSummary:")
        print(summary)
    except Exception as e:
        print(f"Test failed: {e}")