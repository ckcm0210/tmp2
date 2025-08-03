#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INDIRECT 函數解析器
處理 Excel 中的 INDIRECT 函數，將其解析為實際的儲存格引用

支援三種情況：
1. 參數指向其他儲存格 - =INDIRECT(D29&"!"&D26)
2. 硬編碼文字/數字混合 - =INDIRECT("Sheet"&B5&"!A1")  
3. 使用函數計算參數 - =INDIRECT("B"&ROW())
"""

import re
import os
from utils.openpyxl_resolver import read_cell_with_resolved_references

class IndirectResolver:
    """INDIRECT 函數解析器"""
    
    def __init__(self):
        self.cache = {}  # 緩存解析結果
        
    def is_indirect_formula(self, formula):
        """檢查公式是否包含 INDIRECT 函數"""
        if not formula or not isinstance(formula, str):
            return False
        return 'INDIRECT' in formula.upper()
    
    def extract_indirect_functions(self, formula):
        """從公式中提取所有 INDIRECT 函數"""
        if not self.is_indirect_formula(formula):
            return []
        
        # 正則表達式匹配 INDIRECT 函數
        pattern = r'INDIRECT\s*\(\s*([^)]+)\s*\)'
        matches = re.finditer(pattern, formula, re.IGNORECASE)
        
        indirect_functions = []
        for match in matches:
            indirect_functions.append({
                'full_match': match.group(0),
                'parameters': match.group(1),
                'start_pos': match.start(),
                'end_pos': match.end()
            })
        
        return indirect_functions
    
    def resolve_indirect_formula(self, formula, workbook_path, sheet_name):
        """
        解析 INDIRECT 公式，返回解析結果
        
        Args:
            formula: 包含 INDIRECT 的公式
            workbook_path: 當前工作簿路徑
            sheet_name: 當前工作表名稱
            
        Returns:
            dict: 解析結果
        """
        if not self.is_indirect_formula(formula):
            return {
                'is_indirect': False,
                'original_formula': formula,
                'resolved_reference': None,
                'resolution_status': 'not_indirect'
            }
        
        try:
            # 提取 INDIRECT 函數
            indirect_functions = self.extract_indirect_functions(formula)
            
            if not indirect_functions:
                return {
                    'is_indirect': True,
                    'original_formula': formula,
                    'resolved_reference': None,
                    'resolution_status': 'no_indirect_found'
                }
            
            # 處理第一個 INDIRECT 函數（簡化版本）
            first_indirect = indirect_functions[0]
            parameters = first_indirect['parameters']
            
            # 解析參數
            resolved_ref = self._resolve_indirect_parameters(
                parameters, workbook_path, sheet_name
            )
            
            if resolved_ref:
                return {
                    'is_indirect': True,
                    'original_formula': formula,
                    'original_indirect_formula': first_indirect['full_match'],
                    'resolved_reference': resolved_ref,
                    'indirect_parameters': parameters,
                    'resolution_status': 'success'
                }
            else:
                return {
                    'is_indirect': True,
                    'original_formula': formula,
                    'original_indirect_formula': first_indirect['full_match'],
                    'resolved_reference': None,
                    'indirect_parameters': parameters,
                    'resolution_status': 'resolution_failed'
                }
                
        except Exception as e:
            return {
                'is_indirect': True,
                'original_formula': formula,
                'resolved_reference': None,
                'resolution_status': 'error',
                'error_message': str(e)
            }
    
    def _resolve_indirect_parameters(self, parameters, workbook_path, sheet_name):
        """
        解析 INDIRECT 函數的參數
        
        處理三種情況：
        1. 儲存格引用：D29&"!"&D26
        2. 硬編碼混合："Sheet"&B5&"!A1"
        3. 函數計算：ROW(), COLUMN() 等
        """
        try:
            # 第一種情況：檢查是否包含儲存格引用
            if self._contains_cell_references(parameters):
                return self._resolve_cell_reference_parameters(
                    parameters, workbook_path, sheet_name
                )
            
            # 第二種情況：硬編碼文字混合
            if '"' in parameters:
                return self._resolve_hardcoded_parameters(
                    parameters, workbook_path, sheet_name
                )
            
            # 第三種情況：函數計算（暫時返回 None，後續實現）
            if any(func in parameters.upper() for func in ['ROW()', 'COLUMN()', 'CHAR(', 'CONCATENATE(']):
                return self._resolve_function_parameters(
                    parameters, workbook_path, sheet_name
                )
            
            # 如果都不匹配，嘗試直接作為引用
            return parameters.strip('"')
            
        except Exception as e:
            print(f"Error resolving INDIRECT parameters: {e}")
            return None
    
    def _contains_cell_references(self, parameters):
        """檢查參數是否包含儲存格引用"""
        # 簡單的儲存格引用模式：A1, B2, $A$1 等
        cell_pattern = r'\b[A-Z]+\d+\b'
        return bool(re.search(cell_pattern, parameters))
    
    def _resolve_cell_reference_parameters(self, parameters, workbook_path, sheet_name):
        """
        解析包含儲存格引用的參數
        例如：D29&"!"&D26
        """
        try:
            # 找到所有儲存格引用
            cell_pattern = r'\b([A-Z]+\d+)\b'
            cell_refs = re.findall(cell_pattern, parameters)
            
            # 讀取每個儲存格的值
            cell_values = {}
            for cell_ref in cell_refs:
                try:
                    cell_info = read_cell_with_resolved_references(
                        workbook_path, sheet_name, cell_ref
                    )
                    cell_values[cell_ref] = str(cell_info.get('display_value', ''))
                except Exception as e:
                    print(f"Error reading cell {cell_ref}: {e}")
                    cell_values[cell_ref] = ''
            
            # 替換參數中的儲存格引用為實際值
            resolved_params = parameters
            for cell_ref, value in cell_values.items():
                resolved_params = resolved_params.replace(cell_ref, f'"{value}"')
            
            # 評估字串連接表達式
            return self._evaluate_string_concatenation(resolved_params)
            
        except Exception as e:
            print(f"Error resolving cell reference parameters: {e}")
            return None
    
    def _resolve_hardcoded_parameters(self, parameters, workbook_path, sheet_name):
        """
        解析硬編碼文字混合的參數
        例如："Sheet"&B5&"!A1"
        """
        try:
            # 找到儲存格引用並替換為值
            cell_pattern = r'\b([A-Z]+\d+)\b'
            
            def replace_cell_ref(match):
                cell_ref = match.group(1)
                try:
                    cell_info = read_cell_with_resolved_references(
                        workbook_path, sheet_name, cell_ref
                    )
                    return f'"{cell_info.get("display_value", "")}"'
                except:
                    return '""'
            
            resolved_params = re.sub(cell_pattern, replace_cell_ref, parameters)
            
            # 評估字串連接表達式
            return self._evaluate_string_concatenation(resolved_params)
            
        except Exception as e:
            print(f"Error resolving hardcoded parameters: {e}")
            return None
    
    def _resolve_function_parameters(self, parameters, workbook_path, sheet_name):
        """
        解析包含函數的參數（第三種情況）
        暫時返回 None，標記為待實現
        """
        # TODO: 實現函數計算邏輯
        print(f"Function parameters not yet implemented: {parameters}")
        return None
    
    def _evaluate_string_concatenation(self, expression):
        """
        評估字串連接表達式
        例如："Sheet"&"2"&"!A1" -> "Sheet2!A1"
        """
        try:
            # 分割 & 運算符
            parts = expression.split('&')
            result = ""
            
            for part in parts:
                part = part.strip()
                # 移除引號
                if part.startswith('"') and part.endswith('"'):
                    result += part[1:-1]
                else:
                    result += part
            
            return result
            
        except Exception as e:
            print(f"Error evaluating string concatenation: {e}")
            return None