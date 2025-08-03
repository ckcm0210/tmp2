# -*- coding: utf-8 -*-
"""
Created on Wed Jun 25 09:11:26 2025

@author: kccheng
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os
import re
import win32com.client
import win32gui
import win32con
from functools import partial

# Import functions from their new locations
from core.link_analyzer import get_referenced_cell_values
from utils.excel_io import find_matching_sheet, read_external_cell_value
from utils.range_optimizer import parse_excel_address
from core.excel_connector import activate_excel_window, find_external_workbook_path
from openpyxl.utils import get_column_letter, column_index_from_string

def apply_filter(controller, event=None):
    controller.view.result_tree.delete(*controller.view.result_tree.get_children())
    controller.cell_addresses.clear()
    address_filter_str = controller.view.filter_entries['address'].get().strip()
    parsed_address_filters = []
    if address_filter_str and address_filter_str != controller.placeholder_text:
        address_tokens = [token.strip() for token in address_filter_str.split(',') if token.strip()]
        if address_tokens:
            try:
                for token in address_tokens:
                    parsed_address_filters.append(parse_excel_address(token))
            except Exception as e:
                messagebox.showerror("Invalid Excel Address", str(e))
                return
    other_filters = {
        'type': (controller.show_formula.get(), controller.show_local_link.get(), controller.show_external_link.get()),
        'formula': controller.view.filter_entries['formula'].get().lower(),
        'result': controller.view.filter_entries['result'].get().lower(),
        'display_value': controller.view.filter_entries['display_value'].get().lower()
    }
    filtered_formulas = []
    for formula_data in controller.all_formulas:
        if len(formula_data) < 5: continue
        formula_type, address, formula_content, result_val, display_val = formula_data
        type_map = {'formula': other_filters['type'][0], 'local link': other_filters['type'][1], 'external link': other_filters['type'][2]}
        if not type_map.get(formula_type, True): continue
        if other_filters['formula'] and other_filters['formula'] not in str(formula_content).lower(): continue
        if other_filters['result'] and other_filters['result'] not in str(result_val).lower(): continue
        if other_filters['display_value'] and other_filters['display_value'] not in str(display_val).lower(): continue
        if parsed_address_filters:
            addr_upper = address.replace("$", "").upper()
            current_cell_match = re.match(r"([A-Z]+)([0-9]+)", addr_upper)
            if not current_cell_match: continue
            cell_col_str, cell_row_str = current_cell_match.groups()
            cell_col_idx = column_index_from_string(cell_col_str)
            cell_row_idx = int(cell_row_str)
            is_match = False
            for f_type, f_val in parsed_address_filters:
                if f_type == 'cell' and addr_upper == f_val:
                    is_match = True; break
                elif f_type == 'row_range':
                    start_r, end_r = map(int, f_val.split(':'))
                    if start_r <= cell_row_idx <= end_r:
                        is_match = True; break
                elif f_type == 'col_range':
                    start_c, end_c = f_val.split(':')
                    if column_index_from_string(start_c) <= cell_col_idx <= column_index_from_string(end_c):
                        is_match = True; break
                elif f_type == 'range':
                    start_cell, end_cell = f_val.split(':')
                    sc_str, sr_str = re.match(r"([A-Z]+)([0-9]+)", start_cell).groups()
                    ec_str, er_str = re.match(r"([A-Z]+)([0-9]+)", end_cell).groups()
                    if (column_index_from_string(sc_str) <= cell_col_idx <= column_index_from_string(ec_str) and
                        int(sr_str) <= cell_row_idx <= int(er_str)):
                        is_match = True; break
            if not is_match: continue
        filtered_formulas.append(formula_data)
    if controller.current_sort_column:
        col_index = controller.view.tree_columns.index(controller.current_sort_column)
        sort_dir = controller.sort_directions[controller.current_sort_column]
        filtered_formulas.sort(key=lambda x: str(x[col_index]), reverse=(sort_dir == -1))
    count = len(filtered_formulas)
    controller.view.formula_list_label.config(text=f"Formula List ({count} records):")
    for i, data in enumerate(filtered_formulas):
        tag = "evenrow" if i % 2 == 0 else "oddrow"
        item_id = controller.view.result_tree.insert("", "end", values=data, tags=(tag,))
        address_index = controller.view.tree_columns.index("address")
        if address_index < len(data):
            controller.cell_addresses[item_id] = data[address_index]

def sort_column(controller, col_id):
    controller.current_sort_column = col_id
    controller.sort_directions[col_id] *= -1
    apply_filter(controller)
    for column in controller.view.tree_columns:
        original_text = controller.view.result_tree.heading(column, "text").split(' ')[0]
        controller.view.result_tree.heading(column, text=original_text, image='')
    current_direction = " \u2191" if controller.sort_directions[col_id] == 1 else " \u2193"
    current_text = controller.view.result_tree.heading(col_id, "text").split(' ')[0]
    controller.view.result_tree.heading(col_id, text=current_text + current_direction)

def go_to_reference(controller, workbook_path, sheet_name, cell_address):
    try:
        try:
            controller.xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            try:
                controller.xl = win32com.client.Dispatch("Excel.Application")
                controller.xl.Visible = True
            except Exception as e:
                messagebox.showerror("Excel Error", f"Could not start or connect to Excel.\nError: {e}")
                return

        target_workbook = None
        normalized_workbook_path = os.path.normpath(workbook_path) if workbook_path else None

        if normalized_workbook_path:
            for wb in controller.xl.Workbooks:
                if os.path.normpath(wb.FullName) == normalized_workbook_path:
                    target_workbook = wb
                    break
            
            if not target_workbook:
                if os.path.exists(normalized_workbook_path):
                    try:
                        # Save original settings to restore later
                        original_display_alerts = controller.xl.DisplayAlerts
                        original_update_links = getattr(controller.xl, 'AskToUpdateLinks', True)
                        
                        # Disable all alerts and update prompts to prevent dialog interruptions
                        controller.xl.DisplayAlerts = False
                        controller.xl.AskToUpdateLinks = False
                        
                        # Open workbook with parameters to avoid dialog boxes
                        target_workbook = controller.xl.Workbooks.Open(
                            Filename=normalized_workbook_path,
                            UpdateLinks=0,  # Don't update any links
                            ReadOnly=False,
                            Format=1,
                            Password="",
                            WriteResPassword="",
                            IgnoreReadOnlyRecommended=True,
                            Notify=False,
                            AddToMru=False
                        )
                        
                        # Restore original settings
                        controller.xl.DisplayAlerts = original_display_alerts
                        controller.xl.AskToUpdateLinks = original_update_links
                        
                    except Exception as e:
                        # Ensure settings are restored even if opening fails
                        try:
                            controller.xl.DisplayAlerts = original_display_alerts
                            controller.xl.AskToUpdateLinks = original_update_links
                        except:
                            pass
                        messagebox.showerror("Error Opening File", f"Could not open workbook:\n{normalized_workbook_path}\n\nError: {e}")
                        return
                else:
                    found_in_open_workbooks = False
                    filename = os.path.basename(normalized_workbook_path)
                    for wb in controller.xl.Workbooks:
                        if wb.Name.lower() == filename.lower():
                            target_workbook = wb
                            found_in_open_workbooks = True
                            break
                    
                    if not found_in_open_workbooks:
                        for wb in controller.xl.Workbooks:
                            if normalized_workbook_path.lower() in wb.Name.lower():
                                target_workbook = wb
                                found_in_open_workbooks = True
                                break
                    
                    if not found_in_open_workbooks:
                        messagebox.showerror("File Not Found", f"The workbook '{filename}' was not found in open workbooks and the path does not exist:\n{normalized_workbook_path}")
                        return
        else:
            target_workbook = controller.workbook

        if not target_workbook:
            messagebox.showerror("Error", "Could not access the target workbook.")
            return

        target_worksheet = None
        try:
            target_worksheet = target_workbook.Worksheets(sheet_name)
        except Exception:
            messagebox.showerror("Worksheet Not Found", f"Could not find worksheet '{sheet_name}' in workbook '{os.path.basename(target_workbook.FullName)}'.")
            return

        activate_excel_window(controller)
        target_workbook.Activate()
        target_worksheet.Activate()
        target_worksheet.Range(cell_address).Select()

    except Exception as e:
        messagebox.showerror("Navigation Error", f"Could not navigate to cell '{cell_address}'.\nError: {e}")

def go_to_reference_new_tab(controller, workbook_path, sheet_name, cell_address, reference_display):
    try:
        go_to_reference(controller, workbook_path, sheet_name, cell_address)
        
        try:
            if workbook_path:
                file_name = os.path.basename(workbook_path)
                if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
                    file_name = file_name[:-4]
            else:
                file_name = "Current"
            
            tab_name = f"{file_name}|{sheet_name}!{cell_address}"
            
            if len(tab_name) > 25:
                if len(file_name) > 10:
                    file_name = file_name[:7] + "..."
                if len(sheet_name) > 10:
                    sheet_name = sheet_name[:7] + "..."
                tab_name = f"{file_name}|{sheet_name}!{cell_address}"
                
                if len(tab_name) > 25:
                    tab_name = f"{file_name[:5]}...|{sheet_name[:5]}...!{cell_address}"
        except:
            tab_name = f"{reference_display}"
            if len(tab_name) > 20:
                tab_name = tab_name[:17] + "..."
        
        counter = 1
        original_tab_name = tab_name
        while tab_name in controller.tab_manager.detail_tabs:
            tab_name = f"{original_tab_name}({counter})"
            counter += 1
        
        new_detail_text = controller.tab_manager.create_detail_tab(tab_name)
        
        try:
            if not controller.xl:
                controller.xl = win32com.client.GetActiveObject("Excel.Application")
            
            target_workbook = None
            normalized_workbook_path = os.path.normpath(workbook_path) if workbook_path else None
            
            if normalized_workbook_path:
                for wb in controller.xl.Workbooks:
                    try:
                        if os.path.normpath(wb.FullName) == normalized_workbook_path:
                            target_workbook = wb
                            break
                    except Exception:
                        continue
                
                if not target_workbook:
                    filename = os.path.basename(normalized_workbook_path)
                    for wb in controller.xl.Workbooks:
                        try:
                            if wb.Name.lower() == filename.lower():
                                target_workbook = wb
                                break
                        except Exception:
                            continue
            
            if not target_workbook:
                target_workbook = controller.workbook
            
            target_worksheet = None
            try:
                target_worksheet = target_workbook.Worksheets(sheet_name)
            except Exception as ws_error:
                new_detail_text.insert('end', f"Error accessing worksheet '{sheet_name}': {ws_error}\n", "info_text")
                return
            
            target_cell = None
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    target_worksheet.Activate()
                    target_cell = target_worksheet.Range(cell_address)
                    break
                except Exception as cell_error:
                    if attempt == max_retries - 1:
                        new_detail_text.insert('end', f"Error accessing cell '{cell_address}' after {max_retries} attempts: {cell_error}\n", "info_text")
                        return
                    else:
                        import time
                        time.sleep(0.1)
            
            cell_formula = None
            cell_value = None
            cell_display_value = None
            
            for attempt in range(max_retries):
                try:
                    cell_formula = target_cell.Formula if hasattr(target_cell, 'Formula') and target_cell.Formula else target_cell.Value
                    cell_value = target_cell.Value
                    cell_display_value = target_cell.Text if hasattr(target_cell, 'Text') else str(target_cell.Value)
                    break
                except Exception as value_error:
                    if attempt == max_retries - 1:
                        new_detail_text.insert('end', f"Error reading cell values after {max_retries} attempts: {value_error}\n", "info_text")
                        return
                    else:
                        import time
                        time.sleep(0.1)
            
            if target_cell.Formula and target_cell.Formula.startswith('='):
                cell_type = "formula"
                if any(ref in target_cell.Formula for ref in ['[', ']']):
                    cell_type = "external link"
                elif '!' in target_cell.Formula:
                    cell_type = "local link"
            else:
                cell_type = "value"
            
            new_detail_text.insert('end', "Type: ", "label")
            new_detail_text.insert('end', f"{cell_type} / ", "value")
            new_detail_text.insert('end', "Cell Address: ", "label")
            new_detail_text.insert('end', f"{sheet_name}!{cell_address}\n", "value")
            new_detail_text.insert('end', "Workbook: ", "label")
            new_detail_text.insert('end', f"{os.path.basename(target_workbook.FullName)}\n", "value")
            new_detail_text.insert('end', "Calculated Result: ", "label")
            new_detail_text.insert('end', f"{cell_value} / ", "result_value")
            new_detail_text.insert('end', "Displayed Value: ", "label")
            new_detail_text.insert('end', f"{cell_display_value}\n\n", "value")
            
            if cell_type == "formula" or cell_type.endswith("link"):
                new_detail_text.insert('end', "Formula Content:\n", "label")
                new_detail_text.insert('end', f"{cell_formula}\n\n", "formula_content")
                
                if target_cell.Formula and target_cell.Formula.startswith('='):
                    try:
                        read_func = read_external_cell_value
                        referenced_values = get_referenced_cell_values(
                            cell_formula,
                            target_worksheet,
                            target_workbook.FullName,
                            read_func,
                            lambda name, obj: find_matching_sheet(controller.workbook, name)
                        )
                        
                        if referenced_values:
                            new_detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
                            for ref_addr, ref_val in referenced_values.items():
                                display_text = ref_addr
                                if '|' in ref_addr:
                                    _, display_text = ref_addr.split('|', 1)

                                new_detail_text.insert('end', f"  {display_text}: {ref_val}  ", "referenced_value")

                                workbook_path_new = None
                                sheet_name_new = None
                                cell_address_to_go_new = None
                                
                                try:
                                    if '|' in ref_addr:
                                        full_path, display_ref = ref_addr.split('|', 1)
                                        workbook_path_new = full_path
                                        
                                        if ']' in display_ref and '!' in display_ref:
                                            sheet_and_cell = display_ref.split(']', 1)[1]
                                            parts = sheet_and_cell.rsplit('!', 1)
                                            sheet_name_new = parts[0].strip("'")
                                            cell_address_to_go_new = parts[1]
                                    else:
                                        workbook_path_new = target_workbook.FullName
                                        if '!' in ref_addr:
                                            parts = ref_addr.rsplit('!', 1)
                                            sheet_name_new = parts[0]
                                            cell_address_to_go_new = parts[1]

                                    if workbook_path_new and sheet_name_new and cell_address_to_go_new:
                                        def build_handler_new(wp, sn, ca, ref_display):
                                            def handler():
                                                go_to_reference_new_tab(controller, wp, sn, ca, ref_display)
                                            return handler
                                        
                                        btn = tk.Button(new_detail_text, text="Go to Reference", font=("Arial", 7), cursor="hand2", command=build_handler_new(workbook_path_new, sheet_name_new, cell_address_to_go_new, display_text))
                                        new_detail_text.window_create('end', window=btn)

                                except Exception as e:
                                    print(f"INFO: Could not create navigation button for '{ref_addr}': {e}")

                                new_detail_text.insert('end', "\n")
                        else:
                            new_detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
                            new_detail_text.insert('end', "  No individual cell references found or accessible.\n", "info_text")
                    except Exception as ref_error:
                        new_detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
                        new_detail_text.insert('end', f"  Error retrieving referenced values: {ref_error}\n", "info_text")
            else:
                new_detail_text.insert('end', "Content:\n", "label")
                new_detail_text.insert('end', f"{cell_value}\n", "value")
                
        except Exception as e:
            new_detail_text.insert('end', f"Error retrieving cell details: {e}\n", "info_text")
            
    except Exception as e:
        messagebox.showerror("Tab Creation Error", f"Could not create new tab for reference.\nError: {e}")

def on_select(controller, event):
    selected_item = controller.view.result_tree.selection()
    
    try:
        main_tab_info = controller.tab_manager.detail_tabs.get("Main") or controller.tab_manager.detail_tabs.get("Tab_0")
        if not main_tab_info:
            current_detail_text = controller.tab_manager.get_current_detail_text()
        else:
            current_detail_text = main_tab_info["text_widget"]
            controller.tab_manager.detail_notebook.select(main_tab_info["frame"])
    except (AttributeError, KeyError):
        current_detail_text = controller.tab_manager.get_current_detail_text()

    if not selected_item:
        current_detail_text.delete(1.0, 'end')
        return
        
    item_id = selected_item[0]
    values = controller.view.result_tree.item(item_id, "values")
    
    if len(values) < 5:
        current_detail_text.delete(1.0, 'end')
        current_detail_text.insert(1.0, "Selected item has incomplete data.")
        return
        
    formula_type, cell_address, formula, result, display_value = values
    
    current_detail_text.delete(1.0, 'end')
    current_detail_text.insert('end', "Type: ", "label")
    current_detail_text.insert('end', f"{formula_type} / ", "value")
    current_detail_text.insert('end', "Cell Address: ", "label")
    current_detail_text.insert('end', f"{cell_address}\n", "value")
    current_detail_text.insert('end', "Calculated Result: ", "label")
    current_detail_text.insert('end', f"{result} / ", "result_value")
    current_detail_text.insert('end', "Displayed Value: ", "label")
    current_detail_text.insert('end', f"{display_value}\n\n", "value")
    current_detail_text.insert('end', "Formula Content:\n", "label")
    current_detail_text.insert('end', f"{formula}\n\n", "formula_content")
    
    if controller.xl and controller.worksheet:
        read_func = read_external_cell_value
        referenced_values = get_referenced_cell_values(
            formula,
            controller.worksheet,
            controller.workbook.FullName,
            read_func,
            lambda name, obj: find_matching_sheet(controller.workbook, name)
        )
        if referenced_values:
            current_detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
            for ref_addr, ref_val in referenced_values.items():
                display_text = ref_addr
                if '|' in ref_addr:
                    _, display_text = ref_addr.split('|', 1)

                current_detail_text.insert('end', f"  {display_text}: {ref_val}  ", "referenced_value")

                workbook_path = None
                sheet_name = None
                cell_address_to_go = None
                
                try:
                    if '|' in ref_addr:
                        full_path, display_ref = ref_addr.split('|', 1)
                        workbook_path = full_path
                        
                        if ']' in display_ref and '!' in display_ref:
                            sheet_and_cell = display_ref.split(']', 1)[1]
                            parts = sheet_and_cell.rsplit('!', 1)
                            sheet_name = parts[0].strip("'")
                            cell_address_to_go = parts[1]
                        else:
                            if display_ref.startswith('[') and ']' in display_ref:
                                bracket_end = display_ref.find(']')
                                file_name = display_ref[1:bracket_end]
                                remaining = display_ref[bracket_end+1:]
                                if '!' in remaining:
                                    sheet_name, cell_address_to_go = remaining.split('!', 1)
                                    workbook_path = find_external_workbook_path(controller, file_name)
                    else:
                        workbook_path = controller.workbook.FullName
                        if '!' in ref_addr:
                            parts = ref_addr.rsplit('!', 1)
                            sheet_name = parts[0]
                            cell_address_to_go = parts[1]

                    if workbook_path and sheet_name and cell_address_to_go:
                        def build_handler(wp, sn, ca, ref_display):
                            def handler():
                                go_to_reference_new_tab(controller, wp, sn, ca, ref_display)
                            return handler
                        
                        btn = tk.Button(current_detail_text, text="Go to Reference", font=("Arial", 7), cursor="hand2", command=build_handler(workbook_path, sheet_name, cell_address_to_go, display_text))
                        current_detail_text.window_create('end', window=btn)

                except Exception as e:
                    print(f"INFO: Could not create navigation button for '{ref_addr}': {e}")

                current_detail_text.insert('end', "\n")
        else:
            current_detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
            current_detail_text.insert('end', "  No individual cell references found or accessible.\n", "info_text")
    else:
        current_detail_text.insert('end', "Excel connection not active to retrieve referenced values.\n", "info_text")
        
def on_double_click(controller, event):
    selected_item = controller.view.result_tree.selection()
    if not selected_item:
        return
    item_id = selected_item[0]
    cell_address = controller.cell_addresses.get(item_id)
    if cell_address:
        try:
            try:
                controller.xl = win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                controller.xl = None
            if not controller.xl:
                try:
                    controller.xl = win32com.client.Dispatch("Excel.Application")
                    controller.xl.Visible = True
                    if controller.last_workbook_path and os.path.exists(controller.last_workbook_path):
                        controller.workbook = controller.xl.Workbooks.Open(controller.last_workbook_path)
                    else:
                        messagebox.showwarning("File Not Found", "The last scanned Excel file path is not valid or found. Please open Excel manually.")
                        return
                except Exception as e:
                    messagebox.showerror("Excel Launch Error", f"Could not launch Excel or open the workbook.\nError: {e}")
                    return
            target_workbook = None
            
            if controller.last_workbook_path:
                normalized_path = os.path.normpath(controller.last_workbook_path)
                for wb in controller.xl.Workbooks:
                    if os.path.normpath(wb.FullName) == normalized_path:
                        target_workbook = wb
                        break
                
                if not target_workbook and os.path.exists(controller.last_workbook_path):
                    try:
                        # Save original settings to restore later
                        original_display_alerts = controller.xl.DisplayAlerts
                        original_update_links = getattr(controller.xl, 'AskToUpdateLinks', True)
                        
                        # Disable all alerts and update prompts to prevent dialog interruptions
                        controller.xl.DisplayAlerts = False
                        controller.xl.AskToUpdateLinks = False
                        
                        # Open workbook with parameters to avoid dialog boxes
                        target_workbook = controller.xl.Workbooks.Open(
                            Filename=controller.last_workbook_path,
                            UpdateLinks=0,  # Don't update any links
                            ReadOnly=False,
                            Format=1,
                            Password="",
                            WriteResPassword="",
                            IgnoreReadOnlyRecommended=True,
                            Notify=False,
                            AddToMru=False
                        )
                        
                        # Restore original settings
                        controller.xl.DisplayAlerts = original_display_alerts
                        controller.xl.AskToUpdateLinks = original_update_links
                        
                    except Exception as open_e:
                        # Ensure settings are restored even if opening fails
                        try:
                            controller.xl.DisplayAlerts = original_display_alerts
                            controller.xl.AskToUpdateLinks = original_update_links
                        except:
                            pass
                        messagebox.showerror("Workbook Open Error", f"Could not open workbook '{os.path.basename(controller.last_workbook_path)}'.\nError: {open_e}")
                        return
                
                if not target_workbook:
                    filename = os.path.basename(controller.last_workbook_path)
                    for wb in controller.xl.Workbooks:
                        if wb.Name.lower() == filename.lower():
                            target_workbook = wb
                            break
            
            if not target_workbook:
                target_workbook = controller.xl.ActiveWorkbook if controller.xl.ActiveWorkbook else controller.workbook
            
            if not target_workbook:
                messagebox.showerror("Error", "No workbook available to navigate to.")
                return
            
            controller.workbook = target_workbook
            if controller.last_worksheet_name and controller.workbook:
                try:
                    controller.worksheet = controller.workbook.Worksheets(controller.last_worksheet_name)
                except Exception:
                    controller.worksheet = controller.workbook.ActiveSheet
                    messagebox.showwarning("Worksheet Not Found", f"Worksheet '{controller.last_worksheet_name}' not found in '{controller.workbook.Name}'. Activating current sheet.")
            elif controller.workbook:
                controller.worksheet = controller.workbook.ActiveSheet
            else:
                messagebox.showerror("Error", "No active workbook to select cell in.")
                return
            controller.workbook.Activate()
            controller.worksheet.Activate()
            controller.worksheet.Range(cell_address).Select()
            activate_excel_window(controller)
        except Exception as e:
            messagebox.showerror("Excel Selection Error", f"Could not select cell {cell_address} in Excel. Please ensure the workbook and worksheet are still valid.\nError: {e}")