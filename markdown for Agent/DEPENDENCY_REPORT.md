### 檔案依賴關係詳細報告

#### **根目錄 (`C:\Users\user\Excel_tools_develop\Excel_tools_develop_v70`)**

**1. `main.py`**
*   `import tkinter as tk` (第 1 行)
*   `from tkinter import ttk` (第 2 行)
*   `from core.formula_comparator import ExcelFormulaComparator` (第 3 行)
*   `from ui.workspace_view import Workspace` (第 4 行)
*   `from core.mode_manager import ModeManager, AppMode` (第 5 行)


**2. `excel_utils.py`**
*   `import re` (第 8 行)
*   `import os` (第 9 行)
*   `import openpyxl` (第 10 行)
*   `from openpyxl.utils import column_index_from_string, get_column_letter` (第 11 行)













---

#### **`core` 目錄**

**1. `core\formula_comparator.py`**
*   `import tkinter as tk` (第 1 行)
*   `from tkinter import ttk, filedialog, messagebox` (第 2 行)
*   `import openpyxl` (第 3 行)
*   `from openpyxl.styles import Font, PatternFill` (第 4 行)
*   `import os` (第 5 行)
*   `from datetime import datetime` (第 6 行)
*   `import webbrowser` (第 7 行)
*   `from utils.excel_io import calculate_similarity` (第 8 行)
*   `from core.worksheet_tree import apply_filter` (第 231 行)

**2. `core\data_processor.py`**
*   `import re` (第 3 行)
*   `from utils.range_optimizer import smart_range_display` (第 4 行)

**3. `core\dual_pane_controller.py`**
*   `import tkinter as tk` (第 8 行)
*   `from tkinter import ttk` (第 9 行)
*   `from ui.worksheet.controller import WorksheetController` (第 10 行)
*   `import win32com.client` (第 100 行)

**4. `core\excel_connector.py`**
*   `import os` (第 1 行)
*   `import win32com.client` (第 2 行)
*   `from tkinter import messagebox` (第 3 行)
*   `import win32gui` (第 4 行)
*   `import win32con` (第 5 行)

**5. `core\excel_scanner.py`**
*   `import os` (第 1 行)
*   `import time` (第 2 行)
*   `import win32com.client` (第 3 行)
*   `from tkinter import messagebox` (第 4 行)
*   `import psutil` (第 5 行)
*   `import win32gui` (第 6 行)
*   `import win32process` (第 7 行)
*   `import win32con` (第 8 行)
*   `from core.formula_classifier import classify_formula_type` (第 9 行)
*   `from core.worksheet_tree import apply_filter` (第 10 行)
*   `import traceback` (第 120 行)
*   `import traceback` (第 161 行)

**6. `core\formula_classifier.py`**
*   `import re` (第 1 行)
*   `from .link_analyzer import is_external_link_regex_match` (第 2 行)

**7. `core\link_analyzer.py`**
*   `import re` (第 8 行)
*   `import os` (第 9 行)

**8. `core\mode_manager.py`**
*   `import tkinter as tk` (第 8 行)
*   `from tkinter import ttk` (第 9 行)
*   `from enum import Enum` (第 10 行)

**9. `core\models.py`**
*   `from dataclasses import dataclass` (第 1 行)

**10. `core\graph_generator.py`**
*   `import os` (第 3 行)
*   `import webbrowser` (第 4 行)
*   `from pyvis.network import Network` (第 5 行)
*   `import json` (第 6 行)

**11. `core\worksheet_export.py`**
*   `import tkinter as tk` (第 1 行)
*   `from tkinter import filedialog, messagebox` (第 2 行)
*   `import openpyxl` (第 3 行)
*   `from openpyxl.styles import Font, Alignment, Border, Side` (第 4 行)
*   `from datetime import datetime` (第 5 行)
*   `import os` (第 6 行)

**12. `core\worksheet_refresh.py`**
*   `from core.excel_scanner import refresh_data` (第 1 行)

**13. `core\worksheet_summary.py`**
*   `from ui.summary_window import SummaryWindow` (第 1 行)

**14. `core\worksheet_tree.py`**
*   `import tkinter as tk` (第 1 行)
*   `from tkinter import ttk` (第 2 行)
*   `import re` (第 3 行)
*   `from utils.dependency_converter import convert_tree_to_graph_data` (第 4 行)
*   `from core.graph_generator import GraphGenerator` (第 5 行)
*   `from core.worksheet_tree import go_to_reference` (第 879 行)

---

#### **`ui` 目錄**

**1. `ui\summary_window.py`**
*   `import tkinter as tk` (第 1 行)
*   `from tkinter import ttk, messagebox, filedialog` (第 2 行)
*   `import os` (第 3 行)
*   `import re` (第 4 行)
*   `import collections` (第 5 行)
*   `from ui.visualizer import show_visual_chart` (第 6 行)
*   `from utils.excel_helpers import select_ranges_in_excel, replace_links_in_excel` (第 7 行)
*   `from utils.range_optimizer import smart_range_display` (第 8 行)

**2. `ui\visualizer.py`**
*   `import tkinter as tk` (第 1 行)
*   `from tkinter import ttk, messagebox, filedialog` (第 2 行)
*   `from tkinter import messagebox` (第 3 行)
*   `import matplotlib.pyplot as plt` (第 4 行)
*   `import matplotlib.patches as patches` (第 5 行)
*   `from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk` (第 6 行)
*   `import os` (第 7 行)
*   `from utils.range_optimizer import parse_cell_address, format_range` (第 8 行)

**3. `ui\workspace_view.py`**
*   `import tkinter as tk` (第 1 行)
*   `from tkinter import ttk, messagebox, filedialog` (第 2 行)
*   `import pythoncom` (第 3 行)
*   `import win32com.client` (第 4 行)
*   `import win32gui` (第 5 行)
*   `import win32con` (第 6 行)
*   `import time` (第 7 行)
*   `import openpyxl` (第 8 行)
*   `import os` (第 9 行)
*   `from datetime import datetime` (第 10 行)
*   `import threading` (第 11 行)

**4. `ui\worksheet_ui.py`**
*   `import tkinter as tk` (第 1 行)
*   `from tkinter import ttk` (第 2 行)
*   `from core.excel_connector import reconnect_to_excel` (第 14 行)
*   `from core.worksheet_export import export_formulas_to_excel, import_and_update_formulas` (第 15 行)
*   `from core.worksheet_summary import summarize_external_links` (第 16 行)
*   `from core.worksheet_tree import apply_filter, sort_column, on_select, on_double_click` (第 17 行)

**5. `ui\worksheet\view.py`**
*   `from ui.worksheet_ui import create_ui_widgets, bind_ui_commands, _set_placeholder, _on_focus_in, _on_mouse_click, _on_focus_out` (第 14 行)

**6. `ui\modes\inspect_mode.py`**
*   `from core.worksheet_tree import go_to_reference_enhanced` (第 2 行)

---

#### **`utils` 目錄**

**1. `utils\dependency_converter.py`**
*   `import os` (第 3 行)
*   `import re` (第 4 行)
*   `import re` (第 48 行)
*   `import colorsys` (第 70 行)
*   `import colorsys` (第 200 行)

**2. `utils\dependency_exploder.py`**
*   `import re` (第 6 行)
*   `import os` (第 7 行)
*   `from urllib.parse import unquote` (第 8 行)
*   `from utils.openpyxl_resolver import read_cell_with_resolved_references` (第 9 行)
*   `from urllib.parse import unquote` (第 80 行)
*   `import re` (第 83 行)

**3. `utils\excel_helpers.py`**
*   `import tkinter as tk` (第 1 行)
*   `from tkinter import messagebox, ttk` (第 2 行)
*   `import os` (第 3 行)
*   `import re` (第 4 行)
*   `import openpyxl` (第 5 行)
*   `from core.excel_connector import activate_excel_window` (第 6 行)

**4. `utils\excel_io.py`**
*   `import os` (第 8 行)
*   `import re` (第 9 行)
*   `import openpyxl` (第 10 行)
*   `import xlrd` (第 40 行)

**5. `utils\helpers.py`**
*   `import os` (第 3 行)
*   `import datetime` (第 4 行)

**6. `utils\openpyxl_resolver.py`**
*   `import openpyxl` (第 7 行)
*   `import os` (第 8 行)
*   `import re` (第 9 行)

**7. `utils\range_optimizer.py`**
*   `import re` (第 1 行)
*   `import collections` (第 2 行)
*   `from openpyxl.utils import column_index_from_string, get_column_letter` (第 3 行)

**8. `utils\workbook_resolver.py`**
*   `import os` (第 6 行)
*   `import openpyxl` (第 7 行)
*   `from openpyxl.utils import get_column_letter, column_index_from_string` (第 8 行)
*   `import re` (第 9 行)
