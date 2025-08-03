### **行動守則 (Action Guidelines)**

#### **A. 使用者指定的核心原則 (User-Defined Core Principles)**

1.  **進度報告 (Progress Reporting)**: 在每個階段完成後，更新 `PROGRESS_LOG.md`，記錄已完成的步驟、目前狀態及下一步計畫，確保進度清晰可追蹤。
2.  **小步進行，逐步驗證 (Small Steps, Incremental Verification)**: 所有重構任務將被拆解成最小的可驗證單元。我會先提出小範圍的修改，經您確認後才執行。
3.  **專注於程式碼遷移，而非修改 (Focus on Moving, Not Modifying)**: 主要任務是「搬移」程式碼到新結構中，除非絕對必要，否則不修改現有程式碼的內部邏輯。
4.  **檢查並修復依賴性 (Check and Fix Dependencies)**: 每次遷移後，必須立即檢查並修復因此產生的模組導入 (`import`) 錯誤，確保程式在結構改變後功能依然完整。

#### **B. 為確保順利進行的額外核心原則 (Additional Principles for a Smooth Process)**

1.  **保持系統可運作 (Keep the System Working)**: 在每完成一個微小的重構步驟後，整個應用程式理論上都應處於可運作的狀態。
2.  **測試驅動重構 (Test-Driven Refactoring - TDR)**: 在關鍵的重構點後進行手動或腳本化驗證，確保核心功能未被破壞。
3.  **手動版本備份 (Manual Version Backups)**: 在每完成一個成功的小步驟後，建議進行一次手動的資料夾備份。
4.  **提供詳細的測試指引 (Provide Detailed Testing Guidance)**: 在我完成每一次程式碼修改後，都必須提供一份清晰、詳細的指引，說明需要重點測試的功能區域以及具體的驗證步驟。
5.  **利用靜態分析工具 (Leverage Static Analysis Tools)**: 定期運行靜態分析工具 (如 Linter) 以發現潛在問題。
6.  **日誌驅動工作流程 (Log-Driven Workflow)**: 在執行任何操作**前**，先將「下一步行動提議」及「理由」更新至日誌末尾。

#### **C. 業界專業重構準則 (Industry-Standard Refactoring Best Practices)**

1.  **不要混合重構與功能開發 (Don't Mix Refactoring and Feature Addition)**: 一次只做一件事。
2.  **先理解，後重構 (Understand First, Refactor Second)**: 不重構自己不理解的程式碼。
3.  **三次法則 (The Rule of Three)**: 第三次出現重複程式碼時，才是重構的最佳時機。
4.  **有明確的重構目標 (Have a Clear Refactoring Goal)**: 我們的目標是 `REFACTORING_GUIDE.md` 中定義的 MVC 架構。
5.  **嚴格區分移動與修改 (Strict Distinction Between Moving vs. Modifying)**: 任何「修改」都必須被明確聲明並獲得批准。
