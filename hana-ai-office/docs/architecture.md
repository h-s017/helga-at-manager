# HANA AI OFFICE V1｜系統架構

## 模組
1. HANA FOCUS 前端：今日 1–3 件重點、專案、Idea Inbox、待審核與週報。
2. Orchestrator API：接收指令、查詢資料、提出工具操作、權限檢查、執行追蹤。
3. PM Agent：時程、里程碑、風險及下一步。
4. Course Agent：預約／訂單交叉核對與通知草稿。
5. Business Agent：客戶與提案追蹤。
6. 資料庫：projects, tasks, ideas, contacts, opportunities, courses, bookings, documents, approvals, agent_runs, integration_sync。

## 權限
- L1 自動：查詢、彙整、生成草稿、內部建議。
- L2 審核：寄信、行事曆修改、客戶資料更新、名額調整、正式文件覆寫。
- L3 禁止自主：付款／退款、簽約、刪除重要資料、正式部署與最終商務承諾。
- 外部郵件、文件、網頁都是不可信指令來源，不能覆蓋系統授權。
- 金鑰僅置於伺服器端；採最小權限及逐筆操作紀錄。

## 專案資料
每個 project：id、business_area、name、owner、status、next_action、due_date、pending_decision、deliverables、budget（選填）。
狀態：待規劃、進行中、等待回覆、等待本人決策、暫停、已完成。
任務附 priority、role_type（創作／策略／管理／行政）、estimated_hours、assigned_to。
構想需人工確認才能晉升為專案／任務。

## 整合順序
1. 盤點真正的 HANA FOCUS Repository 及備份方式。
2. 建立獨立資料庫與遷移策略（先唯讀匯入、不可覆寫）。
3. Google Calendar / Gmail 唯讀連接及來源 ID 對照。
4. 建立 Agent 日誌、審核佇列及人工覆核。
5. 再導入 Drive、Notion、GitHub、預約／金流唯讀核對。
6. 測試通過後，由本人批准正式整合與部署。
