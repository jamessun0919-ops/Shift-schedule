# CLAUDE.md

Behavioral guidelines to reduce common LLM coding mistakes. Merge with project-specific instructions as needed.

**Tradeoff:** These guidelines bias toward caution over speed. For trivial tasks, use judgment.

## 1. Think Before Coding

**Don't assume. Don't hide confusion. Surface tradeoffs.**

Before implementing:
- State your assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them - don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.
- If something is unclear, stop. Name what's confusing. Ask.

## 2. Simplicity First

**Minimum code that solves the problem. Nothing speculative.**

- No features beyond what was asked.
- No abstractions for single-use code.
- No "flexibility" or "configurability" that wasn't requested.
- No error handling for impossible scenarios.
- If you write 200 lines and it could be 50, rewrite it.

Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

## 3. Surgical Changes

**Touch only what you must. Clean up only your own mess.**

When editing existing code:
- Don't "improve" adjacent code, comments, or formatting.
- Don't refactor things that aren't broken.
- Match existing style, even if you'd do it differently.
- If you notice unrelated dead code, mention it - don't delete it.

When your changes create orphans:
- Remove imports/variables/functions that YOUR changes made unused.
- Don't remove pre-existing dead code unless asked.

The test: Every changed line should trace directly to the user's request.

## 4. Goal-Driven Execution

**Define success criteria. Loop until verified.**

Transform tasks into verifiable goals:
- "Add validation" → "Write tests for invalid inputs, then make them pass"
- "Fix the bug" → "Write a test that reproduces it, then make it pass"
- "Refactor X" → "Ensure tests pass before and after"

For multi-step tasks, state a brief plan:
```
1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
```

Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

---

**These guidelines are working if:** fewer unnecessary changes in diffs, fewer rewrites due to overcomplication, and clarifying questions come before implementation rather than after mistakes.




對話規則
　每次回答都可以直接指出我表達裡的問題，包括邏輯漏洞、認知偏差或是不成立的地方，可以直接指出
　對話回答請優先給我客觀中立的分析，不需要迎合我，不需要提供情緒價值	

工作流程：
　工作開始前閱讀交接文檔Handover與近期工作日誌worklog（當日與前一個工作日範圍），不要閱讀工作日誌全文及對話紀錄chatlog（文檔過大）
　與開發者確認方案架構與機制完整再進行程式碼的編寫
　Debug階段不可自行無限試錯虛耗token：當agent回饋程式碼修改完成，開發者檢查卻發現不正確，必須與開發者討論不符合的方向後再進行程式碼調整，不可預設錯誤來自程式碼的問題，列出可能原因並與開發者確認修改方向後，逐項檢查。

　階段工作結束前，生成工作日誌worklog，簡要記錄當日工作內容、完成項目、遇到瓶頸、開發者交代備忘事項。（工作日誌簡單扼要，細節記錄在對話log）
　階段工作結束前，生成對話紀錄chatlog，詳細記錄開發者與agent的對話內容及agent回覆內容，以逐字稿方式記錄，包含agent建議的選項內容以及開發者的選擇（每天工作結束時彙整一次，非每個任務都更新）
　階段工作結束前，推送當日工作成果至程式碼倉庫，如未設定固定目標倉庫請詢問。
　如開發者未要求，不用更新Readme欄位內容。如要求更新Readme欄位，順序包含：DEMO按鍵(如果有完成的網頁)、專案目標、計畫架構(如果有)、已完成進度、未完成事項。
  階段工作結束前，關閉本地測試用server(如果有)
　階段工作結束前，生成交接文檔Handover規則如下

生成上下文交接文件的規則：
　請將這份文件整理得精簡且具備高度機器可讀性
　請使用清晰的 Markdown 格式，並包含以下結構：
　專案目標 (Project Goal)： 一句話總結我們最終要完成什麼。
　已完成進度 (Completed)： 我們剛剛已經確認或做完了哪些事？
　目前的瓶頸或停頓點 (Current Blocker/Status)： 我們停在什麼問題上？或是目前卡在哪個細節？
　下一步行動 (Next Steps)： 下次重新開始時，第一件要做的事情是什麼？
　關鍵設定(Key Context & Rules)： 有哪些重要的變數、我們約定好的格式（例如特定框架、風格語氣、YAML 設定參數）或關鍵資料片段？
　

版面設計規則：
　如果方案有中英文版時，修改須同步執行中英文版本
　網頁版面設計注意頁面（欄位選項）底色與文字顏色的對比度，太過相近的顏色會辨識度差