# PM Email v3 — Audit Findings Report
## BMAD Red-Team Session — 2026-06-24

**Audited file:** `D:\100.Software\Github\OutlookOkan\.claude\commands\pm-email.md`
**Audit team:** Murat (Test Arch), Mary (Analyst), John (PM), Sally (UX)
**Method:** Đọc trực tiếp file thực tế, đối chiếu từng dòng với báo cáo

---

## VERDICT: Báo cáo trước KHÔNG phản ánh đúng thực trạng file

**13 vấn đề được xác nhận** — 5 Critical, 4 High, 4 Medium.
File thiếu hụt nghiêm trọng so với những gì được báo cáo là "đã hoàn thành".

---

## DANH SÁCH LỖI ĐẦY ĐỦ

### 🔴 CRITICAL — Phải fix trước khi dùng được

| ID | Mô tả | Dòng file | Agent phát hiện |
|----|-------|-----------|-----------------|
| C-01 | **5/8 template bodies không tồn tại** — [3] Kickoff, [4] Approval Request, [5] Meeting Recap, [7] Risk Notification, [8] Project Closure chỉ có tên trong menu, không có nội dung | 341-428 | Mary |
| C-02 | **SLA Monitor tìm sai folder** — `folder_tương_ứng` undefined; reply từ khách luôn vào Inbox, không phải project folder | 211-212 | Murat |
| C-03 | **`risk_keyword` undefined trong Workflow [R]** — biến được dùng nhưng không được gán từ user input | 685 | Murat |
| C-04 | **DASL không hỗ trợ pipe `|` làm OR operator trong Workflow [P]** — search `"risk|issue|delay|..."` sẽ tìm nguyên chuỗi, không tìm được gì | 623 | Murat |
| C-05 | **`project_folder` undefined trong Workflow [T]** — tool chain gọi `search_emails(folder=project_folder)` nhưng không có bước hỏi user về project/folder | 359 | Mary |

---

### 🟡 HIGH — Ảnh hưởng trải nghiệm người dùng

| ID | Mô tả | Dòng file | Agent phát hiện |
|----|-------|-----------|-----------------|
| H-01 | **Batch commands trong Quick Triage [8] có hiển thị nhưng không có implementation** — syntax `"flag 1,3,5"`, `"move 2,4 sang NCB"` được show nhưng không có instruction xử lý | 530-531 | Sally |
| H-02 | **Dashboard [4]: tối đa ~20 COM calls** — snapshot × 2 (30d + 7d) × 5 folders + contact_stats × 10 = ~20 calls, mỗi call 2-5s = có thể 100 giây chờ | 257-263 | Murat |
| H-03 | **Intent "status" mơ hồ** — `status` maps cả [4] Dashboard lẫn [P] PMI Status Report; user type "status" không rõ muốn workflow nào | 80, 86 | Murat |
| H-04 | **Session state path relative** — `.claude/pm-email-state.md` không xác định working directory; có thể save nhầm chỗ. Cần absolute path | 768 | Murat |
| H-05 | **`/pm-email 0` mâu thuẫn** — args "0" → hiển thị menu (dòng 26), nhưng trong menu [0] = EXIT; hành vi không nhất quán | 26, 65 | John |

---

### 🟢 MEDIUM — Cần cải thiện

| ID | Mô tả | Dòng file | Agent phát hiện |
|----|-------|-----------|-----------------|
| M-01 | **Deep Search "refine" syntax không có implementation** — output hướng dẫn gõ `"sender [tên]"` để refine nhưng không có instruction xử lý | 457-458 | Sally |
| M-02 | **Priority score C-level không có data source** — email_stats/snapshot không trả về chức danh; scoring heuristic cần được ghi rõ là "phỏng đoán từ pattern" không phải chính xác | 111-117 | John |
| M-03 | **Workflow [7] Cross-project tool chain không được viết** — Option B chỉ nói "chạy như Dashboard" không có chi tiết cụ thể | 477-478 | Mary |
| M-04 | **Không có progress indicator trong workflows** — user không biết đang ở bước nào nếu có lỗi giữa chừng | Toàn bộ | Sally |

---

## PHÂN TÍCH NGUYÊN NHÂN GỐC RỄ

### Tại sao báo cáo sai lệch?

1. **Gap spec vs. implementation:** File được viết từ spec nhưng một số phần chỉ copy menu items mà không viết nội dung (templates [3][4][5][7][8]).

2. **Undefined variables:** Các bước tool chain dùng biến như `project_folder`, `folder_tương_ứng`, `risk_keyword` nhưng không có bước assign từ user input.

3. **Aspirational features:** Batch processing, refine search, priority scoring C-level — được describe trong output format nhưng không có instruction implementation.

4. **API assumptions:** Pipe `|` trong DASL là assumption sai về API behavior của MCP tools.

---

## SCOPE CỦA SESSION SỬA TIẾP THEO

### Phải có để command hoạt động được (Minimum Viable Fix):

**MV-1:** Thêm 5 template bodies còn thiếu ([3][4][5][7][8]) — ước tính ~100 dòng

**MV-2:** Fix SLA folder logic — thay `folder_tương_ứng` → search Inbox sau đó các project folders

**MV-3:** Thêm bước hỏi user `project_folder` ở đầu Workflow [T]

**MV-4:** Fix `risk_keyword` trong [R] — gán từ user input

**MV-5:** Fix DASL search trong [P] — tách thành 2-3 search riêng

### Nên có (High value, medium effort):

**SH-1:** Fix H-01 Batch commands — thêm instruction xử lý

**SH-2:** Fix H-02 Dashboard COM calls — bỏ snapshot 7d riêng, dùng data từ snapshot 30d

**SH-3:** Fix H-03 Intent "status" — disambiguation logic

**SH-4:** Fix H-04 Session state path — dùng absolute path

**SH-5:** Fix H-05 args "0" logic — phân biệt rõ args "0" vs. menu [0]

### Cải thiện (Medium value, low-medium effort):

**CI-1:** M-01 Refine search implementation
**CI-2:** M-02 Priority score caveat
**CI-3:** M-03 Cross-project tool chain
**CI-4:** M-04 Progress indicator pattern

---

## NEXT SESSION BRIEF (để team đọc trước khi bắt đầu)

### Nhiệm vụ:
Sửa toàn bộ 13 issues trong pm-email.md, ưu tiên theo MV → SH → CI.

### File cần sửa:
- `D:\100.Software\Github\OutlookOkan\.claude\commands\pm-email.md` — file chính (810 dòng)
- `D:\100.Software\Github\OutlookOkan\_bmad\bmm\workflows\pm-email\workflow.md` — cập nhật sau

### Ràng buộc quan trọng:
- KHÔNG thêm .Send() — chỉ .Display()
- KHÔNG auto-commit git
- MY_EMAIL = "thanhnt@softmart.net.vn" từ config.toml
- Tất cả COM operations qua STA thread
- Session state path: `D:\100.Software\Github\OutlookOkan\.claude\pm-email-state.md`

### Template bodies còn thiếu (phải viết mới hoàn toàn):

**[3] Kickoff Email:**
- Subject: [Project] — Project Kickoff — [Ngày]
- Mục tiêu: Giới thiệu PM, team, timeline, ground rules, first steps
- Sections: Team intro / Project scope / Timeline milestones / Communication plan / Next actions

**[4] Approval Request:**
- Subject: [APPROVAL NEEDED] [Project] — [Tài liệu/Quyết định] — Deadline [Ngày]
- Mục tiêu: Yêu cầu phê duyệt cụ thể, rõ deadline, rõ hậu quả nếu không approve
- Sections: What you need approved / Why / Options / Deadline / Impact if delayed

**[5] Meeting Recap:**
- Subject: [Project] — Meeting Recap — [DD/MM] — Action Items
- Mục tiêu: Tóm tắt sau họp, confirm decisions, assign actions có deadline
- Sections: Attendees / Decisions made / Action items (who/what/when) / Next meeting

**[7] Risk Notification:**
- Subject: [Project] — Risk Alert — [Risk Name] — [Level: High/Med/Low]
- Mục tiêu: Thông báo risk mới phát hiện, đơn giản hơn [R] workflow
- Sections: Risk description / Current status / Impact / Proposed mitigation / Ask

**[8] Project Closure:**
- Subject: [Project] — Project Closure Notice — [Giai đoạn]
- Mục tiêu: Đánh dấu kết thúc phase/project, lessons learned, thank you
- Sections: Summary / Achievements / Lessons / Next steps (if any) / Thank you

---

## STATUS TRACKING

```
MV-1 Templates [3][4][5][7][8]:  ⬜ PENDING
MV-2 SLA folder fix:              ⬜ PENDING
MV-3 [T] project_folder ask:     ⬜ PENDING
MV-4 [R] risk_keyword assign:    ⬜ PENDING
MV-5 [P] DASL pipe fix:          ⬜ PENDING
SH-1 Batch commands:             ⬜ PENDING
SH-2 Dashboard COM optimize:     ⬜ PENDING
SH-3 Intent "status" disambig:   ⬜ PENDING
SH-4 Session state absolute path:⬜ PENDING
SH-5 args "0" vs menu [0]:       ⬜ PENDING
CI-1 Refine search impl:         ⬜ PENDING
CI-2 Priority score caveat:      ⬜ PENDING
CI-3 Cross-project tool chain:   ⬜ PENDING
CI-4 Progress indicator:         ⬜ PENDING
```

---

*Audit conducted by BMAD Party Mode — Murat, Mary, John, Sally*
*Document created: 2026-06-24*
*Target fix session: Next session*
