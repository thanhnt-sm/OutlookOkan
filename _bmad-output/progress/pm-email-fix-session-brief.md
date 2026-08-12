# PM Email v3 Fix Session — Team Brief
## BMAD Next Session — Dành cho toàn team đọc trước khi bắt đầu

**Tài liệu này:** Brief cho session tiếp theo để brainstorm → plan → thực thi
**Audit full:** `_bmad-output/analysis/pm-email-v3-audit-findings.md`
**File cần sửa:** `.claude/commands/pm-email.md`

---

## BỐI CẢNH (TÓM TẮT CHO TEAM MỚI)

Chúng ta đang xây dựng một slash command `/pm-email` — công cụ giúp Project Manager làm việc với email Outlook thông qua Claude Code. Tool này dùng MCP server (`outlook-mcp-secure`) để gọi Outlook Desktop qua Windows COM.

**Phiên trước đã làm:** Xây dựng pm-email v3 với 13 workflows + menu nhóm + PMI framework.

**Phiên trước phát hiện ra:** Báo cáo không phản ánh đúng thực tế. File có 13 vấn đề nghiêm trọng, nhiều tính năng được liệt kê trong menu nhưng chưa có implementation.

---

## NHIỆM VỤ PHIÊN NÀY

### Phase 1 — Brainstorm (team debate)
Trước khi sửa, team cần thống nhất:
- Template bodies: viết theo style gì? (formal/professional/friendly?)
- Batch commands: implement bằng cách nào? (parse text, gọi tool loop?)
- Progress indicator: hiển thị dạng gì trong Claude output?
- SLA search strategy: chỉ search Inbox hay scan thêm project folders?

### Phase 2 — Plan (phân công)
Ai sửa gì? Theo thứ tự:
1. MV fixes (5 Critical) — phải xong trước
2. SH fixes (4 High) — sau khi MV xong
3. CI fixes (4 Medium) — cuối cùng nếu còn thời gian

### Phase 3 — Thực thi
Sửa file theo plan, verify từng fix.

---

## DANH SÁCH FIXES CỤ THỂ

### MV-1: 5 Template Bodies còn thiếu

**[3] Kickoff Email** — Email đầu tiên khi bắt đầu project/phase
```
Subject: [Project Name] — Project Kickoff — [DD/MM/YYYY]

Dear [Tên team/stakeholders],

Tôi xin thông báo dự án [tên] chính thức bắt đầu từ [ngày].

TEAM:
• PM: [Tên] — [email]
• [Role khác]: [Tên] — [email]

SCOPE TÓM TẮT:
[1-2 câu mô tả dự án là gì, deliver gì]

TIMELINE:
• Kick-off: [Ngày]
• [Milestone 1]: [Ngày]
• [Milestone 2]: [Ngày]
• Go-live / Hoàn thành: [Ngày]

COMMUNICATION PLAN:
• Họp định kỳ: [Ngày trong tuần, giờ]
• Báo cáo tuần: [Ngày gửi]
• Kênh liên lạc chính: [Email / Teams / Zalo]

NEXT ACTIONS:
• [Ai]: [Làm gì] — trước [Ngày]
• ...

Mọi câu hỏi xin liên hệ tôi trực tiếp.

Trân trọng,
[Tên PM]
```

---

**[4] Approval Request** — Yêu cầu phê duyệt tài liệu/quyết định
```
Subject: [APPROVAL NEEDED] [Project] — [Tên tài liệu/quyết định] — Deadline [DD/MM]

[Tên người có thẩm quyền],

Tôi cần sự phê duyệt của anh/chị cho nội dung sau:

CẦN PHÊ DUYỆT:
[Mô tả rõ: tài liệu gì / quyết định gì]

LÝ DO CẦN PHÊ DUYỆT:
[1-2 câu giải thích bối cảnh]

CÁC LỰA CHỌN (nếu có):
Option A: [Mô tả] — Ưu: ... | Nhược: ...
Option B: [Mô tả] — Ưu: ... | Nhược: ...
Tôi đề xuất: [Option X] vì [lý do ngắn]

DEADLINE: [DD/MM/YYYY]
Lý do: [Nếu không approve trước ngày này, sẽ ảnh hưởng đến...]

TÀI LIỆU ĐÍNH KÈM: [Tên file nếu có]

Anh/chị vui lòng reply "Đồng ý" / "Cần thảo luận" / "Từ chối + lý do".

Trân trọng,
[Tên PM]
```

---

**[5] Meeting Recap** — Tóm tắt sau họp
```
Subject: [Project] — Meeting Recap [DD/MM] — [X] Action Items

Xin chào,

Tóm tắt cuộc họp [loại họp] ngày [DD/MM/YYYY], [HH:mm]-[HH:mm].

NGƯỜI THAM DỰ:
• [Tên] ([Công ty/Role])
• ...

QUYẾT ĐỊNH ĐÃ CHỐT:
✅ [Quyết định 1]
✅ [Quyết định 2]

ACTION ITEMS:
┌─────────────────────────────┬──────────┬────────────┐
│ Việc cần làm                │ Ai       │ Deadline   │
├─────────────────────────────┼──────────┼────────────┤
│ [Mô tả việc 1]              │ [Tên]    │ [DD/MM]    │
│ [Mô tả việc 2]              │ [Tên]    │ [DD/MM]    │
└─────────────────────────────┴──────────┴────────────┘

VẤN ĐỀ CÒN MỞ:
❓ [Câu hỏi / vấn đề chưa giải quyết] — Chờ input từ [Ai]

CUỘC HỌP TIẾP THEO: [Ngày / Giờ / Link]

Vui lòng xác nhận action items của mình trong vòng 24h.

Trân trọng,
[Tên PM]
```

---

**[7] Risk Notification** — Thông báo rủi ro (ngắn hơn Workflow [R])
```
Subject: [Project] — ⚠️ Risk Alert: [Tên rủi ro] — [High/Medium/Low]

[Tên stakeholders],

Tôi muốn thông báo một rủi ro mới vừa được nhận diện trong dự án.

RỦI RO: [Mô tả ngắn gọn 1-2 câu]
MỨC ĐỘ: [🔴 High / 🟡 Medium / 🟢 Low]
KHẢ NĂNG XẢY RA: [Cao / Trung bình / Thấp]

TÁC ĐỘNG NẾU XẢY RA:
• Lịch trình: [ảnh hưởng thế nào]
• Chi phí: [ảnh hưởng thế nào nếu có]

HÀNH ĐỘNG ĐÃ TRIỂN KHAI:
• [Hành động 1 — ai chịu trách nhiệm]
• [Hành động 2]

CẦN HỖ TRỢ TỪ ANH/CHỊ: [Cụ thể hoặc "Không — chỉ thông báo để nắm thông tin"]

Tôi sẽ cập nhật tình trạng vào [ngày check-in tiếp theo].

Trân trọng,
[Tên PM]
```

---

**[8] Project Closure** — Email kết thúc project/phase
```
Subject: [Project] — Project Closure — [Phase / Toàn bộ dự án] hoàn thành

[Tên team/stakeholders],

Tôi xin thông báo [dự án / giai đoạn] [tên] đã chính thức hoàn thành vào [ngày].

THÀNH QUẢ ĐẠT ĐƯỢC:
• [Deliverable 1] — hoàn thành [ngày]
• [Deliverable 2] — hoàn thành [ngày]
• ...

SO VỚI KẾ HOẠCH BAN ĐẦU:
• Lịch trình: [Đúng hạn / Trễ X ngày — lý do]
• Ngân sách: [Trong budget / Vượt X% — lý do]

LESSONS LEARNED (Bài học):
• Làm tốt: [1-2 điểm]
• Cần cải thiện lần sau: [1-2 điểm]

BƯỚC TIẾP THEO: [Maintenance phase / Dự án mới / Không có]

Cảm ơn tất cả các bên đã phối hợp trong suốt thời gian vừa qua.
Tôi luôn sẵn sàng hỗ trợ nếu có câu hỏi về documentation.

Trân trọng,
[Tên PM]
```

---

### MV-2: Fix SLA Monitor folder

**Trước (sai):**
```
search_emails(keyword=subject_gốc, folder=folder_tương_ứng, max=5)
```

**Sau (đúng):**
```
BƯỚC B2 — Check reply trong Inbox trước, sau đó project folders:
search_emails(keyword=subject_gốc, folder="Inbox", max=5).
Nếu không tìm thấy: search_emails(keyword=subject_gốc, folder=PVC.CLIMS, max=3) và các project folders.
Nếu tất cả đều không có reply: đây là SLA breach.
Giới hạn: chỉ check tối đa 10 sent emails để tránh quá nhiều COM calls.
```

---

### MV-3: Fix Workflow [T] — thêm bước hỏi project/folder

**Thêm trước BƯỚC 1:**
```
BƯỚC 0 — Xác định context:
Nếu args chứa tên project/folder rõ ràng: dùng ngay.
Nếu không: hỏi "Template này cho project nào?" (tên folder trong allowed folders).
Lưu vào biến project_folder để dùng trong các bước sau.
```

---

### MV-4: Fix Workflow [R] — gán risk_keyword

**Thêm sau "Hỏi: Risk/issue gì?":**
```
Lấy từ user input: tên/mô tả risk = risk_description.
Tạo risk_keyword = 1-2 từ ngắn từ risk_description để search email.
(Ví dụ: "delay deployment" → risk_keyword = "deploy")
```

---

### MV-5: Fix DASL search trong Workflow [P]

**Trước (sai):**
```
search_emails(folder=folder, keyword="risk|issue|delay|problem|vấn đề|chậm|rủi ro", max=10)
```

**Sau (đúng):**
```
BƯỚC 3 — Tìm risk/issue emails (2 calls riêng):
search_emails(folder=folder, keyword="risk issue delay problem", max=5).
search_emails(folder=folder, keyword="vấn đề chậm rủi ro", max=5).
Gộp kết quả, loại trùng.
```

---

### SH-1: Implement Batch Commands trong Quick Triage [8]

**Thêm sau phần Output:**
```
XỬ LÝ BATCH COMMANDS:
Khi user nhập batch command, parse và thực hiện tuần tự:
- "flag 1,3,5" → gọi flag_email(entry_id) cho items số 1, 3, 5 trong list
- "move 2,4 sang [folder]" → gọi move_email(entry_id, destination=folder) cho items 2, 4
- "mark 6-10 read" → gọi mark_email_read(entry_id) cho items 6 đến 10
- "reply 1" → chạy Workflow 5 (Smart Draft) cho item số 1
Sau khi xong: "Đã xử lý [X] items. Còn [Y] items trong triage."
```

---

### SH-2: Tối ưu Dashboard COM calls

**Thay Bước 3 hiện tại (gọi snapshot 7d riêng) bằng:**
```
BƯỚC 3 (Stakeholder Temperature — không gọi thêm snapshot):
Dùng dữ liệu từ Bước 1 (snapshot 30d) đã có.
Từ danh sách top senders: gọi get_contact_stats(email=contact) cho top 2 contacts/project.
Phân tích trend: email count 30 ngày / 4 tuần = avg/tuần.
So sánh với contact_stats.recent_week_count nếu có trong response.
```

---

### SH-3: Disambiguate Intent "status"

**Thay dòng:**
```
- 4 / "dashboard" / "health" / "status" / "all"  → WORKFLOW 4
```
**Bằng:**
```
- 4 / "dashboard" / "health" / "all"             → WORKFLOW 4
- "status" đứng một mình (không có "project")    → Hỏi: "[4] Dashboard tổng thể hay [P] PMI Status Report cho 1 project?"
- "project status" / "status report" / "báo cáo tiến độ" → WORKFLOW P
```

---

### SH-4: Absolute path cho session state

**Thay dòng 768:**
```
Cuối mỗi session, lưu tóm tắt vào `.claude/pm-email-state.md`:
```
**Bằng:**
```
Cuối mỗi session, lưu tóm tắt vào `D:\100.Software\Github\OutlookOkan\.claude\pm-email-state.md`:
```

---

### SH-5: Clarify args "0" vs. menu [0]

**Thay dòng 26:**
```
Nếu không có args hoặc args là "menu" hoặc args là "0": hiển thị MENU rồi hỏi chọn.
```
**Bằng:**
```
Nếu không có args hoặc args là "menu": hiển thị MENU rồi hỏi chọn số.
Nếu đang trong menu và user chọn [0]: kết thúc session, không làm gì thêm.
Nếu args là "0": hiển thị MENU (vì user có thể muốn xem menu, không phải exit).
```

---

### CI-1: Refine Search Implementation

**Thêm sau BƯỚC 3 trong Workflow [6]:**
```
BƯỚC 4 — Xử lý refine từ user:
Nếu user gõ "sender [tên]": gọi search_emails với sender_email=[tên] + keyword=keyword_cũ.
Nếu user gõ "từ [ngày]": gọi lại với date_from=[ngày].
Nếu user gõ "folder [tên]": gọi lại chỉ 1 folder đó.
Hiển thị kết quả mới, giữ nguyên format nhóm theo folder.
```

---

### CI-2: Priority Score Caveat

**Thêm sau Priority Scoring section:**
```
LƯU Ý: Priority score là phỏng đoán tốt nhất (heuristic) từ email subject và sender address patterns.
Điểm C-level (+3) chỉ áp dụng khi email address chứa rõ "ceo", "cto", "director", "giamdoc" hoặc subject/body chứa chức danh. Không chính xác tuyệt đối.
```

---

### CI-3: Cross-Project Comm Report Tool Chain

**Thay "Option B" hiện tại bằng:**
```
Option B — Cross-project report (tất cả 5 folders):
BƯỚC 1: get_project_snapshot(folder, days_back) × 5 folders. Tuần tự.
BƯỚC 2: Tổng hợp top 3 contacts toàn bộ từ 5 snapshots.
BƯỚC 3: get_contact_stats(email=top_contact) cho top 3 contacts xuyên projects.
BƯỚC 4: Tính total volume, avg/project, peak project.
Output: Bảng so sánh 5 projects + cross-project top contacts + nhận xét tổng thể.
```

---

### CI-4: Progress Indicator Pattern

**Thêm vào QUY TẮC CHUNG:**
```
7. Progress indicator: Khi chạy workflow nhiều bước, trước mỗi COM call hiển thị:
   "⏳ [Bước X/Y] Đang [mô tả action]..."
   Ví dụ: "⏳ [2/4] Đang lấy snapshot NCB.FlexCash..."
   Sau khi xong tất cả: "✅ Hoàn thành. Đang tổng hợp kết quả..."
```

---

## CHECKLIST CHO SESSION TIẾP THEO

```
Phase 1 — Brainstorm (discuss với team):
  [ ] Thống nhất style cho 5 templates mới
  [ ] Thống nhất implement batch commands
  [ ] Thống nhất progress indicator format

Phase 2 — Plan:
  [ ] Phân công: ai sửa MV-1 to MV-5?
  [ ] Phân công: ai sửa SH-1 to SH-5?
  [ ] Thứ tự sửa trong file (từ trên xuống để tránh conflict)

Phase 3 — Thực thi:
  [ ] MV-1: Thêm templates [3][4][5][7][8]
  [ ] MV-2: Fix SLA folder
  [ ] MV-3: Fix [T] project_folder ask
  [ ] MV-4: Fix [R] risk_keyword
  [ ] MV-5: Fix [P] DASL search
  [ ] SH-1: Implement batch commands
  [ ] SH-2: Optimize Dashboard calls
  [ ] SH-3: Disambiguate "status"
  [ ] SH-4: Fix session state path
  [ ] SH-5: Fix args "0" logic
  [ ] CI-1 to CI-4: Medium improvements
  
Post-fix:
  [ ] Verify file: đếm tất cả 8 template bodies có đủ không
  [ ] Verify: không còn undefined variables
  [ ] Update workflow.md
  [ ] Update memory
```

---

*Prepared by BMAD Party Mode — 2026-06-24*
*Intended audience: Next session team — Murat, Mary, John, Sally, Winston, Amelia*
