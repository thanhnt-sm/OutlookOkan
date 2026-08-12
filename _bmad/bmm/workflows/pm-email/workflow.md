---
name: pm-email
description: PM Email Workflow v3 (PMI Edition) — 13 workflows đầy đủ, chuẩn PMI/PMBOK 7, menu nhóm, full chain, bảo toàn context đa session
---

# PM Email Workflow v3 (PMI Edition)

**Mục tiêu:** Tool email toàn diện nhất cho PM — hàng ngày, họp, tracking, soạn thảo, PMI reporting, stakeholder management. Chuẩn hóa theo PMBOK 7.

**Kích hoạt:** `/pm-email` trong Claude Code → hiện menu nhóm rõ ràng.

**Yêu cầu:** outlook-mcp-secure đang chạy, Outlook Desktop đang mở.

---

## ARCHITECTURE v3

```
/pm-email
│
├── SESSION INIT
│   ├── Đọc config.toml → MY_EMAIL = account_name
│   ├── Tải pm-email-state.md (nếu có)
│   └── Tính days_back thông minh (T2=3, khác=1)
│
├── MENU (nhóm rõ ràng)
│   ├── HÀNG NGÀY:     [1] Brief  [8] Triage
│   ├── HỌP:           [2] Pre-Meeting  [S] Stakeholder Temp
│   ├── THEO DÕI:      [3] Pending+SLA  [4] Dashboard
│   ├── SOẠN THẢO:     [5] Smart Draft  [T] Template
│   ├── PHÂN TÍCH:     [6] Search  [7] Report  [9] Thread
│   ├── PMI CHUẨN:     [P] Status  [R] Risk Comm
│   ├── TOÀN DIỆN:     [A] Full Audit
│   └── [0] EXIT
│
└── SESSION SAVE → .claude/pm-email-state.md
```

---

## DANH SÁCH 13 WORKFLOWS

| # | Tên | Mục tiêu | Tool chain chính |
|---|-----|----------|------------------|
| 1 | Daily Brief | Briefing đầy đủ + priority thông minh | stats→snapshot×N→flagged×5→score |
| 2 | Pre-Meeting | Chuẩn bị họp + stakeholder level | snapshot→search→flagged→thread→read |
| 3 | Pending+SLA | Flagged + email chưa được reply | flagged×6→sent+noReply→approval |
| 4 | Dashboard | Health check + Stakeholder Temp | snapshot×5→flagged×5→contact_stats |
| 5 | Smart Draft | Reply đúng tone + confirm | search→read→thread→sentItems→draft |
| 6 | Deep Search | Cross-folder search + drill-down | search×folders→refine→action |
| 7 | Comm Report | Báo cáo đơn/cross-project | snapshot→list→contact_stats→flagged |
| 8 | Quick Triage | Phân loại nhanh + batch action | stats→list→action |
| 9 | Thread Dive | Đọc sâu thread + action items | search→thread→read×2 |
| S | Stakeholder Temp | Ai silent/escalating? | snapshot×2+contact_stats→analysis |
| T | Template Draft | Template chuẩn PM → điền context | search→template→compose |
| P | PMI Status | Status report EVM chuẩn | snapshot→flagged→search risk→compose |
| R | Risk Comm | Risk notification chuẩn PMI | snapshot→search→compose |
| A | Full Audit | Toàn diện tất cả | stats+snapshot×5+flagged×5+SLA+contacts |

---

## KEY IMPROVEMENTS v3 (so với v2)

### P1 Critical Fixes:
- **Weekend logic**: Thứ Hai → `days_back=3`, ngày thường → `days_back=1`
- **User email**: Đọc từ `config.toml[outlook.account_name]` thay vì hardcode
- **Menu [0]**: Đổi thành EXIT/BACK (anti-pattern trước đó là Full Audit)

### P2 UX Improvements:
- **Grouped menu**: 6 nhóm rõ ràng thay vì flat 10 items
- **Deep Search**: Search ngay với input tối thiểu, refine sau
- **Cross-folder default**: Search tất cả folders khi không chỉ định
- **COM optimization**: Chỉ load folders có unread > 0 trong Daily Brief

### P3 New Features:
- **Workflow 3 expanded**: SLA Monitor + Approval Tracking
- **Workflow 4 expanded**: Stakeholder Temperature Check
- **Priority scoring**: Business impact (sender seniority + keywords) thay vì tuổi email
- **Workflow 7 expanded**: Cross-project report option

### New Workflows (PMI Edition):
- **[S] Stakeholder Temperature**: PMBOK 7 Stakeholder Engagement Assessment Matrix
- **[P] PMI Status Report**: EVM indicators (SPI/CPI/EAC) + chuẩn PMI
- **[R] Risk Communication**: Risk notation chuẩn PMI Risk Management
- **[T] Template Draft**: 8 templates chuẩn PM scenario

### Multi-session Context:
- Session state saved to `.claude/pm-email-state.md`
- Loaded automatically ở đầu mỗi phiên

---

## PMI FRAMEWORK INTEGRATION

### PMBOK 7 Performance Domains áp dụng:

| Domain | Áp dụng trong workflow |
|--------|------------------------|
| Stakeholder | [S] Engagement Matrix, [2] Pre-Meeting level assessment |
| Communication | [7] Comm Report health, [5] R.A.G.E. framework |
| Uncertainty/Risk | [R] Risk Communication chuẩn |
| Delivery | [P] Status Report + EVM |
| Team | [4] Dashboard + Stakeholder Temp |
| Measurement | [P] EVM: SPI/CPI/EAC |

### EVM Basics (Workflow P):
- SPI > 1.0 = ahead of schedule (tiến độ tốt hơn kế hoạch)
- CPI > 1.0 = under budget (chi phí thấp hơn kế hoạch)
- SPI < 1.0 = behind schedule → cần status report + escalation plan
- CPI < 1.0 = over budget → cần explain + corrective action

### R.A.G.E. Email Framework (Workflow 5, T, P):
- **R**ole: Vai trò PM trong context project cụ thể
- **A**udience: Cấp bậc/quan hệ người nhận (từ email patterns)
- **G**oal: Xác nhận / làm rõ / escalate / đóng / inform
- **E**ssentials: Thông tin cần thiết từ thread + project context

---

## SESSION STATE FORMAT

File: `.claude/pm-email-state.md`

```markdown
---
date: [DD/MM/YYYY]
---
## Session Summary
- Workflows ran: [1, 3, 5]
- Key findings: [3 điểm quan trọng]
- Urgent items: [Email/item cần tiếp tục]
- Active projects: [Projects có hoạt động trong session]
- Next suggested: [Workflow gợi ý đầu phiên sau]
```

---

## ALLOWED FOLDERS

| Folder | Loại | Dùng trong |
|--------|------|-----------|
| Inbox | System | Tất cả inbound |
| Sent Items | System | SLA check, học tone |
| Drafts | System | Review chưa gửi |
| Deleted Items | System | Recovery |
| PVC.CLIMS | Project | Dự án PVC - CLIMS |
| NCB.FlexCash | Project | Dự án NCB - FlexCash |
| YMH.sCPM | Project | Dự án YMH - sCPM |
| Softmart | Project | Internal |
| PVC.Collection | Project | Dự án PVC - Collection |

---

## QUICK REFERENCE

| Gõ | Workflow |
|----|---------|
| `/pm-email` | Menu đầy đủ |
| `/pm-email 1` | Daily Brief ngay |
| `/pm-email họp NCB` | Pre-Meeting NCB.FlexCash |
| `/pm-email 3` | Pending + SLA check |
| `/pm-email s` | Stakeholder Temperature |
| `/pm-email p` | PMI Status Report |
| `/pm-email r rủi ro delay` | Risk Communication về delay |
| `/pm-email t kickoff` | Template kickoff email |
| `/pm-email a` | Full Audit (có confirm) |
