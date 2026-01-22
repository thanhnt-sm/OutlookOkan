# Hướng Dẫn Sử Dụng OutlookOkan (Tiếng Việt)

**OutlookOkan** là một Add-in mã nguồn mở dành cho Microsoft Outlook, được thiết kế để ngăn chặn việc gửi nhầm email, quên đính kèm tệp tin, hoặc gửi nhầm người nhận.

Tài liệu này giải thích chi tiết các chức năng và cách cấu hình phần mềm.

---

## 📋 Mục Lục

1.  [Giới Thiệu Chung](#1-giới-thiệu-chung)
2.  [Cửa Sổ Xác Nhận (Confirmation Window)](#2-cửa-sổ-xác-nhận-confirmation-window)
3.  [Cấu Hình Chi Tiết (Settings)](#3-cấu-hình-chi-tiết-settings)
    *   [General (Cài đặt chung)](#31-general-cài-đặt-chung)
    *   [Do items / Auto check (Tự động bỏ qua/kiểm tra)](#32-do-items--auto-check)
    *   [General 2 (Danh bạ & Nhóm)](#33-general-2-danh-bạ--nhóm)
    *   [Internal Domain (Tên miền nội bộ)](#34-internal-domain-tên-miền-nội-bộ)
    *   [Whitelist (Danh sách tin cậy)](#35-whitelist-danh-sách-tin-cậy)
    *   [Name and Domain (Tên và Tên miền)](#36-name-and-domain-tên-và-tên-miền)
    *   [Alert & Prohibitions (Cảnh báo & Cấm gửi)](#37-alert--prohibitions-cảnh-báo--cấm-gửi)
    *   [Automation (Tự động CC/BCC)](#38-automation-tự-động-ccbcc)

---

## 1. Giới Thiệu Chung

Khi bạn nhấn nút **Send** (Gửi) trong Outlook, OutlookOkan sẽ chặn email lại và hiện lên một **Cửa sổ xác nhận**. Bạn bắt buộc phải kiểm tra (tick) vào các mục quan trọng trước khi nút Gửi sáng lên.

Điều này giúp bạn có "thời gian suy nghĩ lần 2" để tránh các lỗi tai hại như:
*   Gửi nhầm cho khách hàng (External) thay vì đồng nghiệp (Internal).
*   Quên đính kèm file dù trong mail có viết "gửi anh file đính kèm".
*   Gửi file tài liệu mật ra ngoài.

---

## 2. Cửa Sổ Xác Nhận (Confirmation Window)

Giao diện cửa sổ này xuất hiện mỗi khi gửi mail (trừ khi được cấu hình bỏ qua).

### Các vùng thông tin chính:

1.  **Khu vực Cảnh báo (Alerts - Trên cùng)**
    *   Hiển thị các cảnh báo quan trọng màu đỏ.
    *   Ví dụ: "Thiếu file đính kèm", "Gửi cho tên miền lạ", "Chứa từ khóa nhạy cảm".
    *   Bạn phải tick xác nhận đã đọc từng cảnh báo.

2.  **Thông tin Người nhận (To / Cc / Bcc)**
    *   Liệt kê tất cả địa chỉ email sẽ nhận thư.
    *   **Màu đỏ**: Địa chỉ thuộc tên miền bên ngoài (External Domain).
    *   **Màu đen**: Địa chỉ nội bộ (Internal Domain) hoặc nằm trong Whitelist.
    *   Bạn phải tick chọn từng người nhận để xác nhận họ là đúng.

3.  **Tệp đính kèm (Attachments)**
    *   Liệt kê các file đang được gửi kèm.
    *   Cảnh báo nếu file quá lớn hoặc có định dạng nguy hiểm.
    *   **Nút Open**: Cho phép mở file ngay tại đây để kiểm tra nội dung lần cuối.

4.  **Thông tin Email (Mail Info)**
    *   **Sender**: Người gửi (Hữu ích nếu bạn dùng nhiều tài khoản mail hoặc tính năng "Send As").
    *   **Subject**: Tiêu đề thư.
    *   **Mail Type**: Định dạng thư (HTML/Text/RichText).
    *   **Deferred Delivery**: Tính năng **Gửi chậm**.
        *   Nhập số phút vào ô này (ví dụ: `10`). Email sẽ nằm trong Outbox 10 phút trước khi thực sự được gửi đi. Rất hữu ích để "thu hồi" thư nếu chợt nhớ ra điều gì đó sau khi bấm gửi.

5.  **Nội dung (Mail Body)**
    *   Cho xem trước nội dung thư (dạng Text) để rà soát nhanh.

➡ **Quy tắc**: Nút **Send** chỉ kích hoạt (enable) khi bạn đã tick chọn **TẤT CẢ** các ô kiểm trong cửa sổ này.

---

## 3. Cấu Hình Chi Tiết (Settings)

Để mở cài đặt: Trên thanh Ribbon của Outlook, chọn tab **OutlookOkan** -> **Settings**.

### 3.1. General (Cài đặt chung)

*   **Enable forgotten to attach alert**: Bật cảnh báo nếu trong thư có chữ "đính kèm", "attach"... mà không có file nào được attach.
*   **Is not treated as attachments at html embedded files**: Không coi các hình ảnh chèn trực tiếp trong nội dung (inline images, chữ ký) là file đính kèm. (Nên bật để tránh phiền phức).
*   **Is enable recipients are sorted by domain**: Tự động sắp xếp danh sách người nhận theo tên miền trong cửa sổ xác nhận để dễ nhìn.
*   **Always add the sender's address to Cc**: Luôn tự động Cc cho chính mình.
*   **Show confirmation at Send Meeting Request**: Hiện cửa sổ xác nhận cả khi gửi lời mời họp.

### 3.2. Do items / Auto check

Nhóm cài đặt để giảm bớt thao tác thừa:

*   **Do not show confirmation screen**:
    *   **If all recipients are same domain**: Không hiện cửa sổ xác nhận nếu tất cả người nhận đều cùng tên miền với người gửi (gửi nội bộ).
    *   **If all recipients are in the whitelist**: Không hiện nếu tất cả người nhận đều là người quen (trong Whitelist).
*   **Auto Check Config** (Tự động tick chọn):
    *   **Auto check if all recipients are same domain**: Cửa sổ vẫn hiện, nhưng tự động tick sẵn cho các email cùng tên miền.
    *   **Auto check attachments**: Tự động tick xác nhận file đính kèm (Không khuyến khích, nên tự check tay cho an toàn).

### 3.3. General 2 (Danh bạ & Nhóm)

*   **Address Book**:
    *   **Is warning if recipients is not registered**: Cảnh báo nếu gửi cho ai đó KHÔNG CÓ trong danh bạ Outlook (Contacts) của bạn.
    *   **Is prohibits sending...**: Cấm gửi luôn nếu người nhận không có trong danh bạ (Chế độ bảo mật cao).
*   **Get members of the list**: Hỗ trợ bung (expand) các Distribution List hoặc Contact Group để kiểm tra từng thành viên bên trong.

### 3.4. Internal Domain (Tên miền nội bộ)

**Rất quan trọng**. Hãy khai báo tên miền công ty bạn vào đây.
*   Ví dụ: `company.com`
*   Tác dụng: Các email đuôi `@company.com` sẽ được coi là "người nhà" (Internal), hiển thị màu đen thay vì màu đỏ cảnh báo.

### 3.5. Whitelist (Danh sách tin cậy)

Thêm các email khách hàng thân thiết hoặc đối tác thường xuyên.
*   Nhập email đầy đủ (ví dụ: `partner@gmail.com`) hoặc tên miền (`@partner.com`).
*   **Is Skip Confirmation**: Nếu tick vào cột này, email gửi đến địa chỉ này sẽ không cần xác nhận (hoặc autocheck tùy cài đặt).

### 3.6. Name and Domain (Tên và Tên miền)

Tính năng nâng cao để chống giả mạo hoặc nhầm lẫn.
*   Kiểm tra sự khớp nhau giữa "Tên hiển thị" và "Địa chỉ email".
*   Ví dụ: Bạn định nghĩa Tên "Sếp Tổng" phải đi với email `ceo@company.com`. Nếu một email lạ `hacker@gmail.com` nhưng đặt tên hiển thị là "Sếp Tổng", phần mềm sẽ cảnh báo.

### 3.7. Alert & Prohibitions (Cảnh báo & Cấm gửi)

OutlookOkan cung cấp 3 loại danh sách đen/cảnh báo mạnh:

1.  **Keyword and Recipients**:
    *   Cảnh báo nếu gửi thư chứa từ khóa X cho người nhận Y.
    *   Ví dụ: Từ khóa "Báo cáo tài chính" gửi cho `@gmail.com`.

2.  **Alert Keyword for Body/Subject**:
    *   Quét nội dung hoặc tiêu đề thư.
    *   **Alert Keyword**: Từ khóa cần bắt (ví dụ: "Mật", "Confidential", "Lương").
    *   **Sending Forbid**: Nếu tick chọn, phần mềm sẽ **KHÓA NÚT GỬI**. Bạn buộc phải xóa từ khóa đó đi mới gửi được.

3.  **Alert Mail Address**:
    *   Cảnh báo đặc biệt khi gửi cho địa chỉ cụ thể.
    *   Ví dụ: Gửi cho `all-company@domain.com` -> Hiện cảnh báo "Bạn có chắc chắn muốn gửi email cho TOÀN BỘ CÔNG TY không?".

### 3.8. Automation (Tự động CC/BCC)

Tự động thêm người nhận vào Cc hoặc Bcc dựa trên điều kiện:

*   **By Keyword**: Nếu thư có chữ "Hóa đơn", tự động Cc cho `ketoan@company.com`.
*   **By Mail Address/Domain**: Nếu gửi cho Khách hàng A, tự động Bcc cho `sep@company.com`.
*   **By Attached File**: Nếu có đính kèm file (ví dụ file báo giá), tự động Cc cho Quản lý.

---

## 4. Mẹo Sử Dụng & Tối Ưu Hóa (Power Tips)

Phần này hướng dẫn bạn cài đặt OutlookOkan để đạt được sự cân bằng tốt nhất giữa **An toàn** và **Tốc độ**.

### 5 Mức độ cấu hình được khuyến nghị:

#### Mức 1: An toàn tuyệt đối (Mặc định)
*   **Cấu hình**: Để mặc định mọi thứ.
*   **Hành vi**: Cửa sổ xác nhận luôn hiện ra với mọi email. Bạn phải tick từng người.
*   **Phù hợp**: Người mới dùng, hoặc người thường xuyên gửi sai email quan trọng.

#### Mức 2: Hiệu quả & An toàn (Khuyên dùng)
*   **Cấu hình**:
    *   Bật: `Auto check if all recipients are same domain` (Tự động tick nếu cùng tên miền).
    *   Whitelist: Thêm các tên miền đối tác thân thiết (ví dụ `@partner.com`).
*   **Hành vi**:
    *   Gửi nội bộ: Cửa sổ hiện ra nhưng nút **Send** sáng ngay lập tức (do đã được tự động tick). Bạn chỉ cần liếc qua và bấm Enter.
    *   Gửi ngoài: Phải tự tick chọn người ngoài.
*   **Lợi ích**: Vẫn có bước xác nhận cuối cùng nhưng không tốn click chuột cho email nội bộ.

#### Mức 3: Tự động hóa cao (Auto-Pilot)
*   **Cấu hình**:
    *   Bật: `Do not show confirmation screen > If all recipients are same domain` (Tắt xác nhận khi gửi nội bộ).
    *   Bật: `If all recipients are in the whitelist`.
    *   Vào **Whitelist**: Thêm các email/domain hay gửi, tick chọn cột **Check** hoặc **Skip**.
*   **Hành vi**:
    *   Gửi cho người quen/nội bộ: **Không hiện cửa sổ gì cả**. Email đi thẳng.
    *   Gửi cho người lạ: Mới hiện cửa sổ xác nhận.
*   **Lưu ý**: Rủi ro quên đính kèm file nếu bạn tắt xác nhận hoàn toàn.

### Cách tắt cửa sổ xác nhận khi gửi nội bộ:
1.  Vào **Settings** -> **Do items / Auto check**.
2.  Tick vào ô: `Do not show confirmation screen if all recipients are same domain`.
3.  Đảm bảo bạn đã khai báo tên miền công ty ở tab **Internal Domain**.

### Cách sử dụng Whitelist hiệu quả:
*   Vào tab **Whitelist**.
*   Thêm email/domain (ví dụ: `@gmail.com` - cẩn thận với domain công cộng!).
*   **Cột "Skip Confirmation"**: Nếu tick vào đây, và bạn gửi thư CHỈ cho những người này -> Cửa sổ xác nhận sẽ KHÔNG thiện.

---

## 5. Hỏi & Đáp (Q&A)

**Q: Tại sao nút Send bị mờ (Disable) và tôi không thể gửi thư?**
A: Bạn chưa xác nhận hết các cảnh báo. Hãy chắc chắn:
1.  Đã tick vào tất cả các ô vuông bên cạnh tên người nhận (To/Cc/Bcc).
2.  Đã tick vào tất cả các dòng Cảnh báo màu đỏ (trên cùng) hoặc thông báo file đính kèm.
3.  Kiểm tra xem nội dung thư có chứa "Từ khóa cấm gửi" (Sending Forbid) hay không. Nếu có, bạn bắt buộc phải xóa từ đó đi mới gửi được.

**Q: Tôi đã cấu hình "Không hiện xác nhận khi gửi nội bộ", nhưng sao nó vẫn hiện?**
A: Có thể vì:
1.  Bạn chưa khai báo tên miền công ty trong tab **Internal Domain**.
2.  Trong email có lẫn 1 người nhận bên ngoài (External).
3.  Bạn đang bật chế độ `Is Show Confirmation To Multiple Domain` (Hiện xác nhận khi gửi đa tên miền) - Cài đặt này có độ ưu tiên cao nhất, sẽ ghi đè các cài đặt bỏ qua khác.

**Q: Làm sao để "Thu hồi" email đã lỡ bấm gửi?**
A: Hãy dùng tính năng **Deferred Delivery** (Gửi chậm).
*   Trong cửa sổ xác nhận, nhập số `1` hoặc `2` (phút) vào ô Deferred Delivery.
*   Khi bấm Send, email sẽ nằm im trong **Outbox** 1-2 phút.
*   Nếu chợt nhớ ra sai sót, bạn chỉ cần vào Outbox mở email ra và sửa lại.

**Q: Màu đỏ và màu đen ở tên người nhận có ý nghĩa gì?**
*   **Màu đen**: An toàn. Đây là email nội bộ (Internal) hoặc nằm trong Whitelist.
*   **Màu đỏ**: Cảnh báo. Đây là email bên ngoài (External) chưa được tin cậy. Hãy nhìn kỹ trước khi tick!

**Q: Tôi muốn tự động thêm email của chính mình vào Bcc và KHÔNG muốn hiện cửa sổ xác nhận thì làm thế nào?**  
A: Để thực hiện điều này ("Silent Auto-Bcc"), bạn cần cấu hình phối hợp 2 cài đặt sau:
1.  Vào **Settings** > thẻ **General**.
2.  Trong nhóm **General** (phía trên), tích chọn **"Always add the sender's address to Bcc"**.
    *   *Tác dụng*: Tự động thêm bạn vào Bcc mỗi khi gửi mail.
3.  Trong nhóm **Do not show confirmation screen** (phía dưới), tích chọn **"If all recipients are in the whitelist"** (hoặc "If all recipients are in the same domain" nếu bạn chỉ gửi nội bộ).
    *   *Giải thích*: OutlookOkan thông minh tự động coi địa chỉ người gửi (bạn) là một địa chỉ "An toàn/Whitelist". Do đó, nếu bạn gửi mail cho những người khác cũng nằm trong Whitelist (hoặc nội bộ), và chính bạn (trong Bcc) cũng an toàn, thì toàn bộ danh sách đều an toàn -> Cửa sổ xác nhận sẽ được bỏ qua.
    *   *Lưu ý*: Nếu bạn gửi mail cho một người lạ (External) không có trong Whitelist, cửa sổ xác nhận vẫn sẽ hiện ra để cảnh báo bạn (đây là tính năng an toàn).

---
*Tài liệu được cập nhật ngày 22/01/2026 dựa trên phiên bản mới nhất.*
