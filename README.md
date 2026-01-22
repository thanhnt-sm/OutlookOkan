Okan cho Outlook (Add-in ngăn chặn gửi mail nhầm)
========

English readme is [here](https://github.com/t-miyake/OutlookOkan/blob/master/README_en.md).

Okan cho Outlook (Outlook Okan) là một add-in dành cho Microsoft Office Outlook.  

Để ngăn chặn việc gửi nhầm, cửa sổ xác nhận sẽ hiển thị trước khi gửi mail.  
Nó là một add-in sẽ lo lắng và kiểm tra giúp bạn giống như một người mẹ (Okan).  

Vì là mã nguồn mở hoàn toàn, bạn có thể yên tâm sử dụng ngay cả với các email liên quan đến thông tin mật.  
Ngoài ra, add-in còn có các tính năng tùy chọn tiện lợi như cảnh báo bằng từ khóa, tự động thêm Cc/Bcc.  

Bạn có thể tải xuống từ [releases](https://github.com/t-miyake/OutlookOkan/releases).  
※ Chúng tôi cũng phân phối phiên bản có tên add-in trung tính (ít gây chú ý).

Bạn có thể sử dụng miễn phí và là mã nguồn mở, nhưng không có hỗ trợ và không bảo hành. ([Giấy phép](https://github.com/t-miyake/OutlookOkan/blob/master/LICENSE))  
Nếu bạn cần tùy chỉnh hoặc hỗ trợ riêng, vui lòng liên hệ trực tiếp.  

Cửa sổ xác nhận trước khi gửi  
![Screenshot 1](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v2.5.0_01.png)  

Cửa sổ cài đặt (Cài đặt chung)  
![Screenshot 2](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v2.7.0_04.png)

Cửa sổ cài đặt (Gửi chậm)  
![Screenshot 3](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v2.7.0_05.png)

Thông báo cấm gửi  
![Screenshot 4](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v2.5.0_03.png)

Thông tin phiên bản  
![Screenshot 5](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v2.6.1_02.png)

## Môi trường hỗ trợ

- Windows 7 / 8 / 8.1 / 10 / 11
- Microsoft Outlook 2013 / 2016 / 2019 / 2021 / Microsoft 365 Apps (bản 32bit và 64bit)
- .NET Framework 4.6.2 trở lên

## Danh sách tính năng (Tổng quan)

- Xác nhận trước khi gửi mail, v.v.
  - Hiển thị cửa sổ xác nhận trước khi gửi mail, không thể gửi nếu không tích vào tất cả các mục.
  - Có thể cài đặt để không hiển thị xác nhận đối với mail gửi nội bộ (trong công ty).
  - Tên miền bên ngoài (ngoài công ty) được hiển thị màu đỏ.
  - Hiển thị tiêu đề, địa chỉ người gửi, danh sách file đính kèm, nội dung mail.
  - Cảnh báo nếu quên đính kèm file hoặc file đính kèm có dung lượng lớn.
  - Mở rộng danh sách phân phối (Distribution List) hoặc nhóm liên hệ để hiển thị từng người nhận (Bật/Tắt).
  - Sắp xếp hiển thị người nhận theo tên miền (Bật/Tắt).
  - Luôn tự động thêm địa chỉ người gửi vào Cc hoặc Bcc (Bật/Tắt).

- Chức năng cấm gửi
  - Cấm gửi mail đến các địa chỉ hoặc tên miền đã chỉ định.
  - Cấm gửi mail nếu trong nội dung có chứa từ khóa đã chỉ định.
  - Cấm gửi mail có file đính kèm đến các địa chỉ hoặc tên miền đã chỉ định.
  - Cấm gửi mail có file đính kèm (Bật/Tắt).
  - Cấm gửi mail đến các địa chỉ không có trong danh bạ (Bật/Tắt).
  - Cấm gửi mail nếu số lượng tên miền bên ngoài trong danh sách nhận (To/Cc) quá nhiều (Bật/Tắt).
  - Cấm gửi mail nếu có đính kèm file ZIP đã mã hóa (Bật/Tắt).
  - Khi thuộc trường hợp cấm gửi, sẽ hiển thị thông báo cấm và lý do.

- Danh sách cho phép (Whitelist)
  - Các tên miền hoặc địa chỉ được đăng ký trong danh sách cho phép sẽ không cần kiểm tra các mục trong màn hình xác nhận.

- Đăng ký tên và cảnh báo người nhận
  - Hiển thị cảnh báo nếu tên xuất hiện trong nội dung mail không khớp với địa chỉ hoặc tên miền người nhận.

- Đăng ký từ khóa cảnh báo và cảnh báo
  - Nếu từ khóa đã đăng ký xuất hiện trong nội dung hoặc tiêu đề mail, sẽ hiển thị văn bản cảnh báo đã đăng ký.
  - Cũng có thể cài đặt để luôn hiển thị tin nhắn cảnh báo đã đăng ký.

- Đăng ký địa chỉ cảnh báo và cảnh báo
  - Hiển thị văn bản cảnh báo khi gửi mail đến địa chỉ hoặc tên miền đã đăng ký.  
  - Có thể cài đặt văn bản cảnh báo riêng cho từng người nhận.

- Cảnh báo số lượng tên miền bên ngoài (To/Cc) và tự động chuyển sang Bcc
  - Hiển thị cảnh báo khi số lượng tên miền bên ngoài trong danh sách nhận (To/Cc) quá nhiều.
  - Tự động chuyển đổi địa chỉ bên ngoài trong danh sách nhận (To/Cc) sang Bcc khi số lượng quá nhiều.
  - Bắt buộc chuyển đổi tất cả người nhận sang Bcc.

- Tự động thêm Cc/Bcc (Theo từ khóa)
  - Nếu nội dung mail chứa từ khóa đã chỉ định, tự động thêm địa chỉ đã chỉ định vào Cc hoặc Bcc.

- Tự động thêm Cc/Bcc (Theo người nhận)
  - Khi gửi mail đến người nhận đã chỉ định, tự động thêm địa chỉ đã chỉ định vào Cc hoặc Bcc.

- Tự động thêm Cc/Bcc (Theo file đính kèm)
  - Khi mail có file đính kèm, tự động thêm địa chỉ đã chỉ định vào Cc hoặc Bcc.

- Gửi chậm (Trì hoãn gửi/Tạm giữ)
  - Trì hoãn (tạm giữ) việc gửi mail trong khoảng thời gian đã cài đặt (tính bằng phút).
  - Có thể cài đặt thời gian trì hoãn mặc định cho từng tên miền hoặc địa chỉ email.

- Liên kết tên file đính kèm và người nhận
  - Liên kết tên file đính kèm với địa chỉ email hoặc tên miền người nhận, nếu không khớp sẽ hiển thị cảnh báo.

- Cảnh báo theo người nhận khi có file đính kèm
  - Có thể cài đặt văn bản cảnh báo riêng cho từng người nhận (địa chỉ hoặc tên miền) khi mail có file đính kèm.

- Tự động thêm văn bản vào nội dung mail
  - Có thể tự động thêm văn bản đã chỉ định vào đầu hoặc cuối nội dung mail.

- Bảo mật mail nhận đơn giản
  Các tính năng khi mở mail đã nhận:
  - Cảnh báo nếu tiêu đề có chứa từ khóa đã chỉ định.
  - Hiển thị cảnh báo nếu xác thực bản ghi SPF thất bại.
  - Hiển thị cảnh báo nếu xác thực bản ghi DKIM thất bại.
  - Cảnh báo trước khi mở file đính kèm.
  - Cảnh báo khi mở file ZIP đã mã hóa.
  - Cảnh báo nếu trong file ZIP có chứa file .lnk.
  - Cảnh báo nếu trong file ZIP có chứa file .one.
  - Cảnh báo nếu trong file ZIP có chứa file .docm, xlsm, pptm.
  - Cảnh báo khi mở file đính kèm có chứa macro.

- Khác
  - Hiển thị cảnh báo nếu có đính kèm file ZIP đã mã hóa (Bật/Tắt).
  - Bắt buộc chuyển đổi người nhận sang Bcc.

- Nhập/Xuất cấu hình
  - Nhập/Xuất nội dung cài đặt bằng file CSV.
  - Nhập/Xuất tất cả cài đặt cùng lúc.

- Đa ngôn ngữ
  - Hỗ trợ tổng cộng 10 ngôn ngữ bao gồm tiếng Nhật và tiếng Anh, được thiết kế để có thể thêm ngôn ngữ mới.

## Cách sử dụng

Chi tiết được ghi tại [Wiki(Manual)](https://github.com/t-miyake/OutlookOkan/wiki/Manual).

## Các lỗi đã biết

Chi tiết được ghi tại [Wiki(Known Issues)](https://github.com/t-miyake/OutlookOkan/wiki/Known-Issues).

## Lộ trình phát triển

Chi tiết được ghi tại [Wiki(Roadmap)](https://github.com/t-miyake/OutlookOkan/wiki/Roadmap).
