Outlook Okan (Add-in ngăn chặn gửi mail nhầm)
========

Tiếng Nhật xem tại [đây](https://github.com/t-miyake/OutlookOkan/)。

Outlook Okan là một add-in dành cho Microsoft Office Outlook.  

Add-in này sẽ hiển thị một cửa sổ xác nhận trước khi gửi email.  
Điều đó giúp ngăn chặn việc gửi mail nhầm.  

Đối với các email nhạy cảm, bạn có thể yên tâm rằng add-in này là mã nguồn mở hoàn toàn.  
Ngoài ra còn có các tính năng tùy chọn hữu ích như cảnh báo từ khóa và tự động thêm Cc/Bcc.  

Bạn có thể tải xuống add-in này tại [đây](https://github.com/t-miyake/OutlookOkan/releases).  

Add-in này là mã nguồn mở và sử dụng miễn phí, nhưng không được hỗ trợ và không được đảm bảo.  
([Giấy phép](https://github.com/t-miyake/OutlookOkan/blob/master/LICENSE))  
Nếu bạn cần tùy chỉnh hoặc hỗ trợ, vui lòng liên hệ trực tiếp với chúng tôi.  

Cửa sổ xác nhận trước khi gửi.  
![Screenshot 1](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/en/Screenshot_v2.5.0_01_en.png)  

Cửa sổ cài đặt (cài đặt chung)  
![Screenshot 2](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/en/Screenshot_v2.7.0_04_en.png)  

Cửa sổ cài đặt (gửi chậm)  
![Screenshot 3](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/en/Screenshot_v2.7.0_05_en.png)  

Cửa sổ cảnh báo  
![Screenshot 4](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/en/Screenshot_v2.5.0_03_en.png)  

Cửa sổ giới thiệu  
![Screenshot 5](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/en/Screenshot_v2.6.1_02_en.png)  

## Yêu cầu hệ thống

- Windows 7 / 8 / 8.1 / 10 / 11
- Microsoft Outlook 2013 / 2016 / 2019 / 2021 / Microsoft 365 Apps (32bit hoặc 64bit)
- .NET Framework 4.6.2 trở lên

## Danh sách tính năng (tổng quan)

- Xác nhận trước khi gửi email và các chức năng khác.  
  - Cửa sổ xác nhận hiển thị trước khi gửi mail và tất cả các mục phải được kiểm tra trước khi gửi.
  - Có thể không hiển thị xác nhận trước khi gửi, ví dụ như email gửi đến tên miền nội bộ.
  - Tên miền bên ngoài được hiển thị bằng chữ màu đỏ.
  - Hiển thị chủ đề và địa chỉ người gửi, danh sách tệp đính kèm và nội dung email.
  - Cảnh báo thiếu tệp đính kèm hoặc tệp đính kèm lớn.
  - Mở rộng danh sách phân phối và nhóm liên hệ để hiển thị từng người nhận (có thể bật hoặc tắt).  
  - Sắp xếp và hiển thị người nhận theo tên miền (có thể bật hoặc tắt).  
  - Luôn tự động thêm địa chỉ nguồn vào Cc/Bcc (có thể bật hoặc tắt).  

- Ngăn chặn việc gửi mail khớp với các điều kiện.
  - Ngăn chặn gửi email đến đích hoặc tên miền đã chỉ định.
  - Ngăn chặn gửi email có tệp đính kèm đến các đích hoặc tên miền đã chỉ định.
  - Ngăn chặn gửi email có tệp đính kèm (có thể bật hoặc tắt).
  - Ngăn chặn gửi email đến các địa chỉ không được đăng ký trong Danh bạ (có thể bật hoặc tắt).
  - Ngăn chặn gửi email có chứa từ khóa đã chỉ định trong nội dung.
  - Ngăn chặn gửi mail khi số lượng tên miền bên ngoài của người nhận (To/Cc) quá lớn.
  - Ngăn chặn gửi mail nếu có kèm theo tệp ZIP được mã hóa.

- Danh sách cho phép (Allowlist)
  - Các tên miền và địa chỉ email nằm trong danh sách cho phép không cần phải kiểm tra trên cửa sổ xác nhận.

- Đăng ký tên và người nhận và cảnh báo
  - Nếu tên trong nội dung thư và địa chỉ hoặc tên miền của người nhận không khớp, một cảnh báo sẽ hiển thị.

- Đăng ký từ khóa cảnh báo và tin nhắn cảnh báo.
  - Nếu từ khóa đã đăng ký có trong nội dung hoặc chủ đề của email, tin nhắn cảnh báo đã đăng ký sẽ được hiển thị.
  - Cũng có thể luôn hiển thị tin nhắn cảnh báo đã đăng ký.

- Đăng ký người nhận cảnh báo và tin nhắn cảnh báo.
  - Một tin nhắn cảnh báo hiển thị khi gửi email đến địa chỉ hoặc tên miền đã đăng ký.
  - Tin nhắn cảnh báo cũng có thể được thiết lập theo từng người nhận.

- Cảnh báo về số lượng tên miền bên ngoài của người nhận (To/Cc) và tự động thay đổi thành Bcc.
  - Tin nhắn cảnh báo khi số lượng tên miền bên ngoài của người nhận (To/Cc) quá lớn.
  - Khi số lượng tên miền bên ngoài của người nhận (To/Cc) quá lớn, người nhận bên ngoài (To/Cc) sẽ tự động được chuyển thành Bcc.
  - Buộc tất cả người nhận phải được chuyển đổi thành Bcc.

- Tự động thêm Cc/Bcc (theo từ khóa)
  - Nếu từ khóa đã chỉ định có trong nội dung email, địa chỉ đã chỉ định sẽ tự động được thêm vào Cc và Bcc.

- Tự động thêm Cc/Bcc (theo người nhận)
  - Tự động thêm địa chỉ đã chỉ định vào Cc hoặc Bcc trong email gửi đến người nhận đã chỉ định.

- Tự động thêm Cc/Bcc (theo tệp đính kèm)
  - Tự động thêm địa chỉ đã chỉ định vào Cc và Bcc trong các email có đính kèm tệp.

- Gửi chậm (Trì hoãn gửi)
  - Bạn có thể trì hoãn (tạm giữ) việc gửi email trong một khoảng thời gian đã định (tính bằng phút).
  - Bạn có thể thiết lập thời gian trì hoãn mặc định cho từng tên miền hoặc địa chỉ email.

- Liên kết tên tệp đính kèm với người nhận
  - Liên kết tên tệp đính kèm với địa chỉ email hoặc tên miền của người nhận và hiển thị cảnh báo nếu chúng không khớp.

- Cảnh báo cho từng người nhận khi có tệp đính kèm
  - Cho phép cấu hình văn bản cảnh báo cho từng người nhận (địa chỉ hoặc tên miền) khi có tệp đính kèm.

- Tự động thêm văn bản vào nội dung email
  - Một cụm từ đã chỉ định có thể được tự động thêm vào đầu hoặc cuối nội dung email.

- Khác
  - Hiển thị cảnh báo nếu có đính kèm tệp ZIP được mã hóa (có thể bật hoặc tắt).

- Nhập và xuất cài đặt
  - Bạn có thể nhập và xuất các cài đặt của mình dưới dạng tệp CSV.

- Hỗ trợ đa ngôn ngữ
  - Hỗ trợ tổng cộng 10 ngôn ngữ, bao gồm tiếng Nhật và tiếng Anh. Được thiết kế để cho phép thêm các ngôn ngữ bổ sung.

## Hướng dẫn sử dụng

[Wiki(Manual)](https://github.com/t-miyake/OutlookOkan/wiki/Manual)  

## Các lỗi đã biết

[Wiki(Known Issues)](https://github.com/t-miyake/OutlookOkan/wiki/Known-Issues)  

## Lộ trình phát triển

[Wiki(Roadmap)](https://github.com/t-miyake/OutlookOkan/wiki/Roadmap)  
