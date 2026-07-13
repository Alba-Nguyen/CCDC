# Alba - CCDC

Tool đối chiếu CCDC cho tài khoản 2421 và 2422.

## Cách sử dụng

1. Tự chép file mẫu nội bộ `Check_CCDC 1.xlsx` vào thư mục tool. File này không được lưu trên repository Public.
2. Đặt hai file nguồn vào cùng thư mục với tool và giữ đúng tên:
   - `Bảng kê chứng từ.xlsx`
   - `Bảng tổng hợp chi phí chờ phân bổ.xlsx`
3. Không mở các file nguồn trong Excel khi chạy.
4. Chạy `Chay_Check_CCDC.cmd`.
5. Chờ thông báo `HOÀN TẤT`; báo cáo sẽ có tên `Check CCDC_YYYYMMDD_HHMM.xlsx`.

Tool tự nhận số dòng thực tế, chạy công thức cho toàn bộ dữ liệu nguồn và tạo hai sheet `CheckChiTiet_2421`, `CheckChiTiet_2422`.

Chi tiết cách lấy báo cáo nguồn được ghi trong `TOOL_Check_CCDC/HUONG_DAN.txt` và sheet `Inf` của báo cáo kết quả.
