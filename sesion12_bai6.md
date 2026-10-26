# Bài 6 : Hoàn thiện bộ hồ sơ triển khai nền tảng Smart-Home
## Phần 1. Phân tích: So sánh bảng Excel "Xấu" và bảng "Đúng"

* **Bảng "Xấu" (Vi phạm chuẩn dữ liệu):** Dùng `Merge Cell` (trộn ô) ở tiêu đề, có dòng trống ngắt quãng và dữ liệu bị trộn lẫn kiểu (ví dụ: số gõ chung với chữ). Điều này khiến các hàm như `VLOOKUP` báo lỗi `#REF!` hoặc trả về sai kết quả.
* **Bảng "`Đúng`" (Chuẩn CSDL / Dữ liệu phẳng):** Tuyệt đối không dùng `Merge Cell`. Mỗi dòng là một bản ghi độc lập, mỗi cột chứa một kiểu dữ liệu thuần nhất, tiêu đề nằm ở một dòng duy nhất trên cùng. Giúp các hàm tra cứu và lọc dữ liệu chạy chính xác 100%.

---

## Phần 2. Triển khai học liệu dự án Smart-Home

### 1. Tài liệu Yêu cầu Phần mềm (SRS - Tóm tắt)

* **Mục tiêu:** Xây dựng nền tảng tích hợp Smart-Home kết hợp sàn đấu giá bất động sản trực tuyến, số hóa toàn bộ quy trình kết nối người mua, môi giới và quản trị viên.
* **Tính năng cốt lõi:**
* Quản lý phân quyền người dùng (Khách hàng, Môi giới, Admin).
* Dashboard giám sát và điều khiển thiết bị IoT từ xa qua Web/Mobile.
* Sàn đấu giá nhà đất trực tuyến theo thời gian thực (Real-time bidding).
* Tự động tính toán hoa hồng môi giới bằng hàm `VLOOKUP`.


* **Yêu cầu phi chức năng:** Tốc độ phản hồi dưới 2 giây với 10.000 bản ghi; bảo mật mật khẩu bằng `bcrypt` và giao thức HTTPS.

### 2. Mô hình hóa tính toán hoa hồng trên Excel

Bảng tham chiếu `Table_Matrix` (Cột A: Min_Value, Cột B: Max_Value, Cột C: Commission_Rate) không dùng `Merge Cell`. Công thức tra cứu tự động:


$$\text{Hoa hồng} = \text{Giá trị giao dịch} \times \text{VLOOKUP}(\text{Giá trị giao dịch}, \text{Table\_Matrix}, 3, \text{TRUE})$$

### 3. Bộ Slide Pitch Deck gọi vốn (Quy tắc 10/20/30)

Gồm 10 slide, mỗi slide mang một thông điệp cốt lõi với nhiều khoảng trắng (White Space):

1. **Bìa:** Smart-Home - Tương lai bất động sản thông minh tại Việt Nam.
2. **Vấn đề:** Thị trường BĐS truyền thống thiếu minh bạch, thiết bị IoT rời rạc.
3. **Giải pháp:** Nền tảng khép kín tích hợp giao dịch BĐS và quản trị Smart-Home.
4. **Mô hình kinh doanh:** Doanh thu từ phí môi giới, hoa hồng thiết bị IoT và dịch vụ bảo trì.
5. **Thị trường:** Đón đầu xu hướng chuyển đổi số phân khúc nhà ở cao cấp.
6. **Công nghệ:** Hạ tầng dữ liệu phẳng, tốc độ xử lý giao dịch $< 2$ giây.
7. **Tài chính:** Tự động hóa dòng tiền và tính toán chiết khấu thời gian thực.
8. **Đội ngũ:** Nhân sự CNTT trẻ, chuyên môn vững vàng.
9. **Kêu gọi đầu tư:** Gọi vốn vòng Seed Capital để hoàn thiện sản phẩm và mở rộng thị trường.
10. **Tổng kết:** Q&A - Smart-Home đồng hành kiến tạo không gian sống tương lai.

---

## Phần 3. Kiểm thử: Khắc phục lỗi trôi vùng dữ liệu với Dấu đô la (`$`)

* **Nguyên nhân lỗi:** Khi dùng `VLOOKUP`, nếu không khóa vùng bảng dò (ví dụ dùng `A2:C10`), khi kéo công thức xuống các ô bên dưới, vùng tham chiếu sẽ bị dịch chuyển thành `A3:C11`, `A4:C12`,... dẫn đến sót dữ liệu và báo lỗi `#N/A`.
* **Cách khắc phục chuẩn xác:** Sử dụng **Tham chiếu tuyệt đối** bằng cách thêm dấu `$` để khóa cố định cả cột và hàng của bảng dò:
`=VLOOKUP(D2, $A$2:$C$10, 3, FALSE)`
* *Trong đó:* `D2` thay đổi linh hoạt theo từng dòng, còn `$A$2:$C$10` được giữ nguyên cố định hoàn toàn khi kéo công thức.