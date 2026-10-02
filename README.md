# M365 User Activity & Presence Monitor 🚀

Hệ thống bảng điều khiển (Dashboard) giám sát hoạt động & trạng thái trực tuyến (Realtime Presence) của tài khoản Microsoft 365, chuyên dùng cho Quản trị viên (SysAdmin).

---

## 🌟 Tính Năng Chính
1. **Giám Sát Trực Tuyến Realtime:** Hiển thị tức thời ai đang mở Teams/Excel/Word làm việc (🟢 Online, 🔴 Bận/Họp, 🟡 Vắng mặt, ⚫ Ngoại tuyến).
2. **Theo Dõi Tần Suất Sử Dụng:** Thống kê chính xác thời điểm đăng nhập gần nhất (`signInActivity`), phân loại mức độ sử dụng:
   - *Rất tích cực* (≤ 3 ngày)
   - *Thỉnh thoảng* (4 - 14 ngày)
   - *Ít sử dụng* (15 - 30 ngày)
   - *Có nguy cơ bỏ hoang* (> 14 ngày không truy cập)
3. **Tuân Thủ Tuyệt Đối Quyền Riêng Tư:** Chỉ đọc metadata kỹ thuật phục vụ quản trị tài nguyên công vụ, **hoàn toàn không can thiệp nội dung file hay email**.
4. **Giữ Tenant E5 Developer Active:** Tự động tạo lưu lượng truy vấn Microsoft Graph API đều đặn, giúp tenant được tự động gia hạn 90 ngày vĩnh viễn.
5. **Xuất Báo Cáo:** Hỗ trợ xuất danh sách ra file CSV chỉ bằng 1 cú click.

---

## 🛠️ Yêu Cầu Quyền Microsoft Entra ID (Azure AD)
Ứng dụng sử dụng **Client Credentials Flow** (Application Permissions) để kết nối Microsoft Graph. Bạn cần cấp các quyền sau cho App Registration (`SnapSync302` hoặc tạo App mới):

1. Vào [Microsoft Entra admin center](https://entra.microsoft.com) > **App registrations** > Chọn App của bạn.
2. Vào **API permissions** > **Add a permission** > Chọn **Microsoft Graph** > Chọn **Application permissions**.
3. Tìm và tích chọn 3 quyền sau:
   - `User.Read.All` (Đọc danh sách người dùng)
   - `AuditLog.Read.All` (Đọc lịch sử đăng nhập `signInActivity`)
   - `Presence.Read.All` (Đọc trạng thái đèn xanh/đỏ thời gian thực)
4. Bấm **Grant admin consent for [Tên tổ chức]** (Bắt buộc để các quyền có hiệu lực).

---

## 🚀 Hướng Dẫn Deploy Lên Render.com (Miễn Phí 100%)

### Cách 1: Đẩy mã nguồn lên GitHub rồi liên kết Render
1. Tạo một repository mới trên GitHub (chế độ **Private** để bảo mật).
2. Đẩy toàn bộ thư mục này lên GitHub:
   ```bash
   git init
   git add .
   git commit -m "Initial commit M365 Activity Monitor"
   git remote add origin https://github.com/<your-username>/<your-repo>.git
   git branch -M main
   git push -u origin main
   ```
3. Đăng nhập vào [Render.com](https://render.com).
4. Bấm **New +** > Chọn **Web Service** > Chọn repository GitHub bạn vừa tạo.
5. Thiết lập thông số:
   - **Name:** `m365-activity-monitor`
   - **Environment:** `Node`
   - **Build Command:** `npm install`
   - **Start Command:** `node server.js`
   - **Instance Type:** `Free`
6. Cuộn xuống phần **Environment Variables** (Biến môi trường), thêm 3 biến:
   - `TENANT_ID`: *(Tenant ID của bạn)*
   - `CLIENT_ID`: `49c7d651-aca9-4e46-ae4e-651f1e851e3c`
   - `CLIENT_SECRET`: *(Giá trị Value Secret bạn vừa tạo)*
7. Bấm **Create Web Service**. 
8. Chờ khoảng 1-2 phút, Render sẽ cấp cho bạn một tên miền miễn phí dạng: `https://m365-activity-monitor-xxxx.onrender.com`.

---

## 💻 Chạy Thử Trên Máy Tính Cục Bộ (Local)
1. Tạo file `.env` từ `.env.example`:
   ```bash
   cp .env.example .env
   ```
2. Điền `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` vào file `.env`.
3. Khởi động server:
   ```bash
   npm start
   ```
4. Mở trình duyệt truy cập: `http://localhost:3000`.
