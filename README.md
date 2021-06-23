# Nguyên tắc hoạt động của Project

## Xử lý hình ảnh Webcam từ Browser
 - Khi người dùng bật Webcam từ máy tính cá nhân, luồng Stream từ webcam sẽ được cắt ra thành từng frame. Đây sẽ là đầu vào xử lí của mô hình học sâu (CNN).
 - Các Frame này sẽ được đưa qua một Model đã được Train sẵn để dự đoán các Bounding Boxes (các hình chữ nhật chứa bàn tay) chưa các bàn tay. Đầu ra của kết quả dự đoán gồm 1 mảng chứa các tham số của các Bounding Boxes (x y w h)
    + x, y : Tọa độ tâm của hình chữ nhật
    + w : Chiều rộng của hình chữ nhật
    + h : Chiều cao/dài của hình chữ nhật
 - Các Bounding Boxs này sẽ được vẽ lại vào Canvas cùng với ảnh (với khoảng 24fps).

## Xử lý vị trí của bàn tay
 - Dựa vào vị trí tương đối của bàn tay trong frame trước và frame sau để ta xác định được bàn tay đó là đang di chuyển sang bên trái, phải, lên, xuống  

## Điều khiển file Excel
 - Project "My Add In React" là 1 Project Add-in dựa theo cấu trúc được Microsoft phát triển. Nó có thể tạo ra 1 file Excel mà ta có thể trực tiếp can thiệp vào mọi tiện ích của file Excel đó.
 - Mục đích chính của Project ta đang làm là di chuyển cell (selected) bằng bàn tay thông qua webcam
 - Đầu tiên, ta đọc địa chỉ của cell đang được selected. Tính toán các địa chỉ của cell hàng xóm (trên, dưới, trái, phải).
 - Với các thao tác bàn tay nhận được từ Webcam được xử lí ở trên, ta di chuyển cell đang được selected tới các cell hàng xóm

 ## Chạy chương trình
 - Clone Project về máy. git clone https://github.com/beerus266/DoAnThietKe.git
 - Mở cửa sổ CMD, di chuyển đến từng folder. Ví dụ :  cd "react-app-front-end"
 - chạy lệnh: npm start