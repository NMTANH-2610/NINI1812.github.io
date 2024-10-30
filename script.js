const questions = [
    {
        question: "Trong MS Excel, điều nào sau đây là đúng khi nói về ô trong công thức?",
        options: ["A.Các ô phải đứng cạnh nhau trong cùng một hàng", "B.Các ô phải đứng cạnh nhau trong cùng một cột", "C. Các ô phải liền kề", "D. Có thể chọn ô bất kỳ trong bảng tính"],
        correct: 3
    },
    {
        question: "Trong MS Excel, công thức nào sau đây là biểu thị cho tham chiếu hỗn hợp?",
        options: ["A.^A^1", "B.$A#1", "C.@A@1@", "D.$A1"],
        correct: 3
    },
    {
        question: "Trong MS Excel, khi tham chiếu đến bảng tính khác, cần thực hiện điều nào sau đây?",
        options: ["A.	Dùng dấu ngoặc vuông [] và bảng tính khác phải đóng lại", "B.	Dùng dấu ! và cả hai bảng tính đều phải mở", "C.	Dùng dấu + và cả hai bảng tính đều phải mở", "D.	Dùng dấu ngoặc vuông [] và cả hai bảng tính đều phải mở"],
        correct: 3
    },
    {
        question: "Trong MS Excel, điều nào sau đây mô tả cho vùng dữ liệu?",
        options: ["A.	Tất cả đều đúng", "B.	Kích thước có thể thay đổi", "C.	Có thể đặt tên", "D.	Có thể sử dụng trong công thức"],
        correct: 0
    },
    {
        question: "Trong MS Excel, làm cách nào để xem tên của vùng dữ liệu?",
        options: ["A.	Xem trên thanh trạng thái", "B.	Xem hộp thoại tên bên cạnh thanh công thức", "C.	Có thể đặt tên", "D.	Có thể sử dụng trong công thức"],
        correct: 1
    },
    {
        question: "Trong MS Excel, điều nào sau đây là đúng khi nói về việc đặt tên vùng dữ liệu?",
        options: ["A.	Có thể đặt tên vùng trùng với tên trang tính", "B.	Có thể đặt tên vùng trùng với tên của nhiều trang tính", "C.	Tên vùng không được chứa toán tử", "D.	Tên vùng nằm tại thanh công thức"],
        correct: 2
    },
    {
        question: "Trong MS Excel, cú pháp nào sau đây dùng để thêm vùng dữ liệu có tên SALES vào công thức?",
        options: ["A.	&(SALE)", "B.	SALE ", "C.	SUM(SALES)", "D.	TOTAL#(SALES)"],
        correct: 2
    },
    {
        question: "Trong MS Excel, cú pháp nào sau đây dùng để nhân tất cả các ô trong vùng dữ liệu có tên DATA?",
        options: ["A.	PRODUCT(DATA)", "B.	DATA", "C.	x(DATA)", "D.	Multiply(DATA)"],
        correct: 0
    },
    {
        question: "Tab nào trong Excel cung cấp quyền truy cập vào thư viện công thức và hàm số?",
        options: ["A.	File", "B.	View", "C.	Formulas", "D.	Page layout"],
        correct: 2
    },
    {
        question: "Trong MS Excel, đáp án nào dưới đây định nghĩa về công thức được lập trình sẵn để thực hiện một phép tính?",
        options: ["A.	Địa chỉ tham chiếu", "B.	Dấu ngoặc đơn", "C.	Toán từ", "D.	Hàm"],
        correct: 3
    },
    {
        question: "Trong MS Excel, chức năng Function Library trong tab Formulas sắp xếp các hàm theo __________?",
        options: ["A.	Kích cỡ", "B.	Màu sắc", "C.	Vị trí ", "D.	Thể loại"],
        correct: 3
    },
    {
        question: "Trong MS Excel, nếu muốn tính toán khoản lãi suất phải trả cho việc vay thế chấp, bạn sẽ sử dụng hàm nào trong danh mục Financial trên thanh công thức?",
        options: ["A.	Date & time ", "B.	Financial ", "C.	Text ", "D.	Logical "],
        correct: 1
    },
    {
        question: "Trong MS Excel, phần nào trong một công thức có nhiều đối số cụ thể và trả về một giá trị duy nhất dựa vào những đối số?",
        options: ["A.	Hàm", "B.	Vùng", "C.	Ô dữ liệu", "D.	Trang tính"],
        correct: 0
    },
    {
        question: "Trong MS Excel, vị trí của phím Insert Function nằm ở đâu?",
        options: ["A.	Tab Formulas", "B.	Tab View", "C.	Data tab", "D.	Review tab"],
        correct: 0
    },
    {
        question: "Trong MS Excel, hàm nào dưới đây đưa ra giá trị nhỏ nhất trong danh sách các đối số?",
        options: ["A.	MAX", "B.	AVERAGE", "C.	SUM", "D.	MIN"],
        correct: 3
    },
    {
        question: "Trong MS Excel, để xác định điểm kết thúc của một hàm không có đối số, ký tự nào dưới đây vẫn phải được sử dụng?",
        options: ["A.	Dấu chấm câu", "B.	Dấu ngoặc kép", "C.	Dấu ngoặc đơn", "D.	Dấu phẩy"],
        correct: 2
    },
    {
        question: "Trong MS Excel, nếu bạn đã sắp xếp danh sách nhân viên theo từng bộ phận và muốn tính tổng tiền lương theo từng bộ phận, hàm nào phù hợp nhất để sử dụng?",
        options: ["A.	TOTAL", "B.	SUM", "C.	SUBTOTAL", "D.	MAX"],
        correct: 2
    },
    {
        question: "Trong MS Excel, giá trị được trả về bởi hàm nào sẽ tự động cập nhật thay đổi khi trang tính thay đổi?",
        options: ["A.	TODAY", "B.	INSERT", "C.	NOW", "D.	UPDATE"],
        correct: 0
    },
    {
        question: "Trong MS Excel, hàm nào có chức năng mặc định tính tổng các ô dữ liệu liền kề với ô đầu tiên không phải là số và có chứa hàm SUM trong công thức của nó?",
        options: ["A.	AutoSum", "B.	SUM", "C.	SUMIF", "D.	SUMSQ"],
        correct: 0
    },
    {
        question: "Trong MS Excel, hàm COUNT thuộc loại hàm nào?",
        options: ["A.	Hàm văn bản", "B.	Hàm luận lý", "C.	Hàm thống kê", "D.	Hàm cơ bản"],
        correct: 2
    },
    {
        question: "Trong MS Excel, ký tự nào dưới đây bao quanh các đối số trong một hàm?",
        options: ["A.	dấu phẩy", "B.	dấu ngoặc đơn", "C.	dấu ngoặc vuông", "D.	dấu hoa thị"],
        correct: 1
    },
    {
        question: "Trong MS Excel, hàm nào dưới đây đếm số lượng các ô trong vùng chọn có chứa dữ liệu, bao gồm số, ngày tháng, công thức hoặc nhãn văn bản?",
        options: ["A.	COUNTA", "B.	Auto sum", "C.	sum", "D.	total"],
        correct: 0
    },
    {
        question: "Trong MS Excel, hàm nào dưới đây được sử dụng để đếm số lượng ô trong vùng chọn chỉ chứa số, không chứa văn bản?",
        options: ["A.	auto sum", "B.	COUNTA", "C.	total", "D.	COUNT"],
        correct: 3
    },
    {
        question: "Trong MS Excel, tỷ lệ được sử dụng trong hàm PMT đại diện cho đại lượng nào dưới đây?",
        options: ["A.	giá trị đầu tư hiện tại", "B.	tỷ lệ lãi suất cho vay hàng năm", "C.	giá trị của khoản đầu tư tại thời điểm cuối kỳ đầu tư", "D.	tổng các khoản thanh toán"],
        correct: 1
    },
    {
        question: "Trong MS Excel, bạn có nhiệm vụ theo dõi việc kiểm kê hàng hóa của năm thành phố và tính tổng các chi phí cho từng thành phố. Bạn sẽ sử dụng hàm nào dưới đây để tính toán cho nhiệm vụ này?",
        options: ["A.	Min", "B.	Total", "C.	Subtotal", "D.	Max"],
        correct: 2
    },
    {
        question: "Trong MS Excel, bạn cần cộng giá trị ở các ô B7, B30 và B35. Hàm nào sẽ được bạn sử dụng để cộng các giá trị ở các ô không liền kề nhau?",
        options: ["A.	Auto sum", "B.	SUMIF", "C.	SUM", "D.	SUMIFS"],
        correct: 2
    },
    {
        question: "Trong MS Excel, bạn sử dụng công cụ nào dưới đây để thể hiện mối quan hệ giữa công thức và ô tham chiếu của nó?",
        options: ["A.	mũi tên theo dõi tiền lệ", "B.	tên ô dữ liệu", "C.	dấu", "D.	ô dữ liệu trống"],
        correct: 0
    },
    {
        question: "Trong MS Excel, tab nào chứa chức năng theo dõi các ô tiền lệ?",
        options: ["A.	home", "B.	Formulas", "C.	Data", "D.	Page layout"],
        correct: 1
    },
    {
        question: ": Trong MS Excel, tổ hợp phím nào được sử dụng để hiển thị hoặc ẩn công thức trong một trang tính?",
        options: ["A.	Ctrl +  dấu huyền", "B.	Ctrl + Dấu phẩy", "C.	Ctrl + dấu chấm câu", "D.	Ctrl + Dấu gạch ngang"],
        correct: 0
    },
    {
        question: "Trong MS Excel, trang tính của bạn có một cột được thiết kế để theo dõi các dự án đã hoàn thành và bạn muốn đếm số lượng dự án đã hoàn thành. Bạn sẽ sử dụng hàm nào dưới đây để đếm các ô chứa dữ liệu cần tìm?",
        options: ["A.	Count", "B.	Countif", "C.	Counta", "D.	Decout"],
        correct: 2
    },
    {
        question: "Trong MS Excel, khi tạo hóa đơn thanh toán cho khách hàng trong vòng ba mươi ngày, bạn muốn ngăn chặn việc tự động cập nhật sửa đổi. Bạn có thể sử dụng hàm nào dưới đây?",
        options: ["A.	TODAY", "B.	Year", "C.	Day", "D.	Days360"],
        correct: 0
    },
    {
        question: ": Trong MS Excel, khi bạn chèn một ô vào cột dọc, các ô trong cùng cột sẽ dịch chuyển:",
        options: ["A.	qua phải", "B.	qua trái", "C.	xuống dưới", "D.	Lên trên"],
        correct: 1
    },
    {
        question: "Trong MS Excel, khi bạn chèn một ô vào hàng ngang, các ô trong cùng hàng sẽ dịch chuyển:",
        options: ["A.	qua phải", "B.	qua trái", "C.	xuống dưới", "D.	Lên trên"],
        correct: 0
    },
    {
        question: "Trong MS Excel, điều gì xảy ra khi nhấp chuột phải lên một ô?",
        options: ["A.	Một ô bị xóa", "B.	Một ô được chèn", "C.	Ô bị dịch chuyển", "D.	Danh sách lốt tắt hiên thị"],
        correct: 3
    },
    {
        question: "Trong MS Excel, thao tác nào dưới đây có nghĩa là sắp xếp và sắp hàng các ô ngay thẳng với nhau?",
        options: ["A.	Canh dòng", "B.	Làm phẳng", "C.	Làm thẳng", "D.	Sắp xếp"],
        correct: 0
    },
    {
        question: "Trong MS Excel, tại vị trí nào bạn có thể truy cập vào lệnh Alignment?",
        options: ["A.	hộp thoại Format Cells", "B.	tab Data", "C.	Tab view", "D.	Tab review"],
        correct: 0
    },
    {
        question: "Trong MS Excel, khi bạn nhập số, vị trí mặc định là:",
        options: ["A.	canh dòng bên phải ", "B.	canh dòng bên  trái", "C.	ở ngay trung tâm", "D.	canh đều"],
        correct: 0
    },
    {
        question: "Trong MS Excel, lệnh nào dùng để căn chỉnh văn bản giữa các ô?",
        options: ["A.	Top", "C.	Center", "B.	Bottom", "D.	Justify"],
        correct: 3
    },
    {
        question: ": Trong MS Excel, lệnh nào dùng để căn chỉnh văn bản từ trên xuống dưới trong một ô?",
        options: ["A.	distributed", "B.	Center", "C.	Justify", "D.	Top"],
        correct: 0
    },
    {
        question: "Trên tab Home trong nhóm lệnh căn chỉnh, nút nào được sử dụng để di chuyển văn bản sang phải?",
        options: ["A.	top align", "B.	bottom align", "C.	Increase Indent", "D.	Decrease indent"],
        correct: 2
    },
    {
        question: "Trong MS Excel, tổ hợp phím nào được dùng để tăng khoảng cách thụt đầu các dòng?",
        options: ["A.	Ctrl + H +4", "B.	Alt + H + 6", "C.	Ctrl + H +6", "D.	Shift + H +4"],
        correct: 2
    },
    {
        question: "Trong MS Excel, bạn chọn mục nào trên tab File để thay đổi font chữ mặc định dùng trong tất cả các bảng tính?",
        options: ["A.	Info", "B.	Save", "C.	Share", "D.	Options"],
        correct: 3
    },
    {
        question: "Trong MS Excel, thanh công cụ nào là công cụ định dạng, được hiển thị trên hoặc dưới danh sách tắt khi bạn nhấp chuột phải vào một ô?",
        options: ["A.	Mini ", "B.	Top", "C.	New (Mới)", "D.	Fast"],
        correct: 0
    },
    {
        question: "Trong MS Excel, để làm nổi bật hoặc tạo nền màu cho các ô trong một trang tính, bạn sẽ chọn công cụ nào?",
        options: ["A.	Bold (Đậm)", "B.	Font Color (Màu chữ)", "C.	Fill Color (Tô màu)", "D.	Font"],
        correct: 2
    },
    {
        question: "Trong MS Excel, nhóm công cụ nào trên tab Home giúp định dạng số cho các ô được chọn?",
        options: ["A.	Clipboard (Bảng kẹp)", "B.	Font (Phông chữ)", "C.	Alignment", "D.	Number (Số)"],
        correct: 3
    },
    {
        question: "Trong MS Excel, định dạng số nào sau đây hiển thị mặc định với ký hiệu tiền tệ và có hai chữ số thập phân?",
        options: ["A.	Number (Số)", "B.	Percentage (Phần trăm)", "C.	Accounting (Kế toán)", "D.	Currency (Tiền tệ)"],
        correct: 3
    },
    {
        question: "Trong MS Excel, các định dạng số nào sau đây có sẵn trong danh sách Number Format?",
        options: ["A.	Tất cả đúng", "B.	General (Tổng quan)", "C.	Accounting (Kế toán)", "D.	Fraction (Phân số)"],
        correct: 0
    },
    {
        question: "Trong MS Excel, con trỏ chuột sẽ trông như thế nào khi Format Painter đang hoạt động?",
        options: ["A.	Dấu cộng màu trắng", "B.	Dấu cộng màu trắng và cây chổi sơn", "C.	Chỉ có cây chổi sơn", "D.	Dấu cộng màu đen và cây chổi sơn"],
        correct: 1
    },
    {
        question: "Để cung cấp các tùy chọn đặc biệt khi dán các ô, Excel sẽ hiển thị nút tùy chọn Paste trên trang tính. Đúng không?",
        options: ["A.	Paste (Dán)", "B.	Copy (Sao chép)", "C.	Cut (Cắt)", "D.	Insert (Chèn)"],
        correct: 0
    },
    {
        question: "Trong MS Excel, để tăng cường thao tác với trang tính, có thể sử dụng tập hợp các cài đặt sẵn của ô nào sau đây?",
        options: ["A.	Công thức", "B.	Kiểu", "C.	Ngày tháng năm ", "D.	Ký hiệu"],
        correct: 1
    },
    {
        question: "Trong MS Excel, câu nào là đúng khi áp dụng kiểu định dạng cho các ô khác nhau trong trang tính?",
        options: ["A.	Một bảng tính chỉ có thể có một kiểu duy nhất", "B.	Có thể kết hợp các kiểu định dạng khác nhau theo ý muốn", "C.	Chỉ có thể kết hợp tối đa hai kiểu định dạng", "D.	Không thể sử dụng lệnh Undo để đổi kiểu định dạng"],
        correct: 1
    },
    {
        question: ": Trong MS Excel, để tạo kiểu định dạng riêng, bạn mở danh sách Cell Styles và chọn tùy chọn nào sau đây?",
        options: ["A.	New Cell Style (Kiểu ô mới)", "B.	My Cell Style (Phong cách ô của tôi)", "C.	Custom Cell Style (Kiểu ô tùy chỉnh)", "D.	Rename Cell Style (Đổi tên kiểu ô)"],
        correct: 0
    },
    {
        question: "Trong MS Excel, lệnh nào trên danh sách tắt có thể được sử dụng để xóa một siêu liên kết từ ô đã chọn?",
        options: ["A.	Undo hyperlink", "B.	Earse hyperlink", "C.	Remove Hyperlink", "D.	Delete Hyperlink"],
        correct: 2
    },
    {
        question: "Trong MS Excel, định dạng dữ liệu có điều kiện nằm trong nhóm nào trên tab Home?",
        options: ["A.	Clipboard (Bảng kẹp)", "B.	Font (Phông chữ)", "C.	Styles (Phong cách)", "D.	Alignment (Căn chỉnh)"],
        correct: 2
    },
    {
        question: "Trong MS Excel, tùy chọn nào trên danh sách định dạng có điều kiện cung cấp quyền truy cập vào Conditional Formatting Rules Manager?",
        options: ["A.	Rename Rule (Đổi tên quy tắc)", "B.	Create Rule (Tạo quy tắc)", "C.	Manage Rules (Quản lý quy tắc)", "D.	Write rules"],
        correct: 2
    },
    {
        question: "Trong MS Excel, hàng tiêu đề được xác định bởi điều gì dưới đây?",
        options: ["A.	Tổ hợp phím ", "B.	Biểu tượng ", "C.	Ký tự ", "D.	Số "],
        correct: 3
    },
    {
        question: "Trong MS Excel, cột tiêu đề được xác định bởi điều gì dưới đây?",
        options: ["A.	Tổ hợp phím ", "B.	Biểu tượng ", "C.	Ký tự ", "D.	Số "],
        correct: 2
    },
    {
        question: "Trong MS Excel, để xóa một hàng, chọn hàng tiêu đề, sau đó nhấp vào mũi tên trong nhóm Cells trên tab Home và chọn lệnh nào?",
        options: ["A.	Erase", "B.	Delete", "C.	Cut", "D.	Remove	"],
        correct: 1
    },
    {
        question: "Để Excel có thể tự động điều chỉnh độ rộng của ô dữ liệu cho phù hợp với kích thước của nội dung, bạn thực hiện thao tác gì ngay tại đường ranh giới giữa hàng và cột?",
        options: ["A.	Nhấp chuột", "B.	Kéo", "C.	Nhấp chuột ba lần", "D.	Nhấp đôi chuột	"],
        correct: 3
    },
    {
        question: "Trong MS Excel, khi muốn ẩn các hàng hoặc cột không mong muốn trong trang tính, bạn có thể sử dụng tùy chọn nào từ danh sách tắt?",
        options: ["A.	Option ", "B.	Clesr", "C.	Hide", "D.	Delete	"],
        correct: 2
    },
    {
        question: "Trong MS Excel, màu nào biểu thị cho nội dung bị ẩn trong một trang tính?",
        options: ["A.	Đỏ", "B.	Xanh lá", "C.	Vàng", "D.	Xanh biển"],
        correct: 1
    },
    {
        question: "Trong MS Excel, lệnh Transpose được tìm thấy trong hộp thoại Paste Special thuộc nhóm Clipboard trên tab nào?",
        options: ["A.	Insert ", "B.	View", "C.	Review", "D.	Home"],
        correct: 3
    },
    {
        question: "Trong MS Excel Để định dạng nhanh trang tính, ta chọn bất cứ ô dữ liệu nào đang hoạt động trong trang tính và chọn một biểu mẫu đã thiết kế sẵn ở trong tab nào dưới đây?",
        options: ["A.	Page Layout", "B.	Home", "C.	Review", "D.	View"],
        correct: 0
    },
    {
        question: ": Trong MS Excel Những điều nào dưới đây là đúng khi nói về cách sử dụng các biểu mẫu đã hiệu chỉnh?",
        options: ["A.	Chúng chỉ thích hợp một lần duy nhất", "B.	Chúng chỉ thích hợp với ứng dụng mà chúng được tạo", "C.	Chúng thích hợp cho tất cả các loại tài liệu của Office 2013", "D.	Chúng chỉ thích hợp trong trang tính của Excel"],
        correct: 2
    },
    {
        question: "Trong MS Excel Sử dụng SmartArt hoặc các đồ họa tương tự. trong nhóm các biểu mẫu để nhấn mạnh nội dung của biểu đồ, hình dạng,",
        options: ["A.	Font", "B.	Currency", "C.	Symbols ", "D.	Effects"],
        correct: 3
    },
    {
        question: "Trong MS Excel Để bỏ một hình ảnh nền, nhấp chuột vào chức năng Delete Background trong nhóm Page Setup thuộc tab nào dưới đây?",
        options: ["A.	Home", "B.	Insert", "C.	Page layout", "D.	Data	"],
        correct: 2
    },
    {
        question: "Trong MS Excel Nhấp chuột vào trong nhóm Sheet Option để mở hộp thoại Page Setup.",
        options: ["A.	Nút Open", "B.	Dữ liệu", "C.	Trình khởi chạy hộp thoại", "D.	Hàng	"],
        correct: 2
    },
    {
        question: "Trong MS Excel Để gõ thêm một tiêu đề đầu trang hoặc cuối trang vào trang tính, nhấp chuột vào nút Header & Footer trên tab Insert tại nhóm nào dưới đây?",
        options: ["A.	Table", "B.	Apps", "C.	Text", "D.	Reports"],
        correct: 2
    },
    {
        question: "Trong MS Excel Tab nào dưới đây dùng hiển thị thêm chức năng định dạng Header & Footer?",
        options: ["A.	Design	", "B.	Data	", "C.	Home", "D.	Review"],
        correct: 0
    },
    {
        question: "Trong MS Excel Để tiết kiệm thời gian, có thể sử dụng một trong các tiêu đề đầu trang và chân trang được định sẵn trên tab nào dưới đây?",
        options: ["A.	Data ", "B.	Home", "C.	Design ", "D.	Reviews"],
        correct: 0
    },
    {
        question: "Trong MS Excel Để nhìn thấy bao nhiêu tài liệu được in chia giữa các trang, ta chọn Page Break Preview tại tab nào dưới đây?",
        options: ["A.	View", "B.	Home", "C.	Data", "D.	Design"],
        correct: 0
    },
    {
        question: "Trong MS Excel Chức năng nào dưới đây được sử dụng để thu hẹp hoặc kéo dài tài liệu để in trên một trang duy nhất?",
        options: ["A.	Pasting", "B.	Cutting", "C.	Zooming", "D.	Scaling	"],
        correct: 0
    },
    {
        question: "Trong MS Excel Nhóm nào dưới đây trong tab Page Layout giúp thay đổi được phạm vi trang tính cần in",
        options: ["A.	Arrange", "B.	Page Setup", "C.	Sheet Options", "D.	Theme"],
        correct: 1
    },
    {
        question: "Trong MS Excel Điều gì dưới đây sẽ xảy ra với màu tab trang tính khi sao chép trang tính?",
        options: ["A.	Tab trang tính chuyển sang màu trắng", "B.	Tab trang tính chuyển sang màu xám", "C.	Tab trang tính giữ nguyên màu sắc", "D.	Tab trang tính chuyển sang không  màu"],
        correct: 2
    },
    {
        question: " Trong MS Excel Để ẩn một vài trang tính cùng lúc, đầu tiên phải muốn ẩn, sau đó nhấp chuột phải và chọn",
        options: ["A.	ấn Ctrl, hide", "B.	giữ nút tab, hide", "C.	chọn công cụ hightlight, protect sheet", "D.	ấn tổ hợp phím Ctrl + H"],
        correct: 0
    },
    {
        question: "Trong MS Excel, khi quản lý trang tính, cần phải lưu ý những điểm nào dưới đây?",
        options: [
            "A. Thêm những trang tính mới vào cùng một bảng tính.",
            "B. Sử dụng ô dữ liệu tham chiếu trên cùng bảng tính.",
            "C. Hạn chế số lượng của trang tính được sử dụng trong một dự án.",
            "D. Sử dụng một trang tính cho mỗi dự án và không sử dụng ô tham chiếu."
        ],
        correct: 0 // Thay đổi theo đáp án đúng
    },
    {
        question: "Trong MS Excel, làm thế nào có thể xem nhiều trang tính trong cùng một bảng tính cùng lúc?",
        options: [
            "A. Không thể xem nhiều trang tính trong cùng một bảng tính cùng lúc.",
            "B. Đi đến tab View và chọn New Window.",
            "C. Đi đến tab Page Layout và chọn Orientation.",
            "D. Đi tới tab Insert và chọn Power View."
        ],
        correct: 1 // Đáp án đúng là "B. Đi đến tab View và chọn New Window."
    },
    {
        question: "Trong MS Excel, làm thế nào có thể xem nhiều trang tính của một bảng tính sắp xếp cạnh nhau trong cùng một cửa sổ?",
        options: [
            "A. Di chuyển những trang tính ấy sang những bảng tính riêng biệt, sau đó xếp các bảng tính theo thứ tự.",
            "B. Đi đến tab View, chọn Arrange All và chọn Vertical.",
            "C. Không thể xem nhiều trang tính của một bảng tính sắp xếp cạnh nhau trong cùng một cửa sổ.",
            "D. Đi đến tab View, chọn Arrange All và chọn Tiled."
        ],
        correct: 3 // Đáp án đúng là "D. Đi đến tab View, chọn Arrange All và chọn Tiled."
    },
    {
        question: "Trong MS Excel, có thể tìm thấy chức năng Zoom ở tab Menu nào?",
        options: [
            "A. File",
            "B. Page Layout",
            "C. Review",
            "D. View"
        ],
        correct: 3 // Đáp án đúng là "D. View"
    },
    {
        question: "Trong MS Excel, để phóng to một phần của trang tính, chọn chức năng nào dưới đây trên tab Zoom?",
        options: [
            "A. Zoom",
            "B. 100%",
            "C. Zoom to selection",
            "D. Custom View"
        ],
        correct: 2 // Đáp án đúng là "C. Zoom to selection"
    },
    {
        question: "Trong MS Excel, nếu chọn một vùng dữ liệu để thực hiện lệnh Find and Replace, điều gì dưới đây sẽ xảy ra nếu chọn Replace all?",
        options: [
            "A. Chỉ những ô dữ liệu trong vùng chọn được thay thế.",
            "B. Không có ô dữ liệu nào được thay thế; cần phải chọn toàn bộ trang tính.",
            "C. Tất cả các ô dữ liệu trong bảng tính đáp ứng đúng với các tiêu chí sẽ được thay thế.",
            "D. Lệnh sẽ chỉ ảnh hưởng với giới hạn 100 ô dữ liệu."
        ],
        correct: 0 // Đáp án đúng là "A. Chỉ những ô dữ liệu trong vùng chọn được thay thế."
    },
    {
        question: "Trong MS Excel, nếu một phép tính tham chiếu đến một vùng tên, đáp án nào dưới đây phải được chọn để tìm kiếm?",
        options: [
            "A. Tìm kiếm chính xác tất cả nội dung trong ô dữ liệu.",
            "B. Tìm kiếm giá trị các ô",
            "C. Tìm kiếm chú thích",
            "D. Tìm kiếm công thức"
        ],
        correct: 3 // Đáp án đúng là "D. Tìm kiếm công thức"
    },
    {
        question: "Trong MS Excel, có thể tìm thấy lệnh Go To ở tab và vùng nào?",
        options: [
            "A. View > Find & Select",
            "B. Data > Find & Select",
            "C. Home > Find & Select",
            "D. Insert > Find & Select"
        ],
        correct: 2 // Đáp án đúng là "C. Home > Find & Select"
    },
    {
        question: "Trong MS Excel, phân tích dữ liệu nghĩa là gì?",
        options: [
            "A. Xuất dữ liệu sang bảng tính khác",
            "B. Hiển thị tất cả thông tin",
            "C. Thêm dữ liệu đặc biệt",
            "D. Giải thích dữ liệu bằng bảng tính"
        ],
        correct: 3 // Đáp án đúng là "D. Giải thích dữ liệu bằng bảng tính"
    },
    {
        question: "Trong MS Excel, ký tự phân cách các dữ liệu với nhau gọi là gì?",
        options: [
            "A. Điều kiện",
            "B. Dấu tách",
            "C. Dấu phân số",
            "D. Macro"
        ],
        correct: 1 // Đáp án đúng là "B. Dấu tách"
    },
    {
        question: "Câu 86: Khi dữ liệu được nhập vào một tập tin Excel, định dạng dữ liệu mặc định là gì?",
        options: [
            "A. General",
            "B. Text",
            "C. Date",
            "D. Numbers"
        ],
        correct: 0 // Đáp án đúng là "A. General"
    },
    {
        question: "Câu 87: Trong MS Excel, các quy tắc xác nhận liên quan đến dữ liệu nào?",
        options: [
            "A. Chỉ dữ liệu từ các trang tính Excel khác",
            "B. Dữ liệu mới khi nhập vào trang tính",
            "C. Tất cả các dữ liệu nhập vào và dữ liệu hiện có trong trang tính",
            "D. Dữ liệu đã tồn tại trong trang tính hiện tại"
        ],
        correct: 2 // Đáp án đúng là "C. Tất cả các dữ liệu nhập vào và dữ liệu hiện có trong trang tính"
    },
    {
        question: "Câu 88: Trong MS Excel, sau khi dữ liệu được lọc, điều gì xảy ra với các dòng không có liên quan?",
        options: [
            "A. Chúng bị xóa",
            "B. Chúng được làm nổi bật",
            "C. Chúng được sao chép",
            "D. Chúng được ẩn đi"
        ],
        correct: 3 // Đáp án đúng là "D. Chúng được ẩn đi"
    },
    {
        question: "Câu 89: Auto Filter sử dụng cái gì dưới đây để làm điều kiện lọc dữ liệu?",
        options: [
            "A. Dòng đầu tiên",
            "B. Dòng tiêu đề",
            "C. Dòng cuối cùng",
            "D. Dòng tô màu"
        ],
        correct: 1 // Đáp án đúng là "B. Dòng tiêu đề"
    },
    {
        question: "Câu 90: Trong MS Excel, nhóm và tổ chức dữ liệu trong một bảng tính được thực hiện lúc nào?",
        options: [
            "A. Trước khi lọc dữ liệu",
            "B. Ban đầu",
            "C. Dòng cuối cùng",
            "D. Dòng tô màu"
        ],
        correct: 0 // Đáp án đúng là "A. Trước khi lọc dữ liệu"
    },
    {
        question: "Câu 91: Trong MS Excel, tab nào chứa các lệnh Subtotal?",
        options: [
            "A. Page Layout",
            "B. Insert",
            "C. Formula",
            "D. Data"
        ],
        correct: 3 // Đáp án đúng là "D. Data"
    },
    {
        question: "Câu 92: Trong MS Excel, khi sử dụng các định dạng tại hộp thoại Table, có thể thiết lập bảng theo định dạng nào ngoài việc tạo tiêu đề?",
        options: [
            "A. Bộ lọc",
            "B. Màu",
            "C. Kích thước",
            "D. Phông chữ"
        ],
        correct: 0 // Đáp án đúng là "A. Bộ lọc"
    },
    {
        question: "Câu 93: Trong MS Excel, tab nào xuất hiện khi một kiểu định dạng được áp dụng cho bảng?",
        options: [
            "A. Design",
            "B. Data",
            "C. Style",
            "D. Page Layout"
        ],
        correct: 0 // Đáp án đúng là "A. Design"
    },
    {
        question: "Câu 94: Trong MS Excel, kí tự nào sau đây dùng trong công thức có nghĩa “chỉ nhận giá trị của cột ở trong cùng một hàng”?",
        options: [
            "A. @",
            "B. $",
            "C. &",
            "D. *"
        ],
        correct: 1 // Đáp án đúng là "B. $"
    },
    {
        question: "Câu 95: Trong MS Excel, khi dùng chức năng Remove Duplicates trong việc so sánh các trường dữ liệu được lựa chọn trong hộp thoại, hàng nào sẽ được loại bỏ?",
        options: [
            "A. Hàng ở trên sau khi trùng lặp nội dung",
            "B. Hàng ở dưới để không trùng lặp nội dung",
            "C. Hàng ở trên để không trùng lặp nội dung",
            "D. Hàng ở dưới sau khi trùng lặp nội dung"
        ],
        correct: 0 // Đáp án đúng là "A. Hàng ở trên sau khi trùng lặp nội dung"
    },
    {
        question: "Câu 96: Trong MS Excel, Remove Duplicates sẽ loại bỏ tất cả các thông tin trùng lặp trong một hàng nếu trường dữ liệu được lựa chọn để so sánh.",
        options: [
            "A. Tất cả",
            "B. Ba",
            "C. Năm",
            "D. Hai"
        ],
        correct: 0 // Đáp án đúng là "A. Tất cả"
    },
    {
        question: "Câu 97: Trong MS Excel, để hiển thị tab cần thiết để thực thi macros, nhấp chuột vào tab Files, nhấp Options, nhấp Customized Ribbon, kiểm tra hộp và nhấp chuột vào OK.",
        options: [
            "A. Dữ liệu",
            "B. Phần bổ sung",
            "C. Nhà phát triển",
            "D. Công thức"
        ],
        correct: 2 // Đáp án đúng là "C. Nhà phát triển"
    },
    {
        question: "Câu 98: Trong MS Excel, Macro Security sẽ cảnh báo bất cứ khi nào bạn mở một bảng tính có chứa macro nếu Trust Center được thiết lập tính năng nào?",
        options: [
            "A. Vô hiệu hoá tất cả các macro mà không cần thông báo",
            "B. Vô hiệu hoá tất cả các macro có thông báo",
            "C. Vô hiệu hoá tất cả các macro, trừ macro có chữ ký số",
            "D. Cho phép tất cả các macro"
        ],
        correct: 2 // Đáp án đúng là "C. Vô hiệu hoá tất cả các macro, trừ macro có chữ ký số"
    },
    {
        question: "Câu 99: Trong MS Excel, hàm nào thêm ô vào trong dãy được xác định bởi một điều kiện cho sẵn?",
        options: [
            "A. SUMIF",
            "B. COUNTIF",
            "C. AVERAGEIF",
            "D. VLOOKUP"
        ],
        correct: 0 // Đáp án đúng là "A. SUMIF"
    },
    {
        question: "Câu 100: Trong MS Excel, tab nào chứa chức năng Function Library?",
        options: [
            "A. Home",
            "B. Insert",
            "C. Formulas",
            "D. Data"
        ],
        correct: 2 // Đáp án đúng là "C. Formulas"
    },
    {
        question: "Câu 101: Trong MS Excel, thuật ngữ dành cho các giá trị mà một hàm sử dụng để thực hiện thao tác hoặc tính toán là gì?",
        options: [
            "A. Đối số",
            "B. Điều kiện",
            "C. Vùng dữ liệu",
            "D. Hộp thoại"
        ],
        correct: 0 // Đáp án đúng là "A. Đối số"
    },
    {
        question: "Câu 102: Trong MS Excel, định nghĩa một quy tắc và có thể là một số, văn bản hay biểu thức?",
        options: [
            "A. Hộp thoại",
            "B. Đối số",
            "C. Điều kiện",
            "D. Vùng dữ liệu"
        ],
        correct: 2 // Đáp án đúng là "C. Điều kiện"
    },
    {
        question: "Câu 103: Trong MS Excel, một cửa hàng ngũ kim cần tính tổng lượng búa bán trong tháng. Ngày ở cột A, tên mặt hàng ở cột B, số lượng mặt hàng bán ra ở cột C, và phạm vi băng là A2:C50. Công thức nào dưới đây là đúng?",
        options: [
            "A. SUMIF(B2:B50, \"<>hammer\", A2:C50)",
            "B. SUMIF(A2:C50, \"hammer\", C2:C50)",
            "C. SUMIF(B2:B50, \"hammer\", C2:C50)",
            "D. SUMIF(C2:C50, \"hammer\", B2:B50)"
        ],
        correct: 2 // Đáp án đúng là "C. SUMIF(B2:B50, \"hammer\", C2:C50)"
    },
    {
        question: "Câu 104: Trong MS Excel, chữ cái nào trên button Insert Function hiển thị hộp Function Arguments?",
        options: [
            "A. If",
            "B. Fun",
            "C. InF",
            "D. fx"
        ],
        correct: 3 // Đáp án đúng là "D. fx"
    },
    {
        question: "Câu 105: Trong MS Excel, một trang tính gồm 50 dòng liệt kê hồ sơ học viên. Cột A là giới tính, cột B là màu mắt, và cột C là tuổi. Bạn cần xác định số lượng người tham gia là nữ giới. Vùng dữ liệu trong hàm COUNTIF là gì?",
        options: [
            "A. A2:B50",
            "B. A2:A50",
            "C. B2:B50",
            "D. A2:C50"
        ],
        correct: 1 // Đáp án đúng là "B. A2:A50"
    },
    {
        question: "Câu 106: Trong MS Excel, hàm nào phạm vi đáp ứng một tiêu chí đã cho?",
        options: [
            "A. SUMIF",
            "B. COUNTIF",
            "C. AVERAGEIF",
            "D. VLOOKUP"
        ],
        correct: 1 // Đáp án đúng là "B. COUNTIF"
    },
    {
        question: "Câu 107: Trong MS Excel, Function Library bao gồm các hàm nào cho COUNTIF và AVERAGE?",
        options: [
            "A. Financial",
            "B. Logical",
            "C. Math & trig",
            "D. More functions"
        ],
        correct: 3 // Đáp án đúng là "D. More functions"
    },
    {
        question: "Câu 108: Trong MS Excel, dùng công thức nào để tính giá trị trung bình hàng bàn chải hiện đang tồn kho trong bộ phận dụng cụ khi cột A là bộ phận, cột B là tên dụng cụ và cột C là giá?",
        options: [
            "A. AVERAGEIFS(C2:C25, A2:A25, \"Tool \", B2:B25, \"Brush\")",
            "B. AVERAGEIFS(A2:C25, A2:A25, \"Tool\", B2:B25, \"Brush\")",
            "C. AVERAGEIFS(A2:C25, A2:A25, \"Brush\", B2:B25, \"Tool\")",
            "D. AVERAGEIFS(A2:A25, B2:B25, \"Brush\", C2:C25, \"Tool\")"
        ],
        correct: 0 // Đáp án đúng là "A. AVERAGEIFS(C2:C25, A2:A25, \"Tool \", B2:B25, \"Brush\")"
    },
    {
        question: "Câu 109: Trong MS Excel, trong hàm HLOOKUP, khi TRUE được sử dụng trong Range_lookup, điều gì sẽ xảy ra?",
        options: [
            "A. Tìm kết quả chính xác ở dòng trên.",
            "B. Tìm kết quả gần đúng ở dòng trên.",
            "C. Tìm kết quả gần đúng trong cột đầu tiên.",
            "D. Tìm kết quả chính xác trong cột đầu tiên."
        ],
        correct: 1 // Đáp án đúng là "B. Tìm kết quả gần đúng ở dòng trên."
    },
    {
        question: "Câu 110: Trong MS Excel, hàm nào in hoa chữ cái đầu tiên của một chuỗi văn bản và chuyển đổi các chữ cái khác thành chữ thường?",
        options: [
            "A. Proper",
            "B. Upter",
            "C. Lower",
            "D. Caplock"
        ],
        correct: 0 // Đáp án đúng là "A. Proper"
    },
    {
        question: "Câu 111: Trong Excel, nếu ô K1 có “STUDY HARD”, công thức sử dụng để “STUDY HARD” xuất hiện trong ô K2 là gì?",
        options: [
            "A. UPPER(K1)",
            "B. LOWER(K1)",
            "C. HIGHER(K1)",
            "D. CAPLOCK(K1)"
        ],
        correct: 0 // Đáp án đúng là "A. UPPER(K1)"
    },
    {
        question: "Câu 112: Trong Excel, nếu ô C1 có “STUDY HARD”, công thức sử dụng để “STUDY HARD” xuất hiện trong ô C2 là gì?",
        options: [
            "A. UPPER(C1)",
            "B. PROPER(C1)",
            "C. HIGHER(C1)",
            "D. LOWER(C1)"
        ],
        correct: 0 // Đáp án đúng là "A. UPPER(C1)"
    },
    {
        question: "Câu 113: Trong MS Excel, sau khi sử dụng các hàm cần thiết chuyển đổi văn bản từ cột A sang cột B, bước nào sẽ cho phép tiếp tục thay thế các kết quả mà không cần công thức trong một cột?",
        options: [
            "A. Sao chép kết quả từ cột B và dán ở cột A.",
            "B. Chèn một cột C mới, sao chép các kết quả từ cột B, dán vào cột C, và xóa cột A và B.",
            "C. Sao chép kết quả từ cột B, sử dụng tùy chọn Paste Special để dán công thức vào cột B, và xóa cột A.",
            "D. Sao chép kết quả từ cột B, sử dụng tùy chọn Paste Special để dán giá trị vào cột B, và xóa cột A."
        ],
        correct: 3 // Đáp án đúng là "D. Sao chép kết quả từ cột B, sử dụng tùy chọn Paste Special để dán giá trị vào cột B, và xóa cột A."
    },
    {
        question: "Câu 114: Trong MS Excel, CONCATENATE là gì?",
        options: [
            "A. Tách văn bản thành hai cột.",
            "B. Chỉ kết hợp số với nhau.",
            "C. Thêm một kí tự ở giữa một chuỗi văn bản.",
            "D. Kết hợp các chuỗi văn bản lại với nhau."
        ],
        correct: 3 // Đáp án đúng là "D. Kết hợp các chuỗi văn bản lại với nhau."
    },
    {
        question: "Câu 115: Trong MS Excel, tổ hợp phím dùng để hiển thị tất cả các công thức trong một trang tính là gì?",
        options: [
            "A. Alt + Ctrl",
            "B. CTRL + `",
            "C. Shift +",
            "D. Ctrl + "
        ],
        correct: 1 // Đáp án đúng là "B. CTRL + `"
    },
    {
        question: "Câu 116: Trong MS Excel, các _________ của hàm IF là Logical_Test, [Value_If_True], và [Value_If_False].",
        options: [
            "A. Đối số",
            "B. Giá trị",
            "C. Biến",
            "D. Đối tượng"
        ],
        correct: 0 // Đáp án đúng là "A. Đối số"
    },
    {
        question: "Câu 117: Trong MS Excel, hàm nào tạo ra một số ngẫu nhiên mà không cần điều kiện?",
        options: [
            "A. RANDBETWEEN",
            "B. NUMBER",
            "C. RAND",
            "D. RNUM"
        ],
        correct: 2 // Đáp án đúng là "C. RAND"
    },
    {
        question: "Câu 118: Trong MS Excel, hàm nào cho phép tạo một số ngẫu nhiên trong vùng xác định?",
        options: [
            "A. RANDBETWEEN",
            "B. SELECT NUMBER",
            "C. RAND",
            "D. RNUMBETWEEN"
        ],
        correct: 0 // Đáp án đúng là "A. RANDBETWEEN"
    },
    {
        question: "Câu 119: Trong MS Excel, ngoài việc xem, bạn còn có thể cho phép người dùng làm gì khi mời mọi người chia sẻ một bảng tính?",
        options: [
            "A. Xóa trang tính",
            "B. Tạo",
            "C. Chia sẻ với những người khác",
            "D. Chỉnh sửa"
        ],
        correct: 3 // Đáp án đúng là "D. Chỉnh sửa"
    },
    {
        question: "Câu 120: Trong MS Excel, có thể tìm thấy Show/Hide Comments, chức năng cho phép hiển thị tất cả bình luận trên một trang tính tại tab nào?",
        options: [
            "A. Insert",
            "B. Formulas",
            "C. Data",
            "D. Review"
        ],
        correct: 3 // Đáp án đúng là "D. Review"
    },
    {
        question: "Câu 121: Trong MS Excel, Document Inspector nằm ở đâu?",
        options: [
            "A. FILE > Info > Check for Issues > Inspect",
            "B. Formulas > Trace Precedents",
            "C. Review > Info > Check for Issues > Inspect",
            "D. Formulas > Error > Checking"
        ],
        correct: 0 // Đáp án đúng là "A. FILE > Info > Check for Issues > Inspect"
    },
    {
        question: "Câu 122: Trong MS Excel, để Mark as Final, nhấp chuột vào gì?",
        options: [
            "A. Save > Review > Mark as Final",
            "B. Save > File > Protect Workbook Button > Mark as Final",
            "C. Save > Home > Protect Workbook > Mark as Final",
            "D. Save > View > Mark as Final"
        ],
        correct: 1 // Đáp án đúng là "B. Save > File > Protect Workbook Button > Mark as Final"
    },
    {
        question: "Câu 123: Trong MS Excel, tab nào được sử dụng để lập biểu đồ?",
        options: [
            "A. File",
            "B. Insert",
            "C. Page Layout",
            "D. DATA"
        ],
        correct: 1 // Đáp án đúng là "B. Insert"
    },
    {
        question: "Câu 124: Trong MS Excel, làm thế nào để di chuyển một biểu đồ sang trang tính khác?",
        options: [
            "A. Vào tab Page Layout, nhấp vào nút Move Chart, bấm New sheet, và đặt tên.",
            "B. Kéo biểu đồ vào trang tính mới.",
            "C. Đi đến tab Insert, nhấp New Chart, và đặt tên.",
            "D. Đi đến tab Design, nhấp vào nút Move Chart, bấm New sheet, và đặt tên."
        ],
        correct: 0 // Đáp án đúng là "A. Vào tab Page Layout, nhấp vào nút Move Chart, bấm New sheet, và đặt tên."
    },
    {
        question: "Câu 125: Trong MS Excel, loại biểu đồ nào dùng để đánh giá xu hướng?",
        options: [
            "A. cột",
            "B. đường",
            "C. radar",
            "D. hình tròn"
        ],
        correct: 1 // Đáp án đúng là "B. đường"
    },
    {
        question: "Câu 126: Trong MS Excel, biểu đồ nào dùng để so sánh dữ liệu giữa các hạng mục qua từng thời kỳ?",
        options: [
            "A. cột",
            "B. đường",
            "C. hình tròn",
            "D. tán xạ"
        ],
        correct: 0 // Đáp án đúng là "A. cột"
    },
    {
        question: "Câu 127: MS Excel, cái gì đại diện cho dữ liệu hoặc giá trị của ô trong trang tính?",
        options: [
            "A. biểu tượng",
            "B. biểu đồ",
            "C. điểm dữ liệu",
            "D. chuỗi dữ liệu"
        ],
        correct: 2 // Đáp án đúng là "C. điểm dữ liệu"
    },
    {
        question: "Câu 128: Trong MS Excel, làm thế nào để chỉnh sửa chuỗi dữ liệu trong biểu đồ từ Chart Tools?",
        options: [
            "A. Đến tab Format và chọn Format Selection",
            "B. Đến tab Design và chọn Add Chart Elements.",
            "C. Đến tab Design và chọn Change Chart Type.",
            "D. Đến tab Design và chọn Selection Data"
        ],
        correct: 3 // Đáp án đúng là "D. Đến tab Design và chọn Selection Data"
    },
    {
        question: "Câu 129: Trong MS Excel, làm thế nào để chỉnh sửa các thành phần bên trong biểu đồ tại Chart Tools?",
        options: [
            "A. Đến tab Format và chọn Format Selection",
            "B. Đến tab Design và chọn Add Chart Elements.",
            "C. Đến tab Design và chọn Select Data",
            "D. Đến tab Design và chọn Change Chart Type."
        ],
        correct: 1 // Đáp án đúng là "B. Đến tab Design và chọn Add Chart Elements."
    },
    {
        question: "Câu 130: Trong MS Excel, điều gì sẽ xảy ra khi một nhóm dữ liệu mới được thêm vào biểu đồ bằng cách nhấp chọn Select Data từ Công cụ biểu đồ?",
        options: [
            "A. Xóa và tạo lại biểu đồ",
            "B. Biểu đồ tự động cập nhật",
            "C. Chọn lại vùng dữ liệu",
            "D. Không thể thêm mới nhóm dữ liệu"
        ],
        correct: 1 // Đáp án đúng là "B. Biểu đồ tự động cập nhật"
    },
    {
        question: "Câu 131: Trong MS Excel, điều gì sẽ xảy ra với biểu đồ khi dữ liệu mới được thêm vào bảng dữ liệu trong Chart Range?",
        options: [
            "A. Tự động cập nhật",
            "B. Biểu đồ bị xóa và tạo mới",
            "C. Vùng dữ liệu được chọn lại",
            "D. Không thể thêm mới dữ liệu"
        ],
        correct: 0 // Đáp án đúng là "A. Tự động cập nhật"
    },
    {
        question: "Câu 132: Trong MS Excel, bạn có thể thay đổi lại vị trí của biểu đồ trên bảng tính sau khi nhấp vào biểu đồ và thấy điều gì?",
        options: [
            "A. Một mũi tên",
            "B. Mũi tên màu đen có hai đầu (trái và phải)",
            "C. Mũi tên màu đen có hai đầu (trên và dưới)",
            "D. Mũi tên màu đen có bốn đầu"
        ],
        correct: 3 // Đáp án đúng là "D. Mũi tên màu đen có bốn đầu"
    },
    {
        question: "Câu 133: Trong MS Excel, có thể dịch chuyển vị trí của nội dung biểu đồ bằng cách nhấp vào?",
        options: [
            "A. vùng trắng, mũi tên hai đầu",
            "B. đường viền của biểu đồ, mũi tên hai đầu",
            "C. chuỗi tiêu đề, mũi tên 4 đầu",
            "D. nội dung, mũi tên 4 đầu"
        ],
        correct: 1 // Đáp án đúng là "B. đường viền của biểu đồ, mũi tên hai đầu"
    },
    {
        question: "Câu 134: MS Excel, làm thế nào để một Field List hiển thị lại trong PivotTable khi nó biến mất?",
        options: [
            "A. Nhấp vào ô có dữ liệu bất kỳ",
            "B. Nhấp vào ô bên ngoài bảng",
            "C. Nhấp vào ô trống",
            "D. Nhấp vào biểu tượng fx"
        ],
        correct: 0 // Đáp án đúng là "A. Nhấp vào ô có dữ liệu bất kỳ"
    },
    {
        question: "Câu 135: Trong MS Excel, màu của đối tượng nào sẽ thay đổi nếu chọn Change Colors tại tab Design?",
        options: [
            "A. màu biểu đồ",
            "B. màu trang tính",
            "C. màu bảng tính",
            "D. màu tiêu đề"
        ],
        correct: 0 // Đáp án đúng là "A. màu biểu đồ"
    },
    {
        question: "Câu 136: Trong MS Excel, lệnh nào cung cấp tùy chọn nhanh các thành phần được ưu tiên xuất hiện và vị trí của chúng trong biểu đồ?",
        options: [
            "A. Quick chart",
            "B. Quick format",
            "C. Change color",
            "D. Quick layout"
        ],
        correct: 3 // Đáp án đúng là "D. Quick layout"
    },
    {
        question: "Câu 137: Trong MS Excel, lệnh nào giúp thêm hoặc thay đổi thành phần biểu đồ?",
        options: [
            "A. Format section",
            "B. Select data",
            "C. Effects",
            "D. Add chart elements"
        ],
        correct: 3 // Đáp án đúng là "D. Add chart elements"
    },
    {
        question: "Câu 138: Trong MS Excel, tổ hợp phím dùng để mở nhanh Quick Analysis Gallery?",
        options: [
            "A. Ctrl + A",
            "B. Ctrl + Q",
            "C. Ctrl + G",
            "D. Ctrl + Shift + A"
        ],
        correct: 1 // Đáp án đúng là "B. Ctrl + Q"
    },
    {
        question: "Câu 139: Trong MS Excel, tại Quick Analysis Tools, có thể chọn Iconset bằng tùy chọn nào?",
        options: [
            "A. Formating",
            "B. Tables",
            "C. Chart",
            "D. Total"
        ],
        correct: 0 // Đáp án đúng là "A. Formating"
    },
    {
        question: "Câu 140: Trong MS Excel, tại Quick Analysis Tools, có thể chọn Pivot bằng tùy chọn nào?",
        options: [
            "A. Formating",
            "B. Charts",
            "C. Totals",
            "D. Tables"
        ],
        correct: 3 // Đáp án đúng là "D. Tables"
    },
    {
        question: "Câu 141: Trong MS Excel, dùng cách nào nhanh để thu gọn và sắp xếp lượng lớn dữ liệu?",
        options: [
            "A. Trang tính",
            "B. PivotTable",
            "C. Biểu đồ",
            "D. Đồ thị"
        ],
        correct: 1 // Đáp án đúng là "B. PivotTable"
    },
    {
        question: "Câu 142: Trong MS Excel, nút gì dùng để chèn một hình ảnh từ máy tính?",
        options: [
            "A. Table",
            "B. Photo",
            "C. Chart",
            "D. Picture"
        ],
        correct: 3 // Đáp án đúng là "D. Picture"
    },
    {
        question: "Câu 143: Trong MS Excel, thao tác thực hiện để vẽ đường xiên góc 45 độ bắt đầu kéo từ một điểm?",
        options: [
            "A. giữ phím Shift, thả chuột",
            "B. giữ phím Ctrl, thả chuột",
            "C. giữ chuột, thả phím Ctrl",
            "D. giữ chuột, thả phím Shift"
        ],
        correct: 0 // Đáp án đúng là "A. giữ phím Shift, thả chuột"
    },
    {
        question: "Câu 144: Trong MS Excel, trên thanh Menu, tab nào có nút Text Box?",
        options: [
            "A. file",
            "B. INSERT",
            "C. PAGE LAYOUT",
            "D. view"
        ],
        correct: 1 // Đáp án đúng là "B. INSERT"
    },
    {
        question: "Câu 145: Trong MS Excel, khi các hình ảnh chồng lên nhau, bạn sử dụng tính năng gì để thay đổi thứ tự sắp xếp?",
        options: [
            "A. Size",
            "B. Page Setup",
            "C. Arrange",
            "D. Orientation"
        ],
        correct: 2 // Đáp án đúng là "C. Arrange"
    },
    {
        question: "Câu 146: Trong MS Excel, lệnh nào cho phép xem danh sách các hình ảnh trong trang tính?",
        options: [
            "A. Bring forward",
            "B. Send backward",
            "C. Selection pane",
            "D. Align"
        ],
        correct: 2 // Đáp án đúng là "C. Selection pane"
    },
    {
        question: "Câu 147: Trong MS Excel, lựa chọn nào là mặc định?",
        options: [
            "A. di chuyển và thay đổi kích thước ô",
            "B. di chuyển nhưng không thay đổi kích thước ô",
            "C. không di chuyển hoặc thay đổi kích thước ô",
            "D. văn bản bị khóa"
        ],
        correct: 1 // Đáp án đúng là "B. di chuyển nhưng không thay đổi kích thước ô"
    },
    {
        question: "Câu 148: Trong MS Excel, tab nào trên thanh Menu chứa nhóm minh họa với SmartArt?",
        options: [
            "A. HOME",
            "B. INSERT",
            "C. PAGE LAYOUT",
            "D. DATA"
        ],
        correct: 1 // Đáp án đúng là "B. INSERT"
    },
    {
        question: "Câu 149: Trong MS Excel, điều khiển mũi tên ở đâu để mở Text pane, nơi có thể nhập tên của mỗi bước?",
        options: [
            "A. cạnh trên",
            "B. cạnh phải",
            "C. cạnh dưới",
            "D. cạnh trái"
        ],
        correct: 1 // Đáp án đúng là "B. cạnh phải"
    },
    {
        question: "Câu 150: Trong MS Excel, trong Defined Names, có thể thêm tên vùng dữ liệu vào công thức bằng cách chọn?",
        options: [
            "A. Consolidate",
            "B. Calculate now",
            "C. Use in formula",
            "D. Define name"
        ],
        correct: 2 // Đáp án đúng là "C. Use in formula"
    }
];
// Chọn ngẫu nhiên 30 câu hỏi từ 150 câu hỏi
function getRandomQuestions() {
    const shuffled = questions.sort(() => 0.5 - Math.random());
    return shuffled.slice(0, 30);
}

// Hiển thị câu hỏi lên trang
function loadQuiz() {
    const quizContainer = document.getElementById('quiz-form');
    quizContainer.innerHTML = ''; // Xóa câu hỏi cũ
    const randomQuestions = getRandomQuestions();
    randomQuestions.forEach((q, index) => {
        const questionElement = document.createElement('div');
        questionElement.classList.add('question');
        questionElement.innerHTML = `<p>${index + 1}. ${q.question}</p>`; // Hiển thị số thứ tự câu hỏi
        
        q.options.forEach((option, i) => {
            questionElement.innerHTML += `
                <div class="options">
                    <label>
                        <input type="radio" name="question${index}" value="${i}">
                        ${option}
                    </label>
                </div>
            `;
        });
        
        questionElement.innerHTML += `<div class="feedback" id="feedback${index}"></div>`; // Khu vực hiển thị kết quả
        quizContainer.appendChild(questionElement);
    });
}

// Kiểm tra câu trả lời và chấm điểm
function submitQuiz() {
    const totalQuestions = 30; // Tổng số câu hỏi
    let score = 0;
    const quizContainer = document.getElementById('quiz-form');
    const selectedAnswers = Array.from(quizContainer.querySelectorAll('input[type="radio"]:checked'));

    selectedAnswers.forEach((answer, index) => {
        const questionIndex = parseInt(answer.name.replace('question', ''));
        const feedbackElement = document.getElementById(`feedback${questionIndex}`);

        if (parseInt(answer.value) === questions[questionIndex].correct) {
            score++;
            feedbackElement.innerHTML = '<span class="correct-answer">Đáp án đúng!</span>';
        } else {
            feedbackElement.innerHTML = `<span class="wrong-answer">Sai! Đáp án đúng là: ${questions[questionIndex].options[questions[questionIndex].correct]}</span>`;
        }
    });

    const percentage = (score / totalQuestions) * 10; // Tính điểm theo thang điểm 10
    document.getElementById('result').innerHTML = `<h2>Kết quả:</h2><p>Điểm của bạn: ${percentage.toFixed(2)}</p>`;
}

// Hàm reset quiz
function resetQuiz() {
    document.getElementById('result').innerHTML = ''; // Xóa kết quả
    loadQuiz(); // Tải lại câu hỏi
}

// Gọi hàm để tải câu hỏi ngay khi trang được tải
window.onload = loadQuiz;