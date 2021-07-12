# Công cụ trộn đề thi trắc nghiệm trên định dạng WORD
**Đây là Tiểu luận tốt nghiệp của [Văn Công Toàn](https://www.facebook.com/tonten2211/) - học kỳ 8 2021** 

## Giới thiệu
Xây dựng bằng ngôn ngữ `Visual Basic for Application(VBA)`, chạy ổn định trên phiên bản MS Word 2013 - 2016.
Công cụ hỗ trợ trộn đề trắc nghiệm cho tài liệu `Word`, 
nhận biết và khoanh vùng câu hỏi, nhận diện **câu hỏi**, **câu trả lời**, **câu trả lời đúng** .
Công cụ là một tập tin Word (docm).

Người sử dụng sao chép đề cần trộn vào tập tin, sử dụng các chức năng trên thanh menu "Trộn trắc nghiệm" để thao tác.
![Tool Menu](/img/menuGUI.png)

## Chức năng
* Trộn đề thi
  * - [x] Chỉ trộn câu hỏi
  * - [x] Chỉ trộn câu trả lời
  * - [x] Trộn cả hai lựa chọn trên
  * - [x] Trộn đề hiện tại
  * - [x] Trộn vùng được tô đen
  * - [x] Trộn đề và xuất đề
  * - [x] Xuất đáp án
  * - [x] Chèn đáp án
  * - [x] Trộn đề và xuất đề kèm đáp án
* - [x] Tạo đề mới từ ngân hàng trắc nghiệm
* Format đề thi
  - [x] Canh lề tabstop
  - [x] Đánh dấu đáp án đúng (gạch chân | in đỏ)
  - [x] Bỏ đánh dấu
  - [x] Đánh dấu lại thứ tự câu
  - [x] Điều chỉnh size chữ vùng câu hỏi trắc nghiệm
  - [x] Thêm dấu chấm cuối câu trả lời

## Quy định về cấu trúc câu hỏi & ngân hàng trắc nghiệm
### Câu hỏi
* Các câu hỏi bắt đầu bằng từ khoá `"Câu"`.
* Câu hỏi được phép chứa hình ảnh.

![Question contain image](/img/structureQ-1.png)

* Mỗi câu hỏi phải có lớn hơn hoặc từ 2 câu trả lời.
* Ký tự đầu tiên của câu trả lời được đánh dấu theo thứ tự alphabet.
* Một câu trả lời có thể bao gồm nhiều hàng.
* Một hàng có thể bao gồm nhiều câu trả lời.

![Question contain image](/img/structureQ-3.png)

* Câu trả lời có thể chứa các định dạng công thức "Equation".
* Các câu trả lời đúng được đánh dấu bằng `Gạch chân` hoặc `Tô đỏ`.
### Ngân hàng trắc nghiệm
* Các câu hỏi bắt buộc phải thuộc một "Chương" và "Mức" bất kỳ.
* Chương > Mức > Câu hỏi

![Bank QUestion](/img/bankQ.png)

* Cấu trúc của câu hỏi tương tự như phần trên.
