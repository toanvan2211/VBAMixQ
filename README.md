# Macro trộn đề thi trắc nghiệm trên định dạng WORD
**Đây là Tiểu luận tốt nghiệp của [Văn Công Toàn](https://www.facebook.com/tonten2211/) - học kỳ 8 2021** 

## Đặc điểm
Sử dụng ngôn ngữ `Visual Basic for Application(VBA)` 

Xử lí đề trắc nghiệm trên định dạng `Word`

Macro có thể khoanh vùng câu hỏi, nhận diện **câu hỏi**, **câu trả lời**, **câu trả lời đúng** 

Có thể trộn trực tiếp trên file hiện tại, hoặc xuất ra file khác

## Quy định về chuẩn đề trắc nghiệm
* Các câu hỏi bắt đầu bằng từ khoá `"Câu"`.
* Mỗi câu hỏi phải có lớn hơn hoặc từ 2 câu trả lời.
* Các câu trả lời đúng được đánh dấu bằng `Gạch chân` hoặc `Tô đỏ`.
* Các phần câu hỏi dạng chùm (1 nội dung có nhiều câu hỏi cho nội dung đó): đánh dấu `#` ở đầu và cuối phần câu hỏi chùm.

## Chức năng
* Trộn đề thi
  * - [x] Chỉ trộn câu hỏi
  * - [x] Chỉ trộn câu trả lời
  * - [x] Trộn cả hai lựa chọn trên
    * - [x] Trộn đề trên tài liệu hiện tại
    * - [x] Trộn vùng được tô đen
    * - [x] Trộn đề và xuất đề
    * - [ ] Trộn đề và xuất đề kèm đáp án
* Format đề thi
  - [x] Canh lề tabstop
  - [x] Đánh dấu đáp án đúng (gạch chân | in đỏ)
  - [x] Bỏ đánh dấu
  - [x] Điều chỉnh size chữ vùng câu hỏi trắc nghiệm
  - [x] Điều chỉnh alignment vùng câu hỏi trắc nghiệm
  - [ ] Thêm dấu chấm cuối câu trả lời
  
```VBA
Dim str As String
str = "Hello VBA"
Debug.Print str
```
