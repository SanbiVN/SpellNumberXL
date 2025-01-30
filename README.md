# SpellNumberXL - Hàm đọc số thành chữ, chuyển chữ thành số cho Excel
 Đọc số tiền, đọc số thập phân, đọc số trong chuỗi

[Click vào đây để tải xuống](https://github.com/SanbiVN/SpellNumberXL/releases/download/SpellNumber/SpellNumberXL.xlsm)

[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/SpellNumberXL/total.svg)](https://github.com/SanbiVN/SpellNumberXL/releases/download/SpellNumber/SpellNumberXL.xlsm) 

***Mật khẩu VBA là 1 (nếu có)


*** Hướng dẫn
Hàm: 	```=SpellNumVN(Text,[Đối_số_cài_đặt])```

| **Hàm đối số cài đặt**           | **Diễn tả**                                                                                                                                                      |
| -------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **snNumberIndexs**(Indexs)       | Nếu để số: 0 thì đọc tất cả số, 1 đọc số thứ nhất, {1;3} hay {1\\ 3} đọc số vị trí 1 và 3, Số âm -2 thì trích và đọc duy nhất vị trí số thứ 2 trong chuỗi chứa số. |
| **snNotReadZero**()              | Không đọc số không hàng Trăm và Triệu                                                                                                                              |
| **snNotGroupDivision**()         | Không nhóm mỗi 1000 hoặc 1 tỷ đơn vị với dấu ngắt.                                                                                                                 |
| **snReplaceNumbers**()           | Lựa chọn thay thế hoặc thêm vào đằng sau số được đọc                                                                                                               |
| **snSpellPercent**()             | Đọc phần phân số                                                                                                                                                   |
| **snDotPercent**()               | Đọc là chấm thay cho phẩy                                                                                                                                          |
| **snSpellDivision**()            | Đọc cả phép chia ở phần thập phân                                                                                                                                  |
| **snSentenceSpace**(" ")         | Dấu cách khi đọc số                                                                                                                                                |
| **snUnitCode(**"USD",1,True**)** | Thêm đơn vị tiền tệ "Đô-la" và quốc gia đại diện "Mỹ"                                                                                                              |
| **snText**([Left],[Right])       | Thêm chuỗi bên trái và bên phải nếu cần thiết                                                                                                                      |
| **snSoutherners()**              | Đọc theo miền nam (ngàn+lẻ)                                                                                                                                        |


Các kiểu đọc đơn vị tiền tệ (Hàm đối số):
(Có gần 200 đơn vị tiền tệ)

| Hàm                              | Diễn tả |
| -------------------------------- | ---------------------------------------------- |
| **snUnitCode(**"USD"**)**        | 1 - thêm đơn vị tiền tệ "Đô-la."               |
| **snUnitCode(**"USD",2**)**      | 2 - thêm đơn vị tiền tệ "(Đô-la.)"             |
| **snUnitCode(**"USD",3**)**      | 3 - thêm đơn vị tiền tệ "[Đô-la.]"             |
| **snUnitCode(**"USD",1,True**)** | "Đô-la Mỹ."  đơn vị tiền tệ và tên địa lý      |
| **snUnitCode(**"VND",1,True**)** | "Việt Nam đồng."  đơn vị tiền tệ và tên địa lý |

Có 6 kiểu chữ viết Hoa thường:

| **Hàm**               | **Kiểu chữ viết**                | **Ví dụ**                                                                             |
| --------------------- | -------------------------------- | ------------------------------------------------------------------------------------- |
| **snCaseLower**()     | Chữ thường                       | một triệu không trăm năm mươi ngàn đồng                                               |
| **snCaseSentence**()  | Chữ hoa ký tự đầu tiên của chuỗi | Hai triệu đồng                                                                        |
| **snCaseTitle**()     | Chữ Hoa ký tự đầu tất cả từ      | Bảy Mươi Triệu Hai Trăm Năm Mươi Ngàn                                                 |
| **snCaseUpper**()     | Chữ Hoa                          | BA TRĂM NĂM MƯƠI TRIỆU                                                                |
| **snCaseThousands**() | Chữ Hoa sau mỗi 1000 đơn vị      | Chín triệu không trăm năm mươi tỷ, Ba trăm hai mươi bảy triệu, Năm trăm.              |
| **snCaseBillion**()   | Chữ Hoa sau mỗi 1 tỷ đơn vị      | Hai mươi lăm triệu, không trăm năm mươi sáu tỷ. Ba trăm hai mươi bảy triệu, năm trăm. |


Hàm: 	```=TxtToNum(Text,[Đối_số_cài_đặt])```

| **Hàm đối số cài đặt**      | **Diễn giải**                                                                                                                                                      |
| --------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **ttnNumberIndexs**(Indexs) | Nếu để số: 0 thì đọc tất cả số, 1 đọc số thứ nhất, {1;3} hay {1\\ 3} đọc số vị trí 1 và 3, Số âm -2 thì trích và đọc duy nhất vị trí số thứ 2 trong chuỗi chứa số. |
| **ttnSkipBreaks**()         | Xóa bỏ dấu ngắt câu: chấm, phẩy                                                                                                                                    |

Lưu ý: Để sử dụng được Hàm SpellNumVN trong dự án mới, hãy sao chép module zzzzSpellNumber

*** Các ví dụ
| **Số và số nằm trong văn bản**                                                                      | **SpellNumVn - Đọc số**                                                                                                                                                                                     |
| --------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **0.0**                                                                                             | Không                                                                                                                                                                                                       |
| **1**                                                                                               | Một                                                                                                                                                                                                         |
| **100**                                                                                             | Một trăm                                                                                                                                                                                                    |
| **1000**                                                                                            | Một nghìn                                                                                                                                                                                                   |
| **10000**                                                                                           | Mười nghìn                                                                                                                                                                                                  |
| **100000**                                                                                          | Một trăm ngàn                                                                                                                                                                                               |
| **1000000**                                                                                         | Một triệu                                                                                                                                                                                                   |
| **10000000**                                                                                        | Mười triệu                                                                                                                                                                                                  |
| **100000000**                                                                                       | Một trăm triệu                                                                                                                                                                                              |
| **1000000000**                                                                                      | Một tỷ                                                                                                                                                                                                      |
| **10000000000**                                                                                     | Mười tỷ                                                                                                                                                                                                     |
| **100000000000**                                                                                    | Một trăm tỷ                                                                                                                                                                                                 |
| **1000000000000**                                                                                   | Một nghìn tỷ                                                                                                                                                                                                |
| **1,000000000,000000000,000000000**                                                                 | Một tỷ tỷ tỷ                                                                                                                                                                                                |
| **1,000000000,000000000,001002030**                                                                 | Một tỷ tỷ tỷ. không lẻ một triệu, không trăm lẻ hai nghìn, không trăm ba mươi                                                                                                                               |
| **1,000000000,000000000,001002030**                                                                 | Một tỷ tỷ tỷ. không lẻ một triệu, không trăm lẻ hai nghìn, không trăm ba mươi                                                                                                                               |
| 9.259_210.000.000_000.000.000                                                                       | Chín nghìn, hai trăm năm mươi chín tỷ. hai trăm mười triệu, tỷ                                                                                                                                              |
| **9.26E+21**                                                                                        | Chín nghìn, hai trăm năm mươi chín tỷ. hai trăm mười triệu, tỷ                                                                                                                                              |
| **9.26E-21**                                                                                        | Không, phẩy không chục triệu tỷ tỷ. không triệu chín trăm hai mươi lăm nghìn chín trăm hai mươi mốt                                                                                                         |
| **Số tiền là: 8,079,060,001 VNĐ**                                                                   | Số tiền là: 8,079,060,001 VNĐ (Đọc: Tám tỷ. không trăm bảy mươi chín triệu, không trăm sáu mươi nghìn, không trăm lẻ một Việt Nam đồng.)                                                                    |
| Tám tỷ. không trăm bảy mươi chín triệu, không trăm sáu mươi nghìn, không trăm lẻ một Việt Nam đồng. |
| **Số tiền lương: 8,079,060,001 VNĐ và số tiền thưởng: 3,000,000 VNĐ.**                              | Số tiền lương: 8,079,060,001 VNĐ (Đọc: Tám tỷ. không trăm bảy mươi chín triệu, không trăm sáu mươi ngàn, không trăm lẻ một Việt Nam đồng.) và số tiền thưởng: 3,000,000 VNĐ. (Đọc: Ba triệu Việt Nam đồng.) |
| (Đọc: Ba triệu Việt Nam đồng.)                                                                      |





