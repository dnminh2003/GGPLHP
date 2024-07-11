const express = require("express");
const path = require('path');
const { google } = require("googleapis");

const app = express();
app.use(express.json()); // Middleware để parse JSON body
// Cấu hình EJS làm view engine
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Cấu hình thư mục tĩnh cho các file assets
app.use('/assets', express.static(path.join(__dirname, 'assets')));

app.use(express.urlencoded({ extended: true }));

// Đường dẫn tới file credentials.json và ID của spreadsheet
const credentials = require("./credentials.json");
const spreadsheetId = "1Vnwo4KcFaHawodEu6pYS40m2vHTzU-sppB7GOBwRsho";

// Cấu hình OAuth2
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

// Middleware để xây dựng và xác thực client của Google Sheets
async function googleSheetsMiddleware(req, res, next) {
  try {
    // Tạo client
    const client = await auth.getClient();
    req.googleSheets = google.sheets({ version: "v4", auth: client });
    next();
  } catch (error) {
    console.error("Error setting up Google Sheets client:", error);
    res.status(500).send("Error setting up Google Sheets client");
  }
}

// Sử dụng middleware
app.use(googleSheetsMiddleware);

// Route cơ bản
app.get('/', (req, res) => {
  res.render('index');
});

// Route cho trang 'pages-blank'
app.get('/pages-blank', (req, res) => {
  res.render('pages-blank');
});


// Route để hiển thị danh sách học sinh của 1 lớp http://localhost:3000/LOP6A
app.get("/:classname", async (req, res) => {
  try {
    const { googleSheets } = req;
    const { classname } = req.params;

    // Đọc dữ liệu từ bảng tính
    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${classname}!A:AQ`,
    });

    // Kiểm tra và xử lý kết quả trả về
    const rows = response.data.values || [];

    // Trả về dữ liệu cho client
    // res.status(200).json({ rows });
    res.render('student', { rows, classname });
  } catch (err) {
    console.error("Error retrieving data:", err);
    res.status(500).send("Error retrieving data");
  }
});


// Route để thêm dữ liệu học sinh mới / Thêm học sinh.
// http://localhost:3000/add
app.post('/add', async (req, res) => {
  try {
    const { googleSheets } = req;
    const { classname, studentName, parentName, parentPhone, note, parentEmail } = req.body;

    const parentPhoneFinal = "";
    if (parentPhone){
      const parentPhoneFinal = "'" + parentPhone.toString();
    }
    
    // Get the current month (1-12)
    const currentMonth = new Date().getMonth() + 1;

    // Generate the monthly data array
    const monthlyData = Array(12).fill(null).map((_, index) => {
      if (index < currentMonth - 1) {
        return ['X', '', ''];
      } else {
        return ['Chưa', '', ''];
      }
    }).flat();

    // Create the new row with the student data
    const newRow = [
      studentName,
      parentPhoneFinal,
      parentName,
      note,
      parentEmail,
      ...monthlyData, // Default "Chưa" and no payment information for the current and future months, "X" for past months.
      'Chưa hoàn thành', // Status
      new Date().toLocaleString(), // Update time
    ];

    // Lấy thông tin về dòng cuối cùng có dữ liệu
    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${classname}!A:A`,
    });
    const rows = response.data.values || [];
    const lastRow = rows.length + 1; // Dòng mới sẽ là dòng kế tiếp sau dòng cuối cùng có dữ liệu

    // Sao chép định dạng từ dòng cuối cùng đến dòng mới
    const copyPasteRequest = {
      requests: [
        {
          copyPaste: {
            source: {
              sheetId: 0,
              startRowIndex: lastRow - 2, // Dòng cuối cùng (0-based index)
              endRowIndex: lastRow - 1,
              startColumnIndex: 0,
              endColumnIndex: 43, // Số cột bạn muốn sao chép định dạng và dữ liệu validation
            },
            destination: {
              sheetId: 0,
              startRowIndex: lastRow - 1, // Dòng mới (0-based index)
              endRowIndex: lastRow,
              startColumnIndex: 0,
              endColumnIndex: 43,
            },
            pasteType: 'PASTE_FORMAT', // Sao chép định dạng
          },
        },
        {
          copyPaste: {
            source: {
              sheetId: 0,
              startRowIndex: lastRow - 2, // Dòng cuối cùng (0-based index)
              endRowIndex: lastRow - 1,
              startColumnIndex: 0,
              endColumnIndex: 43, // Số cột bạn muốn sao chép định dạng và dữ liệu validation
            },
            destination: {
              sheetId: 0,
              startRowIndex: lastRow - 1, // Dòng mới (0-based index)
              endRowIndex: lastRow,
              startColumnIndex: 0,
              endColumnIndex: 43,
            },
            pasteType: 'PASTE_DATA_VALIDATION', // Sao chép dữ liệu validation
          },
        },
      ],
    };

    await googleSheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: copyPasteRequest,
    });

    // Thêm dữ liệu vào dòng mới
    await googleSheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${classname}!A${lastRow}:AQ${lastRow}`,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [newRow],
      },
    });

    res.status(200).json({ success: true, message: 'Student added successfully' });
  } catch (err) {
    console.error('Error adding student:', err);
    res.status(500).json({ success: false, message: 'Error adding student', err: err.message });
  }
});


// Route để sửa thông tin học sinh  // http://localhost:3000/update/10
// Route để sửa thông tin học sinh
app.post('/update/:rowId', async (req, res) => {
  try {
    const { googleSheets } = req;
    const { rowId } = req.params;
    const { classname, studentName, parentName, parentPhone, note, parentEmail, monthsPaid, monthsPaid_amount, otherMonth } = req.body;

    const parentPhoneFinal = "";
    if (parentPhone){
      const parentPhoneFinal = "'" + parentPhone.toString();
    }
    // Chuẩn bị dữ liệu mới cho hàng cần sửa
    const updatedRow = [
      studentName,
      parentPhoneFinal,
      parentName,
      note,
      parentEmail,
      otherMonth,otherMonth,otherMonth,otherMonth,otherMonth,otherMonth,otherMonth,otherMonth,otherMonth,otherMonth,otherMonth,otherMonth,otherMonth
    ];

    // Cập nhật trạng thái thanh toán cho từng tháng đã chỉ định
    if (monthsPaid) {
      if (monthsPaid >= 1 && monthsPaid <= 12) {
        const index = (monthsPaid - 1) * 3 + 5; // Xác định vị trí bắt đầu của các tháng
        updatedRow[index] = 'Đã nộp'; // Cập nhật trạng thái đã nộp
        updatedRow[index + 1] = new Date().toLocaleDateString(); // Ngày nộp
        updatedRow[index + 2] = monthsPaid_amount; // Số tiền học thêm      
      }
    }
    updatedRow[42] = new Date().toLocaleString(); // Ngày cập nhật, ở cột cuối cùng

    // Cập nhật dữ liệu vào hàng đã chỉ định
    await googleSheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${classname}!A${rowId}:AQ${rowId}`,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [updatedRow],
      },
    });

    res.status(200).json({ success: true, message: 'Student updated successfully' });
  } catch (err) {
    console.error('Error updating student:', err);
    res.status(500).json({ success: false, message: 'Error updating student', err: err.message });
  }
});

// Route để xoá thông tin học sinh
// app.post('/delete/:rowId', async (req, res) => {
//   try {
//     const { googleSheets } = req;
//     const { rowId } = req.params;
//     const { classname } = req.body;

//     // Tạo request để xóa hàng
//     const deleteRequest = {
//       requests: [
//         {
//           deleteDimension: {
//             range: {
//               sheetId: 0, // ID của sheet, 0 là sheet đầu tiên
//               dimension: 'ROWS',
//               startIndex: parseInt(rowId) - 1, // Chỉ số hàng cần xoá (0-based index)
//               endIndex: parseInt(rowId), // endIndex là chỉ số hàng tiếp theo của hàng cần xoá
//             },
//           },
//         },
//       ],
//     };

//     // Thực hiện xóa hàng bằng cách gọi spreadsheets.batchUpdate
//     await googleSheets.spreadsheets.batchUpdate({
//       spreadsheetId,
//       resource: deleteRequest,
//     });

//     res.status(200).json({ message: 'Student deleted successfully' });
//   } catch (err) {
//     console.error('Error deleting student:', err);
//     res.status(500).json({ error: 'Error deleting student' });
//   }
// });

app.delete('/delete/:rowId', async (req, res) => {
  try {
    const { googleSheets } = req;
    const { rowId } = req.params;
    const { classname } = req.body; // Get classname from body

    // Convert rowId to a 0-based index for the API request
    const rowIndex = parseInt(rowId) - 1;

    // Create the request body for deleting the row
    const requestBody = {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId: (await googleSheets.spreadsheets.get({
                spreadsheetId
              })).data.sheets.find(sheet => sheet.properties.title === classname).properties.sheetId,
              dimension: 'ROWS',
              startIndex: rowIndex,
              endIndex: rowIndex + 1
            }
          }
        }
      ]
    };

    // Send the request to delete the row
    await googleSheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: requestBody
    });

    res.status(200).json({ success: true, message: 'Student deleted successfully' });
  } catch (err) {
    console.error('Error deleting student:', err);
    res.status(500).json({ success: false, error: 'Error deleting student', err: err.message });
  }
});



// Route thu tiền học phí:
app.get("/:classname/collect/:rowId", async (req, res) => {
  try {
    const { googleSheets } = req;
    const { classname, rowId } = req.params;

    // Đọc dữ liệu từ bảng tính
    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${classname}!A:AZ`, // Điều chỉnh phạm vi nếu cần
    });

    const rows = response.data.values || [];
    
    // Lấy hàng dữ liệu tương ứng với rowId
    const studentRow = rows[rowId];

    if (!studentRow) {
      return res.status(404).send("Student not found");
    }

    const studentName = studentRow[0];

    // Lọc thông tin các tháng có giá trị ở ô "Nộp" không phải là "Chưa"
    const feeInfo = [];
    const monthHeaders = ['Tháng 1', 'Tháng 2', 'Tháng 3', 'Tháng 4', 'Tháng 5', 'Tháng 6', 'Tháng 7', 'Tháng 8', 'Tháng 9', 'Tháng 10', 'Tháng 11', 'Tháng 12'];
    const statusStartIndex = 5; // Vị trí bắt đầu của cột "Nộp" tháng 1
    const increment = 3; // Khoảng cách giữa các cột

    for (let i = 0; i < 12; i++) {
      const statusIndex = statusStartIndex + i * increment;
      const dateIndex = statusIndex + 1;
      const amountIndex = statusIndex + 2;
      
      if (studentRow[statusIndex] !== 'X') {
        feeInfo.push({
          month: monthHeaders[i],
          status: studentRow[statusIndex],
          date: studentRow[dateIndex] || '',
          amount: studentRow[amountIndex] || '',
        });
      }
    }

    // Trả về dữ liệu cho client
    res.render('collect', { feeInfo, classname, rowId, studentName });
  } catch (err) {
    console.error("Error retrieving data:", err);
    res.status(500).send("Error retrieving data");
  }
});



// Khởi động server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
