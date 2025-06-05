// ====== ตั้งค่า ======
const LINE_TOKEN = 'YOUR_LINE_CHANNEL_ACCESS_TOKEN'; // ใส่ Token ของคุณ
const WEBHOOK_URL = 'https://your-backend.com/api/excel-sync'; // เปลี่ยนเป็น endpoint backend ของคุณ

// ====== Trigger หลัก: ทำงานเมื่อมีการแก้ไขใน Sheet ======
function onEdit(e) {
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const editedColumn = e.range.getColumn();
  const editedRow = e.range.getRow();

  // กรณีแก้ไขคอลัมน์สถานะ (8)
  if (editedColumn === 8 && editedRow > 1) {
    handleStatusChange(e, sheet, editedRow);
  }
  // กรณีแก้ไขคอลัมน์ Textbox (13)
  else if (editedColumn === 13 && editedRow > 1) {
    handleTextboxChange(e, sheet, editedRow);
  }
}

// ====== ฟังก์ชันจัดการการเปลี่ยนสถานะ ======
function handleStatusChange(e, sheet, editedRow) {
  const data = sheet.getRange(editedRow, 1, 1, 13).getValues()[0];

  if (data.length < 12) {
    Logger.log("❌ ไม่พบข้อมูลครบ 12 คอลัมน์");
    return;
  }

  const payload = {
    ticket_id: data[0],
    user_id: data[1],
    email: data[2],
    name: data[3],
    phone: data[4],
    department: data[5],
    created_at: data[6] || new Date().toISOString(),
    status: data[7],
    appointment: data[8],
    requeste: data[9],
    report: data[10],
    textbox: data[11]
  };

  // ส่งข้อมูลไปยัง backend
  sendWebhook(payload);

  // ส่ง Flex Message ไปยัง LINE
  sendStatusFlexMessage(payload.user_id, payload);
}

// ====== ฟังก์ชันจัดการการเปลี่ยน Textbox ======
function handleTextboxChange(e, sheet, editedRow) {
  const data = sheet.getRange(editedRow, 1, 1, 13).getValues()[0];
  if (data.length < 13) return;

  const newText = data[12];
  const oldText = e.oldValue || '';

  // ตรวจสอบว่าเป็นข้อความจากเจ้าหน้าที่ (ไม่มี [ระบบ] หรือ [ผู้ใช้])
  if (!newText || newText === oldText ||
      newText.includes("[ระบบ]") ||
      newText.includes("[ผู้ใช้]")) {
    return;
  }

  const userId = data[1];
  if (!userId) return;

  // แยกข้อความล่าสุดที่เพิ่มโดยเจ้าหน้าที่
  const messages = newText.split('\n').filter(m => m.trim() !== '');
  const lastMessage = messages[messages.length - 1];

  // ตรวจสอบว่าเป็นข้อความจากเจ้าหน้าที่จริงๆ
  if (!lastMessage.includes("[เจ้าหน้าที่]")) {
    const timestamp = new Date().toLocaleString();
    const updatedText = `${newText}\n💼 [เจ้าหน้าที่] ${timestamp}: ${lastMessage}`;
    sheet.getRange(editedRow, 13).setValue(updatedText);
    return;
  }

  // ส่งข้อความไปยังผู้ใช้
  const messageContent = lastMessage.split(':').slice(1).join(':').trim();
  const sendResult = sendLineMessage(userId, messageContent, true); // true = เป็นการตอบกลับ

  if (sendResult === 'success') {
    sheet.getRange(editedRow, 8).setValue('Pending');
  }
}

// ====== ฟังก์ชันส่งข้อมูลไป backend ======
function sendWebhook(payload) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const res = UrlFetchApp.fetch(WEBHOOK_URL, options);
    Logger.log('✅ ส่งข้อมูลไป backend สำเร็จ: ' + res.getContentText());
    return 'success';
  } catch (err) {
    Logger.log('❌ ส่งข้อมูลไป backend ล้มเหลว: ' + err.message);
    return 'error';
  }
}

// ====== ฟังก์ชันส่ง Flex Message สถานะ Ticket ไป LINE ======
function sendStatusFlexMessage(userId, payload) {
  const url = 'https://api.line.me/v2/bot/message/push';

  // จัดรูปแบบวันที่นัดหมาย (ถ้ามี)
  const appointmentDate = payload.appointment ?
    new Date(payload.appointment).toLocaleString('th-TH', {
      day: 'numeric',
      month: 'numeric',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    }) : 'ยังไม่ได้กำหนด';

  // สีของสถานะ (ปรับตามสถานะ)
  let statusColor;
  switch(payload.status) {
    case 'Pending':
      statusColor = '#FF9900'; // สีส้ม
      break;
    case 'Completed':
      statusColor = '#00AA00'; // สีเขียว
      break;
    case 'Rejected':
      statusColor = '#FF0000'; // สีแดง
      break;
    case 'In Progress':
      statusColor = '#0066FF'; // สีน้ำเงิน
      break;
    default:
      statusColor = '#666666'; // สีเทา
  }

  const flexMessage = {
    to: userId,
    messages: [{
      type: "flex",
      altText: "อัปเดตสถานะ Ticket ของคุณ",
      contents: {
        type: "bubble",
        size: "giga",
        header: {
          type: "box",
          layout: "vertical",
          contents: [
            {
              type: "text",
              text: "📢 อัปเดตสถานะ Ticket",
              weight: "bold",
              size: "lg",
              color: "#FFFFFF",
              align: "center"
            }
          ],
          backgroundColor: "#005BBB",
          paddingAll: "20px"
        },
        body: {
          type: "box",
          layout: "vertical",
          contents: [
            // ข้อมูล Ticket
            {
              type: "box",
              layout: "horizontal",
              contents: [
                {
                  type: "text",
                  text: "หมายเลข",
                  weight: "bold",
                  size: "sm",
                  flex: 2,
                  color: "#666666"
                },
                {
                  type: "text",
                  text: payload.ticket_id,
                  size: "sm",
                  flex: 4,
                  align: "end"
                }
              ],
              spacing: "sm",
              margin: "md"
            },
            {
              type: "separator",
              margin: "md"
            },
            {
              type: "box",
              layout: "horizontal",
              contents: [
                {
                  type: "text",
                  text: "ชื่อ",
                  weight: "bold",
                  size: "sm",
                  flex: 2,
                  color: "#666666"
                },
                {
                  type: "text",
                  text: payload.name,
                  size: "sm",
                  flex: 4,
                  align: "end"
                }
              ],
              spacing: "sm",
              margin: "md"
            },
            {
              type: "separator",
              margin: "md"
            },
            {
              type: "box",
              layout: "horizontal",
              contents: [
                {
                  type: "text",
                  text: "แผนก",
                  weight: "bold",
                  size: "sm",
                  flex: 2,
                  color: "#666666"
                },
                {
                  type: "text",
                  text: payload.department,
                  size: "sm",
                  flex: 4,
                  align: "end"
                }
              ],
              spacing: "sm",
              margin: "md"
            },
            {
              type: "separator",
              margin: "md"
            },
            {
              type: "box",
              layout: "horizontal",
              contents: [
                {
                  type: "text",
                  text: "เบอร์ติดต่อ",
                  weight: "bold",
                  size: "sm",
                  flex: 2,
                  color: "#666666"
                },
                {
                  type: "text",
                  text: payload.phone,
                  size: "sm",
                  flex: 4,
                  align: "end"
                }
              ],
              spacing: "sm",
              margin: "md"
            },
            {
              type: "separator",
              margin: "md"
            },
            {
              type: "box",
              layout: "horizontal",
              contents: [
                {
                  type: "text",
                  text: "Type",
                  weight: "bold",
                  size: "sm",
                  flex: 2,
                  color: "#666666"
                },
                {
                  type: "text",
                  text: payload.textbox,
                  size: "sm",
                  flex: 4,
                  align: "end"
                }
              ],
              spacing: "sm",
              margin: "md"
            },
            {
              type: "separator",
              margin: "md"
            },
            {
              type: "box",
              layout: "horizontal",
              contents: [
                {
                  type: "text",
                  text: "ปัญหา",
                  weight: "bold",
                  size: "sm",
                  flex: 2,
                  color: "#666666"
                },
                {
                  type: "text",
                  text: payload.report,
                  size: "sm",
                  flex: 4,
                  align: "end",
                  wrap: true
                },
                {
                  type: "text",
                  text: payload.requeste,
                  size: "sm",
                  flex: 4,
                  align: "end",
                  wrap: true
                },
              ],
              spacing: "sm",
              margin: "md"
            },
            {
              type: "separator",
              margin: "md"
            },
            {
              type: "box",
              layout: "horizontal",
              contents: [
                {
                  type: "text",
                  text: "วันที่นัดหมาย",
                  weight: "bold",
                  size: "sm",
                  flex: 2,
                  color: "#666666"
                },
                {
                  type: "text",
                  text: appointmentDate,
                  size: "sm",
                  flex: 4,
                  align: "end"
                }
              ],
              spacing: "sm",
              margin: "md"
            },
            // สถานะ (ไฮไลท์)
            {
              type: "box",
              layout: "vertical",
              contents: [
                {
                  type: "text",
                  text: "สถานะล่าสุด",
                  weight: "bold",
                  size: "sm",
                  color: "#666666",
                  margin: "md"
                },
                {
                  type: "text",
                  text: payload.status,
                  weight: "bold",
                  size: "xl",
                  color: statusColor,
                  align: "center",
                  margin: "sm"
                }
              ],
              backgroundColor: "#F5F5F5",
              cornerRadius: "md",
              margin: "xl",
              paddingAll: "md"
            }
          ],
          spacing: "md",
          paddingAll: "20px"
        },
        footer: {
          type: "box",
          layout: "vertical",
          contents: [
            {
              type: "text",
              text: "ขอบคุณที่ใช้บริการของเรา",
              size: "xs",
              color: "#888888",
              align: "center"
            }
          ],
          paddingAll: "10px"
        }
      }
    }]
  };

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify(flexMessage),
    muteHttpExceptions: true
  };

  try {
    const res = UrlFetchApp.fetch(url, options);
    return res.getResponseCode() === 200 ? 'success' : 'error';
  } catch (err) {
    Logger.log('❌ ส่ง LINE Flex Message ล้มเหลว: ' + err.message);
    return 'error';
  }
}

// ====== ฟังก์ชันส่งข้อความไป LINE (Push) ======
function sendLineMessage(userId, message, isReply = false) {
  const url = 'https://api.line.me/v2/bot/message/push';

  let payload;

  if (isReply) {
    // กรณีตอบกลับจากเจ้าหน้าที่ (รูปแบบ Flex Message)
    payload = {
      to: userId,
      messages: [{
        type: "flex",
        altText: "ข้อความจากเจ้าหน้าที่",
        contents: {
          type: "bubble",
          body: {
            type: "box",
            layout: "vertical",
            contents: [
              {
                type: "text",
                text: "💼 ตอบกลับจากเจ้าหน้าที่",
                weight: "bold",
                size: "lg",
                color: "#005BBB"
              },
              {
                type: "text",
                text: message,
                wrap: true,
                margin: "md"
              },
              {
                type: "text",
                text: "พิมพ์ 'จบ' เพื่อสิ้นสุดการสนทนา",
                size: "sm",
                color: "#AAAAAA",
                margin: "md"
              }
            ]
          }
        }
      }]
    };
  } else {
    // กรณีแจ้งสถานะ (รูปแบบข้อความธรรมดา)
    payload = {
      to: userId,
      messages: [{ type: 'text', text: message }]
    };
  }

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const res = UrlFetchApp.fetch(url, options);
    return res.getResponseCode() === 200 ? 'success' : 'error';
  } catch (err) {
    Logger.log('❌ ส่ง LINE ล้มเหลว: ' + err.message);
    return 'error';
  }
}

// ====== ฟังก์ชันทดสอบ (รันใน Script Editor ได้) ======
function testSendWebhook() {
  const payload = {
    TicketID: 'TICKET001',
    Name: 'นายทดสอบ ระบบ',
    Status: 'Pending'
  };
  sendWebhook(payload);
}

function testSendLine() {
  const testUserId = 'YOUR_LINE_USER_ID';
  sendLineMessage(testUserId, 'ทดสอบส่งข้อความจาก Apps Script');
}
