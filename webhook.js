const express = require('express');
const { MongoClient } = require('mongodb');
const https = require('https');
const fs = require('fs');
const path = require('path');
const line = require('@line/bot-sdk');
const crypto = require('crypto');

const axios = require('axios');
const exceljs = require('exceljs');

const app = express();

const config = {
    channelAccessToken:'ziclhCLm1Z1KS7oY38QsiYkTExm99ALjVEdacYmN6oj7Xa0ck6XMoKvtltk1lq0KrEgdBYVoSfe8BKq6PfNqLn3k1WU9SLGfHgcJLqffDvdkFKYY+W3miHrx14HyQF48nfJFCnhltsUIC8sgTvh5VwdB04t89/1O/w1cDnyilFU=',
    channelSecret:'b341dfa826fad7b8da5a914b431b9364'
};

const client = new line.Client(config);
// 保存文件的本地目录
const fileDirectory = path.join(__dirname, 'files');

// 创建文件目录（如果不存在的话）
if (!fs.existsSync(fileDirectory)) {
    fs.mkdirSync(fileDirectory);
}
let currentTask = {};
// 設定 axios 的 headers
const headers = {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${config.channelAccessToken}`
};

// 加載 SSL 憑證
const sslOptions = {
    key: fs.readFileSync(path.join(__dirname, 'key.pem')),
    cert: fs.readFileSync(path.join(__dirname, 'cert.pem')),
  };

app.get('/', (req, res) =>{
    res.send('Hello World');
});
const mongoUri = 'mongodb://localhost:27017'; // 替換為你的 MongoDB 連接字串
const mongoClient = new MongoClient(mongoUri);


// 中间件获取原始请求体
app.use(express.json({ verify: (req, res, buf) => { req.rawBody = buf.toString(); }}));
app.post('/webhook', line.middleware(config), (req, res) =>{
    // 获取签名
    const signature = req.headers['x-line-signature'];
    console.log('X-Line-Signature:', signature);

    // 计算签名
    const hash = crypto.createHmac('sha256', config.channelSecret)
    .update(req.rawBody)
    .digest('base64');


    // 比较签名
    if (hash !== signature) {
    return res.status(401).send('Signature validation failed');
    }
    // 处理每个来自LINE的事件
    Promise
    .all(req.body.events.map(handleEvent))    
    .then((result) => res.json(result))
    .catch((err) => {
        console.error(err);
        res.status(500).end();
    });
});
async function handleEvent(event) {

    if (event.type !== 'message' || (event.message.type !== 'text' && event.message.type !== 'image')) {
        return Promise.resolve(null);
    }

    const userId = event.source.userId;

    if (!currentTask[userId]) {
        currentTask[userId] = { material: null, manufacturer: null, items: [], currentStep: 'start' };
    }

    if (event.type === 'message' && event.message.type === 'text') {
        const text = event.message.text.trim();

        if (event.message.text === '#自檢表') {
            currentTask[userId].currentStep = 'awaiting_material';
            return replyMessage(event.replyToken, '請輸入檢查料號，如:N01' );
        }

        if (currentTask[userId].currentStep === 'awaiting_material') {
            currentTask[userId].material = text;
             // 查詢資料庫中的廠商資訊
            const manufacturer = await getManufacturerByMaterialCode(text);
            if (manufacturer) {
                currentTask[userId].manufacturer = manufacturer;
                currentTask[userId].currentStep = 'awaiting_items';
                return replyMessage(event.replyToken, `料號對應的廠商是：${manufacturer}。請輸入檢查項目和合格狀態（如：「平整檢查，ok」或「R值檢查，x」）。完成所有項目填寫，請輸入「#填寫完成」`);
            } else {
                return replyMessage(event.replyToken, '未找到對應的廠商，請檢查料號是否正確。');
            }
        }
        
        if (currentTask[userId].currentStep === 'awaiting_items') {
            if(text === '#填寫完成'){
                if (currentTask[userId].items.length === 0) {
                    return replyMessage(event.replyToken, '您尚未輸入任何檢查項目，請繼續輸入或重新開始。');
                }
                currentTask[userId].currentStep = 'awaiting_images';
                return replyMessage(event.replyToken, '請開始為每個項目上傳照片，依次傳送所有照片。完成所有照片上傳後，請輸入「#完成」。');
            }else{
                const parts = text.split('，');
                if (parts.length < 2) {
                    return replyMessage(event.replyToken, '輸入格式不正確，請使用格式：「項目說明，合格狀態」。');
                }

                const description = parts.slice(0, -1).join('，').trim();
                const status = parts[parts.length - 1].trim().toLowerCase();
                if (status !== 'ok' && status !== 'x' && status !== '合格' && status !== '不合格') {
                    return replyMessage(event.replyToken, '合格狀態必須是「ok」、「x」、「合格」或「不合格」。');
                }

                currentTask[userId].items.push({
                    description: description,
                    status: status === 'ok' || status === '合格' ? '合格' : '不合格',
                    images: []
                });

                return replyMessage(event.replyToken, '已記錄該項目，請繼續輸入下一個檢查項目，或輸入「#填寫完成」。');
            }
        }
        if (currentTask[userId].currentStep === 'awaiting_images' && text === '#完成') {
            return sendExcelFile(userId, event.replyToken);
        }
        
    }

    if (event.type === 'message' && event.message.type === 'image' && currentTask[userId].currentStep === 'awaiting_images') {
        return client.getMessageContent(event.message.id)
            .then(stream => {
                const buffer = [];
                stream.on('data', chunk => buffer.push(chunk));
                stream.on('end', () => {
                    const lastItem = currentTask[userId].items[currentTask[userId].items.length - 1];
                    lastItem.images.push(Buffer.concat(buffer));
                    return replyMessage(event.replyToken, '照片已接收。繼續上傳下一張照片，或輸入「#完成」結束上傳。');
                });
            })
            .catch(err => {
                console.error('getMessageContent error:', err);
                return client.replyMessage(event.replyToken, { type: 'text', text: '圖片下載失敗，請稍後再試。' });
            });
    }
    
    return Promise.resolve(null);
}

// 生成 Excel 文件并保存到本地
function generateExcel(userId) {
    const workbook = new exceljs.Workbook();
    const worksheet = workbook.addWorksheet('自檢表');
    
    worksheet.getColumn(1).width = 30; // Column A: Item Number
    worksheet.getColumn(2).width = 15; // Column B: Pass/Fail Status
    worksheet.getColumn(3).width = 30; // Column C: Notes
    // 添加標題信息：廠商、料號、檢表日期
    worksheet.mergeCells('A1:B1');
    worksheet.getCell('A1').value = `廠商: ${currentTask[userId].manufacturer}`;
    worksheet.mergeCells('C1:D1');
    worksheet.getCell('C1').value = `料號: ${currentTask[userId].material}`;
    worksheet.mergeCells('E1:F1');
    worksheet.getCell('E1').value = `日期: ${new Date().toLocaleDateString()}`;
    worksheet.addRow([]); // Blank row

    worksheet.addRow(['項目', '合/不合格', '備註']);
    worksheet.getRow(3).font = { bold: true };

    let currentRow = 4;
    currentTask[userId].items.forEach(item => {
        // 插入檢查項目和合格狀態
        worksheet.addRow([item.description, item.status, '']);
        currentRow++;// 移動到下一行
    });

    //新增圖片之前保留空白行
    currentRow++;
    worksheet.addRow([]);
    currentRow++;

    let imageCol = 1;//以類似網格的格式新增影像（每行兩個影像）
    currentTask[userId].items.forEach(item => {
        item.images.forEach(image => {
            const imageId = workbook.addImage({
                buffer: image,
                extension: 'png'
            });

            worksheet.addImage(imageId, {
                tl: { col: imageCol - 1, row: currentRow - 1 },
                ext: { width: 200, height: 150 } // 圖片尺寸
            });

            if (imageCol === 2) {
                // 移至兩張影像後的下一行
                imageCol = 1;
                currentRow += 8; // 調整下一組影像的行位置
            } else {
                imageCol++; // 移至下一列
            }
        });
    });

    const filePath = path.join(fileDirectory, '自檢表.xlsx');
    return workbook.xlsx.writeFile(filePath).then(() => filePath);
}
// 設置静態文件路径，讓文件可以通过 HTTP 访问
app.use('/files', express.static(fileDirectory));

// 回傳 Excel 文件
function sendExcelFile(userId, replyToken) {
    return generateExcel(userId).then(filePath => {
        const fileUrl = `https://039c-60-248-110-97.ngrok-free.app/files/自檢表.xlsx`;

        return client.replyMessage(replyToken, {
            type: 'text',
            text: `文件` +fileUrl
        });
    }).catch(err => {
        console.error('Error generating Excel:', err);
        return client.replyMessage(replyToken, { type: 'text', text: '生成 Excel 文件時出錯，請稍後再試。' });
    });
}

async function getManufacturerByMaterialCode(materialCode) {
    try {
        await mongoClient.connect();
        const database = mongoClient.db('Self_Inspection'); // 替換為你的資料庫名稱
        const materials = database.collection('MaterialsFactory');

        const material = await materials.findOne({ materialCode: materialCode });
        return material ? material.manufacturer : null;
    } catch (err) {
        console.error('Database query error:', err);
        return null;
    } finally {
        await mongoClient.close();
    }
}
  

// 回覆訊息函數
function replyMessage(replyToken, message) {

    const body = {
      replyToken: replyToken,
      messages: [
        {
          type: 'text',
          text: message
        }
      ]
    };
  
    axios.post('https://api.line.me/v2/bot/message/reply', body, { headers })
      .then(() => {
        console.log('Reply message sent successfully');
      })
      .catch(error => {
        console.error('Error sending reply message:', error);
      });
  }
// 建立 HTTPS 伺服器
https.createServer(sslOptions, app).listen(23458, () => {
    console.log('HTTPS server running on port 23458');
  });