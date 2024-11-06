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
        currentTask[userId] = { images: [], currentStep: 'start' };
    }

    if (event.message.type === 'text') {
        if (event.message.text === '#自檢表') {
            currentTask[userId].currentStep = 'awaiting_images';
            return replyMessage(event.replyToken, '請上傳一個或多個圖片。' );
        }
        if (event.message.text === '#完成') {
            if (currentTask[userId].images.length === 0) {
                return replyMessage(event.replyToken,'沒有足夠的圖片和信息來生成自檢表，請重新開始。' );
            }
            console.log('Generating Excel for user:', userId); // Debug log
            return sendExcelFile(userId, event.replyToken);
        }
        if (currentTask[userId].currentStep === 'awaiting_info') {
            const [description, status, note] = event.message.text.split('，').map(s => s.trim());
            const lastImage = currentTask[userId].images[currentTask[userId].images.length - 1];
            if (lastImage) {
                lastImage.info = { description, status, note };
                console.log('Recorded image info:', lastImage.info); // Debug log

                return replyMessage(event.replyToken,'圖片資訊已記錄。請繼續上傳下一張圖片，或輸入 `#完成` 結束並生成自檢表。');
            }
        }
        
    }
    if (event.message.type === 'image') {
        if (currentTask[userId].currentStep === 'awaiting_images' || currentTask[userId].currentStep === 'awaiting_info') {
            return client.getMessageContent(event.message.id)
                .then(stream => {
                    const buffer = [];
                    stream.on('data', chunk => buffer.push(chunk));
                    stream.on('end', () => {
                        currentTask[userId].images.push({ buffer: Buffer.concat(buffer) });
                        currentTask[userId].currentStep = 'awaiting_info';
                        console.log('Image uploaded, awaiting info.');
                        client.replyMessage(event.replyToken, { type: 'text', text: '請輸入圖片說明、合格或不合格以及備註（格式：說明, 合格/不合格, 備註）。' });
                    });
                })
                .catch(err => {
                    console.error('getMessageContent error:', err);
                    return client.replyMessage(event.replyToken, { type: 'text', text: '圖片下載失敗，請稍後再試。' });
                });
        }
    }
    return Promise.resolve(null);
}

// 生成 Excel 文件并保存到本地
function generateExcel(userId) {
    const workbook = new exceljs.Workbook();
    const worksheet = workbook.addWorksheet('自檢表');

    worksheet.columns = [
        { header: '圖片', key: 'image', width: 65},
        { header: '圖片說明', key: 'description', width: 30  },
        { header: '合格/不合格', key: 'status', width: 15  },
        { header: '備註', key: 'note', width: 15  }
    ];

    // 假设 currentTask[userId].images 存储了图片的 Buffer 数据
    currentTask[userId].images.forEach((img, index) => {
        // 在单元格中插入图片
        const imageId = workbook.addImage({
            buffer: img.buffer, // 图片的 Buffer 数据
            extension: 'png'    // 或者是 'png' 或其他支持的格式
        });
         // 设置图片的大小和位置
         worksheet.addImage(imageId, {
            tl: { col: 0, row: index + 2 }, // 将图片插入到第一列
            ext: { width: 500, height: 500 } // 设置图片的尺寸
        });
        // 添加其他列的文本数据
        worksheet.addRow({
            description: img.info.description,
            status: img.info.status,
            note: img.info.note
        });
    });

    const filePath = path.join(fileDirectory, '自檢表.xlsx');
    return workbook.xlsx.writeFile(filePath).then(() => filePath);
}
// 设置静态文件路径，让文件可以通过 HTTP 访问
app.use('/files', express.static(fileDirectory));

// 发送 Excel 文件
function sendExcelFile(userId, replyToken) {
    return generateExcel(userId).then(filePath => {
        const fileUrl = `https://3e58-60-248-110-97.ngrok-free.app/files/自檢表.xlsx`;

        return client.replyMessage(replyToken, {
            type: 'text',
            text: `文件` +fileUrl
        });
    }).catch(err => {
        console.error('Error generating Excel:', err);
        return client.replyMessage(replyToken, { type: 'text', text: '生成 Excel 文件時出錯，請稍後再試。' });
    });
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