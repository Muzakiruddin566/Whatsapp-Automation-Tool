const express = require("express");
const app = express();
const xlsx = require('xlsx');
const bodyParser = require("body-parser");
const webdriver = require('selenium-webdriver');
const multer = require("multer");
const path = require("path");   
const { By, Key } = webdriver;
const upload = multer({ dest: 'uploads/phoneNumbers/' }); // Specify the destination folder for uploaded files
const PORT = 5000;
// set the view engine to ejs
app.set('view engine', 'ejs');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended:true}));
let driver;

async function sendMessagesFromExcel(message , pdf, pdf2) {
    const workbook = xlsx.readFile(pdf);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(worksheet);

    for (const row of rows) {
        const contactNumber = row['Phone'];
        const name = row['Name']; // Assuming 'Name' contains the message
        console.log("Number", contactNumber);
        console.log("Name", name);

        if (contactNumber && message) {
            try {
                 await sendWhatsAppMessage(contactNumber, message, pdf2);
            } catch (error) {
                console.log(error);
            }
        }
    }
}


async function pauseNow(Time = 5) {
    return new Promise((resolve, reject) => {
        setTimeout(() => resolve("Done"), Time * 1000);
    });
}
    


async function sendWhatsAppMessage(contactNumber, message, pdf2) {
    return new Promise(async (resolve, reject) => {
        try {
            // Open WhatsApp Web with the specified phone number
            const response =  await driver.get(`https://web.whatsapp.com/send?phone=${contactNumber}`);
                await driver.wait(webdriver.until.elementLocated(By.css('div[title="Type a message"][contenteditable="true"]'), 10000));
                const messageInput = await driver.findElement(By.css('div[title="Type a message"][contenteditable="true"]'));
                messageInput.sendKeys(message);
                await attachFile(pdf2);

                // Click the send button
                await pressSend();

                console.log('Message sent successfully.');
                await pauseNow();
                resolve('Message sent successfully');
            // }
        } catch (error) {
            if (error.name === 'UnexpectedAlertOpenError') {
                // Handle the alert here, for example, by dismissing it
                const alert = await driver.switchTo().alert();
                console.log("alert = = ==== ", alert);
                await alert.dismiss();
                console.log('Dismissed unexpected alert:', alert.getText());
                await pauseNow();
                // You can also add further logic here to verify and continue the process.
                resolve('Unexpected alert dismissed.');
            } else {
                console.error("An error occurred while sending the message:", error);
                await pauseNow();
                reject(error);
            }
        }
    });
}

async function attachFile(filePath) {
    const absoluteFilePath = path.join(__dirname, filePath);
    console.log("path", absoluteFilePath);
    const attachButton = await driver.findElement(By.css('div[title="Attach"]'));
    await attachButton.click();

    const input = await driver.findElement(By.css('input[type="file"]'));
    await input.sendKeys(absoluteFilePath);

    // Wait for the file to be attached
    await driver.wait(webdriver.until.elementLocated(By.css('.copyable-text.selectable-text')), 10000);
}





async function pressSend() {
    return new Promise(async (resolve, reject) => {
        try {
            const sendButton = await driver.wait(webdriver.until.elementLocated(By.css('div[aria-label="Send"][role="button"]')), 10000);            
            await driver.wait(webdriver.until.elementIsNotVisible(driver.findElement(By.css('.element-that-covers-send-button'))), 5000).catch(() => {});
            await sendButton.click();
            console.log('Send button clicked successfully.');
            await pauseNow(10);
            resolve('Send button clicked successfully.');
        } catch (error) {
            console.error("An error occurred while clicking the send button:", error);
            reject(error);
        }
    });
}




async function AutomatedWhatsappMessage(message, pdf, pdf2) {
    await pauseNow(10);
    await sendMessagesFromExcel(message, pdf, pdf2);
}

async function main(message = "", pdf, pdf2) {
    await driver.get('https://web.whatsapp.com/');
    await pauseNow(20);
    AutomatedWhatsappMessage(message , pdf, pdf2)

   
}

app.get("/", (req, res) => {
    res.render("pages/index");
})

app.post("/result", upload.fields([{name : "pdf"}, {name : "pdf2"}]), (req, res) => {
    console.log("Result MEssage", req.body);
    console.log("file", req.files['pdf2']);
    if (req.files['pdf'] && req.files['pdf2']) {
        // const filePath = req.file.path;
        console.log("Running");
        const pdf1Files = req.files['pdf'].map(file => file.path);
        const pdf2Files = req.files['pdf2'].map(file => file.path);
        
        driver =  new webdriver.Builder().forBrowser('chrome').build();
        main(req.body.message || "" , pdf1Files[0], pdf2Files[0]);
    }
})



app.listen(PORT, () => {
    console.log(`http://localhost:${PORT}/`);
    
})

