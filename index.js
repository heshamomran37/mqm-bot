require('dotenv').config();
const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcodeTerminal = require('qrcode-terminal');
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const data = JSON.parse(fs.readFileSync('./services.json', 'utf8'));

const LEADS_FILE = './mqm_leads.xlsx';
async function initExcel() {
      if (!fs.existsSync(LEADS_FILE)) {
                const workbook = new ExcelJS.Workbook();
                const sheet = workbook.addWorksheet('Leads');
                sheet.columns = [
                  { header: 'Date', key: 'date', width: 20 },
                  { header: 'User', key: 'user', width: 30 },
                  { header: 'Phone', key: 'phone_provided', width: 20 },
                  { header: 'Service', key: 'service', width: 30 },
                  { header: 'Message', key: 'full_msg', width: 50 }
                          ];
                await workbook.xlsx.writeFile(LEADS_FILE);
      }
}
initExcel();

async function addLead(userData) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(LEADS_FILE);
      const sheet = workbook.getWorksheet('Leads');
      sheet.addRow({
                date: new Date().toLocaleString(),
                user: userData.from,
                phone_provided: userData.phone || 'N/A',
                service: userData.service || 'N/A',
                full_msg: userData.body
      });
      await workbook.xlsx.writeFile(LEADS_FILE);
}

const userStates = {};

function getEmoji(id) {
      return id;
}

const client = new Client({
      authStrategy: new LocalAuth(),
      puppeteer: {
                executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined,
                handleSIGINT: false,
                args: [
                              '--no-sandbox',
                              '--disable-setuid-sandbox',
                              '--disable-dev-shm-usage',
                              '--disable-accelerated-2d-canvas',
                              '--no-first-run',
                              '--no-zygote',
                              '--disable-gpu'
                          ]
      }
});

client.on('qr', async (qr) => {
      console.log('QR Code generated');
      qrcodeTerminal.generate(qr, { small: true });
      try {
                await QRCode.toFile('./qr_code.png', qr);
      } catch (err) { console.error(err); }
});

client.on('ready', () => {
      console.log('Bot ready');
});

client.on('message', async msg => {
      let chat;
      try {
                chat = await msg.getChat();
      } catch (e) {}

              const userMessage = msg.body.toLowerCase().trim();
      const userId = msg.from;

              if (userStates[userId] && userStates[userId].state === 'AWAITING_PHONE') {
                        await addLead({
                                      from: userId,
                                      phone: msg.body,
                                      service: userStates[userId].service,
                                      body: msg.body
                        });
                        delete userStates[userId];
                        await client.sendMessage(msg.from, 'Success! Team will contact you.');
                        return;
              }

              const greetings = ['hello', 'hi', 'hey', 'start', '.', '?'];
      if (greetings.some(g => userMessage.includes(g))) {
                const welcomeMsg = 'Welcome to MQM_IT\n' +
                              'Please choose service:\n' +
                              data.services.map(s => s.id + ' ' + s.name).join('\n') +
                              '\n9 Portfolio\n0 Talk to us';
                await client.sendMessage(msg.from, welcomeMsg);
                return;
      }

              if (userMessage === '9') {
                        const portfolioPath = path.join(__dirname, 'portfolio', 'mqm_portfolio.png');
                        if (fs.existsSync(portfolioPath)) {
                                      const media = MessageMedia.fromFilePath(portfolioPath);
                                      await client.sendMessage(msg.from, media, { caption: 'Our Portfolio' });
                        } else {
                                      await client.sendMessage(msg.from, 'Portfolio not available.');
                        }
                        return;
              }

              const selectedService = data.services.find(s => s.id === userMessage);
      if (selectedService) {
                await client.sendMessage(msg.from, selectedService.name + '\n' + selectedService.description + '\nOrder? Reply "order"');
                userStates[userId] = { service: selectedService.name };
                return;
      }

              if (userMessage === 'order') {
                        userStates[userId] = { ...userStates[userId], state: 'AWAITING_PHONE' };
                        await client.sendMessage(msg.from, 'Send your phone number.');
                        return;
              }

              if (userMessage === '0') {
                        await client.sendMessage(msg.from, 'Connecting...');
                        return;
              }

              await client.sendMessage(msg.from, 'Choose a number from list.');
});

client.initialize();
