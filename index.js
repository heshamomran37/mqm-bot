require('dotenv').config();
const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcodeTerminal = require('qrcode-terminal');
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { GoogleGenAI } = require('@google/genai');

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
const data = JSON.parse(fs.readFileSync('./services.json', 'utf8'));
const LEADS_FILE = './mqm_leads.xlsx';

async function initExcel() {
      if (!fs.existsSync(LEADS_FILE)) {
                const workbook = new ExcelJS.Workbook();
                const sheet = workbook.addWorksheet('Leads');
                sheet.columns = [
                  { header: '\u0627\u0644\u062a\u0627\u0631\u064a\u062e', key: 'date', width: 20 },
                  { header: '\u0627\u0644\u0627\u0633\u0645/\u0627\u0644\u0631\u0642\u0645 \u0627\u0644\u0645\u0641\u062a\u0627\u062d', key: 'user', width: 30 },
                  { header: '\u0631\u0642\u0645 \u0627\u0644\u0647\u0627\u062a\u0641 \u0627\u0644\u0645\u062a\u0631\u0648\u0643', key: 'phone_provided', width: 20 },
                  { header: '\u0627\u0644\u062e\u062f\u0645\u0629 \u0627\u0644\u0645\u0647\u062a\u0645 \u0628\u0647\u0627', key: 'service', width: 30 },
                  { header: '\u0627\u0644\u0631\u0633\u0627\u0644\u0629 \u0628\u0627\u0644\u0643\u0627\u0645\u0644', key: 'full_msg', width: 50 }
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
                date: new Date().toLocaleString('ar-EG'),
                user: userData.from,
                phone_provided: userData.phone || 'N/A',
                service: userData.service || 'N/A',
                full_msg: userData.body
      });
      await workbook.xlsx.writeFile(LEADS_FILE);
}

const userStates = {};

function getEmoji(id) {
      const emojis = {
                '1': '1\ufe0f\u20e3', '2': '2\ufe0f\u20e3', '3': '3\ufe0f\u20e3', '4': '4\ufe0f\u20e3', '5': '5\ufe0f\u20e3', '6': '6\ufe0f\u20e3', '7': '7\ufe0f\u20e3', '8': '8\ufe0f\u20e3', '9': '9\ufe0f\u20e3', '0': '0\ufe0f\u20e3'
      };
      return emojis[id] || '\ud83d\udd39';
}

async function getGeminiResponse(userMsg) {
      try {
                const systemPrompt = \`You are a consultant.\`;
                        const response = await ai.models.generateContent({
                                    model: 'models/gemini-flash-latest',
                                                systemInstruction: systemPrompt,
                                                            contents: [{ role: 'user', parts: [{ text: userMsg }] }]
                                                                    });
                                                                            return response.text;
                                                                                } catch (error) { return null; }
                                                                                }

                                                                                const client = new Client({
                                                                                    authStrategy: new LocalAuth(),
                                                                                        puppeteer: {
                                                                                                executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined,
                                                                                                        handleSIGINT: false,
                                                                                                                args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-accelerated-2d-canvas', '--no-first-run', '--no-zygote', '--disable-gpu']
                                                                                                                    }
                                                                                                                    });
                                                                                                                    
                                                                                                                    client.on('qr', async (qr) => {
                                                                                                                        console.log('--- QR CODE ---');
                                                                                                                            qrcodeTerminal.generate(qr, { small: true });
                                                                                                                                try {
                                                                                                                                        await QRCode.toFile('./qr_code.png', qr);
                                                                                                                                                console.log('\u062a\u0645 \u062d\u0641\u0638 \u0627\u0644\u0631\u0645\u0632 \u0643\u0635\u0648\u0631\u0629 \u0641\u064a: qr_code.png');
                                                                                                                                                    } catch (err) { console.error(err); }
                                                                                                                                                    });
                                                                                                                                                    
                                                                                                                                                    client.on('ready', () => {
                                                                                                                                                        console.log('\u062a\u0645 \u0627\u0644\u0631\u0628\u0637 \u0628\u0646\u062c\u0627\u062d!');
                                                                                                                                                        });
                                                                                                                                                        
                                                                                                                                                        client.on('message', async msg => {
                                                                                                                                                            const userId = msg.from;
                                                                                                                                                                const userMessage = msg.body.toLowerCase().trim();
                                                                                                                                                                
                                                                                                                                                                    if (userStates[userId] && userStates[userId].state === 'AWAITING_PHONE') {
                                                                                                                                                                            await addLead({ from: userId, phone: msg.body, service: userStates[userId].service, body: msg.body });
                                                                                                                                                                                    delete userStates[userId];
                                                                                                                                                                                            await client.sendMessage(userId, '\u062a\u0645 \u0627\u0633\u062a\u0644\u0627\u0645 \u0628\u064a\u0627\u0646\u0627\u062a\u0643 \u0628\u0646\u062c\u0627\u062d!');
                                                                                                                                                                                                    return;
                                                                                                                                                                                                        }
                                                                                                                                                                                                        
                                                                                                                                                                                                            const greetings = ['hello', '\u0633\u0644\u0627\u0645', '\u0645\u0631\u062d\u0628\u0627', 'hi'];
                                                                                                                                                                                                                if (greetings.some(g => userMessage.includes(g))) {
                                                                                                                                                                                                                        const welcomeMsg = \`*\${data.business_name}*\n\n\` + data.services.map(s => \`\${getEmoji(s.id)} *\${s.name}*\`).join('\n');
                                                                                                                                                                                                                                await client.sendMessage(userId, welcomeMsg);
                                                                                                                                                                                                                                        return;
                                                                                                                                                                                                                                            }
                                                                                                                                                                                                                                            
                                                                                                                                                                                                                                                const selectedService = data.services.find(s => s.id === userMessage);
                                                                                                                                                                                                                                                    if (selectedService) {
                                                                                                                                                                                                                                                            await client.sendMessage(userId, \`*\${selectedService.name}*\n\${selectedService.description}\n\u0637\u0644\u0628\u061f\`);
                                                                                                                                                                                                                                                                    userStates[userId] = { service: selectedService.name };
                                                                                                                                                                                                                                                                            return;
                                                                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                    if (userMessage === '\u0637\u0644\u0628' || userMessage === '10') {
                                                                                                                                                                                                                                                                                            userStates[userId] = { ...userStates[userId], state: 'AWAITING_PHONE' };
                                                                                                                                                                                                                                                                                                    await client.sendMessage(userId, '\u0631\u0642\u0645 \u0627\u0644\u0647\u0627\u062a\u0641\u061f');
                                                                                                                                                                                                                                                                                                            return;
                                                                                                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                    const aiResponse = await getGeminiResponse(msg.body);
                                                                                                                                                                                                                                                                                                                        if (aiResponse) await client.sendMessage(userId, aiResponse);
                                                                                                                                                                                                                                                                                                                        });
                                                                                                                                                                                                                                                                                                                        
                                                                                                                                                                                                                                                                                                                        client.initialize();
                                                                                                                                                                                                                                                                                                                        
