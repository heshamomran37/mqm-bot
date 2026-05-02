require('dotenv').config();
const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcodeTerminal = require('qrcode-terminal');
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { GoogleGenAI } = require('@google/genai');

// Initialize Gemini AI
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Load services data
const data = JSON.parse(fs.readFileSync('./services.json', 'utf8'));

// Initialize Excel Leads File
const LEADS_FILE = './mqm_leads.xlsx';
async function initExcel() {
              if (!fs.existsSync(LEADS_FILE)) {
                                const workbook = new ExcelJS.Workbook();
                                const sheet = workbook.addWorksheet('Leads');
                                sheet.columns = [
                                          { header: 'Date', key: 'date', width: 20 },
                                          { header: 'Name', key: 'user', width: 30 },
                                          { header: 'Phone', key: 'phone_provided', width: 20 },
                                          { header: 'Service', key: 'service', width: 30 },
                                          { header: 'Full Msg', key: 'full_msg', width: 50 }
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

// User States to track conversation flow
const userStates = {};

// Helper for emojis
function getEmoji(id) {
              return '>>';
}

// Gemini Response Function
async function getGeminiResponse(userMsg) {
              try {
                                const systemPrompt = `
                                You are a "Senior Digital Solutions Consultant" at MQM_IT. 
                                Your name is "MQM Smart Assistant". You must act as a professional, human employee.

                                LANGUAGE RULES:
                                - Detect the user's language. If they speak Arabic, respond in professional, friendly Arabic.
                                - If they speak English, respond in professional, clear, and business-oriented English.
                                - Do not mix languages unless requested.

                                MQM_IT CONTEXT:
                                - Business Name: \${data.business_name}
                                - Core Expertise: Full Digital Solutions (Software Development, Digital Marketing, Graphic Design, SEO, Video Production).
                                - Website: \${data.contact.website}

                                SERVICES DATA:
                                \${data.services.map(s => \`- \${s.name}: \${s.description}\`).join('\\n')}

                                INTERACTION GUIDELINES:
                                1. GREETING: Always start with a warm, professional greeting.
                                2. TONE: Be helpful, persuasive, and innovative. You are an expert solving client problems.
                                3. PRICING: If asked about price, explain that we provide "Custom-tailored solutions" and we need to understand their project requirements first to give the "Best Value Quote". Encourage them to press 0 for a direct consultant call.
                                4. SALES STRATEGY: Focus on the benefits (ROI, growth, professional image) not just technical features.
                                5. CALL TO ACTION (CTA): End every message with a suggestion (e.g., "Would you like to see our portfolio? Press 9" or "Shall I help you place an order for this service?").
                                6. NO AI DISCLOSURE: Never state you are an AI. You are a member of the MQM_IT technical & sales team.
                                7. WHATSAPP FORMATTING: Use bold text (*text*) and symbols sparingly for better readability. Keep paragraphs concise.
                                `;

                  const response = await ai.models.generateContent({
                                        model: 'models/gemini-flash-latest',
                                        systemInstruction: systemPrompt,
                                        contents: [{ role: 'user', parts: [{ text: userMsg }] }]
                  });

                  return response.text;
              } catch (error) {
                                console.error('Gemini Error:', error);
                                return null;
              }
}

// Initialize WhatsApp client
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
              console.log('--- QR CODE ---');
              qrcodeTerminal.generate(qr, { small: true });
              try {
                                await QRCode.toFile('./qr_code.png', qr);
                                console.log('QR Code saved as image in: qr_code.png');
              } catch (err) { console.error(err); }
});

client.on('ready', () => {
              console.log('Connected! The upgraded bot is ready from MQM_IT');
});

client.on('message', async msg => {
              let chat;
              try {
                                chat = await msg.getChat();
              } catch (e) {
                                console.log('Warn: Could not get chat object, continuing...');
              }
              const userMessage = msg.body.toLowerCase().trim();
              const userId = msg.from;

              console.log(`[Message] from ${userId}: ${msg.body}`);

              // If user is in a state of providing phone number
              if (userStates[userId] && userStates[userId].state === 'AWAITING_PHONE') {
                                await addLead({
                                                      from: userId,
                                                      phone: msg.body,
                                                      service: userStates[userId].service,
                                                      body: msg.body
                                });
                                delete userStates[userId];
                                await client.sendMessage(msg.from, 'Data received! The *MQM_IT* team will contact you soon. Thank you!');
                                return;
              }

              // Triggers: Greetings or start
              const greetings = ['hello', 'hi', 'hey', 'start', '.', '?'];
              const isGreeting = greetings.some(g => userMessage.includes(g));

              if (isGreeting) {
                                console.log('Sending professional welcome message...');
                                const welcomeMsg = `* ${data.business_name} Digital Solutions *\n\n` +
                                                      `Welcome! We are here to take your business to the next level.\\n` +
                                                      `Please select a service number:\\n\\n` +
                                                      data.services.map(s => `- *${s.name}*`).join('\\n') +
                                                      `\\n\\n9. Our Portfolio\\n` +
                                                      `0. Talk to a representative\\n\\n` +
                                                      `_MQM_IT Team is always at your service_`;

                  await client.sendMessage(msg.from, welcomeMsg);
                                return;
              }

              // Logic for Portfolio
              if (userMessage === '9' || userMessage.includes('portfolio')) {
                                const portfolioPath = path.join(__dirname, 'portfolio', 'mqm_portfolio.png');
                                if (fs.existsSync(portfolioPath)) {
                                                      const media = MessageMedia.fromFilePath(portfolioPath);
                                                      await client.sendMessage(msg.from, media, { caption: 'Samples of our work at *MQM_IT*. We create excellence!' });
                                } else {
                                                      await client.sendMessage(msg.from, 'Portfolio is being updated, please contact support for details.');
                                }
                                return;
              }

              // Logic for service details
              const selectedService = data.services.find(s => s.id === userMessage);
              if (selectedService) {
                                const response = `*${selectedService.name}*\\n\\n${selectedService.description}\\n\\nWould you like to order this service now? (Reply with "order" or "10")`;
                                await client.sendMessage(msg.from, response);
                                userStates[userId] = { service: selectedService.name }; // Save context
                  return;
              }

              // Ordering logic
              if (userMessage === 'order' || userMessage === '10') {
                                userStates[userId] = { ...userStates[userId], state: 'AWAITING_PHONE' };
                                await client.sendMessage(msg.from, 'Great! Please send your phone number so we can discuss the details.');
                                return;
              }

              if (userMessage === '0') {
                                await client.sendMessage(msg.from, 'You will be connected to a representative now. Thank you for choosing MQM_IT.');
                                return;
              }

              // --- NEW: Gemini AI Catch-All Response ---
              console.log('Unknown message, consulting Gemini...');
              const aiResponse = await getGeminiResponse(msg.body);

              if (aiResponse) {
                                await client.sendMessage(msg.from, aiResponse);
              } else {
                                const defaultMsg = `Sorry, I didn't quite understand that. 
                                        const defaultMsg = `Sorry, I didn't quite understand that. ???\n\nI am the *MQM_IT* smart bot, you can always pick a number:\n\n` +
                                                      data.services.map(s => `- *${s.name}*`).join('\n') +
                                                      `\n\n9. Our Portfolio\n` +
                                                      `0. Talk to a representative`;
                                await client.sendMessage(msg.from, defaultMsg);
              }
});

client.initialize();
