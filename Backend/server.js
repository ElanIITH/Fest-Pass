const { google } = require('googleapis');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const handlebars = require('handlebars');
const bwipjs = require('bwip-js');
const axios = require('axios');
require('dotenv').config();

// Configure Google Sheets API
const auth = new google.auth.GoogleAuth({
    keyFile: 'elan-pass-mailing-450613-dbf3999bb350.json', // Your Google Cloud service account key
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
});

// Configure email transporter
const transporter = nodemailer.createTransport({
    host: 'smtp.gmail.com',
    port: 587,
    secure: false, // Use TLS
    auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_APP_PASSWORD
    },
    tls: {
        rejectUnauthorized: false
    }
});

// Add a verification step before processing
async function verifyEmailConfig() {
    try {
        await transporter.verify();
        console.log("Email configuration verified successfully");
        return true;
    } catch (error) {
        console.error("Email verification failed:", error);
        return false;
    }
}

// Read HTML template
const templatePath = path.join(__dirname, '..', 'pass.html');
const template = handlebars.compile(fs.readFileSync(templatePath, 'utf8'));

async function generateBarcode(email) {
    try {
        const barcodeText = `ELAN_24_${email}`;
        
        // Generate barcode as PNG buffer
        const png = await new Promise((resolve, reject) => {
            bwipjs.toBuffer({
                bcid: 'code128',       // Barcode type
                text: barcodeText,     // Text to encode
                scale: 3,              // 3x scaling factor
                height: 10,            // Bar height, in millimeters
                includetext: true,     // Show human-readable text
                textxalign: 'center',  // Center the text
            }, function (err, png) {
                if (err) {
                    reject(err);
                } else {
                    resolve(png);
                }
            });
        });

        // Convert to base64 for embedding in HTML
        return `data:image/png;base64,${png.toString('base64')}`;
    } catch (err) {
        console.error('Error generating barcode:', err);
        throw err;
    }
}

async function sendPass(participant) {
    try {
        // Generate barcode
        const barcodeBuffer = await new Promise((resolve, reject) => {
            bwipjs.toBuffer({
                bcid: 'code128',
                text: `ELAN_24_${participant.email}`,
                scale: 3,
                height: 10,
                includetext: true,
                textxalign: 'center',
                backgroundcolor: 'FFFFFF', // Add white background
                padding: 10 // Add some padding around the barcode
            }, function (err, png) {
                if (err) reject(err);
                else resolve(png);
            });
        });

        // Read header image
        let headerBuffer;
        try {
            const headerImagePath = path.join(__dirname, '..', 'header.png');
            headerBuffer = fs.readFileSync(headerImagePath);
        } catch (error) {
            console.warn('Header image not found:', error.message);
            headerBuffer = null;
        }

        // Generate HTML content with CID references
        const htmlContent = template({
            Name: participant.name,
            Pass: participant.passType,
            ALT: `ELAN_24_${participant.email}`,
            barcode: 'cid:barcodeImage', // Reference to content ID
            College: participant.college,
            City: participant.city,
            headerImage: headerBuffer ? 'cid:headerImage' : null, // Reference to content ID
            useHeaderFallback: !headerBuffer
        });

        // Configure email with attachments using content IDs
        const mailOptions = {
            from: process.env.EMAIL_USER,
            to: participant.email,
            subject: 'Booking confirmed | Papon Live at IIT Hyderabad | Elan & nVision Fest Pass',
            html: htmlContent,
            attachments: [
                {
                    filename: 'barcode.png',
                    content: barcodeBuffer,
                    cid: 'barcodeImage' // Content ID for barcode
                }
            ]
        };

        // Add header image attachment if available
        if (headerBuffer) {
            mailOptions.attachments.push({
                filename: 'header.png',
                content: headerBuffer,
                cid: 'headerImage' // Content ID for header
            });
        }

        await transporter.sendMail(mailOptions);
        console.log(`Pass sent successfully to ${participant.email}`);
        return true;
    } catch (error) {
        console.error(`Error sending pass to ${participant.email}:`, error);
        return false;
    }
}

async function processRegistrations() {
    try {
        const sheets = google.sheets({ version: 'v4', auth });
        
        // Updated range to match your sheet structure
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: '1QRHIHgXVMGebIX_3YeFLE1joCt-HzDO4tADA0SmeEv0',
            range: 'Sheet1!A2:H', // Updated to include all columns
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) {
            console.log('No data found.');
            return;
        }

        // Process each registration with updated column indices
        for (const row of rows) {
            const participant = {
                timestamp: row[0],
                name: row[1],        // Name is in column 2
                phone: row[2],       // Phone is in column 3
                email: row[3],       // Email is in column 4
                college: row[4],     // College name
                age: row[5],         // Age
                city: row[6],        // City
                source: row[7],      // How they heard about the event
                passType: 'General'  // Default pass type if not specified
            };

            if (participant.email) {
                console.log(`Processing registration for ${participant.name} (${participant.email})`);
                await sendPass(participant);
                // Add delay to avoid hitting rate limits
                await new Promise(resolve => setTimeout(resolve, 1000));
            }
        }

    } catch (error) {
        console.error('Error processing registrations:', error);
    }
}

// Main execution
async function main() {
    console.log('Starting to process registrations...');
    
    // Verify email configuration first
    const emailConfigValid = await verifyEmailConfig();
    if (!emailConfigValid) {
        console.error('Email configuration is invalid. Please check your credentials.');
        return;
    }
    
    await processRegistrations();
    console.log('Finished processing registrations.');
}

// Run the script
main().catch(console.error);
