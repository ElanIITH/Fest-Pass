const { google } = require('googleapis');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const handlebars = require('handlebars');
const puppeteer = require('puppeteer');
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

async function generatePDF(htmlContent) {
    const browser = await puppeteer.launch({ 
        headless: 'new',
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();
    
    // Adjusted viewport size for single page
    await page.setViewport({
        width: 800,
        height: 1000,  // Reduced height
        deviceScaleFactor: 1,
    });
    
    await page.setContent(htmlContent);
    
    const pdf = await page.pdf({
        format: 'A4',
        printBackground: true,
        margin: {
            top: '15px',    // Reduced margins
            right: '15px',
            bottom: '15px',
            left: '15px'
        },
        preferCSSPageSize: true,
        scale: 0.9         // Slightly scale down content
    });
    
    await browser.close();
    return pdf;
}

async function sendPass(participant) {
    try {
        const barcodeDataUrl = await generateBarcode(participant.email);
        
        // Read the header image with error handling
        let headerBase64;
        try {
            // Update the path to be relative to the project root
            const headerImagePath = path.join(__dirname, '..', 'header.png');  // Changed from 'assets/header.png'
            const headerImage = fs.readFileSync(headerImagePath);
            headerBase64 = `data:image/png;base64,${headerImage.toString('base64')}`;
            console.log('Header image loaded successfully');
        } catch (error) {
            console.warn('Header image not found at path:', path.join(__dirname, '..', 'header.png'));
            console.warn('Error:', error.message);
            headerBase64 = null;
        }

        // Prepare HTML content with header image or fallback
        const htmlContent = template({
            Name: participant.name,
            Pass: participant.passType,
            ALT: `ELAN_24_${participant.email}`,
            barcode: barcodeDataUrl,
            College: participant.college,
            City: participant.city,
            headerImage: headerBase64,
            useHeaderFallback: !headerBase64
        });

        // Generate PDF
        const pdf = await generatePDF(htmlContent);

        // Send email with more personalized content
        const mailOptions = {
            from: process.env.EMAIL_USER,
            to: participant.email,
            subject: 'Your Elan & nVision 2025 Pass',
            html: `
                <p>Dear ${participant.name},</p>
                <p>Thank you for registering for Elan & nVision 2025! Please find your event pass attached to this email.</p>
                <p>Your Registration Details:</p>
                <ul>
                    <li>Name: ${participant.name}</li>
                    <li>College: ${participant.college}</li>
                    <li>City: ${participant.city}</li>
                </ul>
                <p>Please keep this pass handy during the event.</p>
                <p>Best regards,<br>Elan & nVision, IIT Hyderabad</p>
            `,
            attachments: [
                {
                    filename: 'Elan-nVision-Pass.pdf',
                    content: pdf
                }
            ]
        };

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
