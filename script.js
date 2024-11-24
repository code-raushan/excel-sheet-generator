const { Client } = require('pg');
const ExcelJS = require('exceljs');
const path = require('path');
const dotenv = require('dotenv');

dotenv.config();

// Database configuration
const dbConfig = {
    user: process.env.DB_USER,
    host: process.env.DB_HOST,
    database: process.env.DB_DATABASE,
    password: process.env.DB_PASSWORD,
    port: process.env.DB_PORT,
};

async function exportPhonesToExcel() {
    const client = new Client(dbConfig);
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Phone Numbers');

    try {
        // Connect to database
        await client.connect();

        // Execute query
        const query = `
      SELECT "User"."phone" 
      FROM "User" 
      WHERE "User"."phone" IS NOT NULL
      AND "User"."authProvider" = 'PHONE'
      ORDER BY "User"."createdAt" DESC
    `;

        const result = await client.query(query);

        // Set up worksheet header
        worksheet.columns = [
            { header: 'Phone Numbers', key: 'phone', width: 20 }
        ];
        let invalidPhoneCount = 0;

        // Add data to worksheet
        result.rows.forEach(row => {
            // Check if phone number contains any alphabetic characters
            if (/[a-zA-Z]/.test(row.phone)) {
                console.log(`Skipping phone number with letters: ${row.phone}`);
                invalidPhoneCount++;
                return;
            }

            // Basic phone number validation: must be only numbers and spaces, optionally starting with +
            const phoneRegex = /^\+?[\d\s]+$/;

            // Skip if phone number doesn't match the regex pattern
            if (!phoneRegex.test(row.phone)) {
                console.log(`Skipping invalid phone number: ${row.phone}`);
                invalidPhoneCount++;
                return;
            }

            // Remove spaces from the phone number before adding to worksheet
            const cleanPhone = row.phone.replace(/\s/g, '');
            worksheet.addRow({ phone: cleanPhone });
        });

        // Style the header row
        worksheet.getRow(1).font = { bold: true };

        // Generate filename with timestamp
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const filename = `phone_numbers_${timestamp}.xlsx`;
        const filepath = path.join(__dirname, filename);

        // Save the workbook
        await workbook.xlsx.writeFile(filepath);

        console.log(`Excel file created successfully at: ${filepath}`);
        console.log(`Total phone numbers exported: ${result.rows.length}`);
        console.log(`Total invalid phone numbers: ${invalidPhoneCount}`);
    } catch (error) {
        console.error('Error occurred:', error);
    } finally {
        // Close database connection
        await client.end();
    }
}

// Run the export function
exportPhonesToExcel();