const { Client } = require('pg');
const ExcelJS = require('exceljs');
const path = require('path');
const dotenv = require('dotenv');

dotenv.config();

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
      SELECT 
        "User"."phone",
        "User"."firstName",
        "User"."lastName",
        "User"."email"
      FROM "User" 
      WHERE "User"."phone" IS NOT NULL
      ORDER BY "User"."createdAt" DESC
    `;

        const result = await client.query(query);

        worksheet.columns = [
            { header: 'First Name', key: 'firstName', width: 20 },
            { header: 'Last Name', key: 'lastName', width: 20 },
            { header: 'Email', key: 'email', width: 30 },
            { header: 'Phone Numbers', key: 'phone', width: 20 }
        ];

        result.rows.forEach(row => {
            if (/[a-zA-Z]/.test(row.phone)) {
                return;
            }

            const phoneRegex = /^\+?[\d\s]+$/;

            if (!phoneRegex.test(row.phone)) {
                return;
            }

            const cleanPhone = row.phone.replace(/\s/g, '');

            worksheet.addRow({
                firstName: row.firstName,
                lastName: row.lastName,
                email: row.email,
                phone: cleanPhone
            });
        });

        worksheet.getRow(1).font = { bold: true };

        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const filename = `phone_numbers_${timestamp}.xlsx`;
        const filepath = path.join(__dirname, filename);

        await workbook.xlsx.writeFile(filepath);

        console.log(`Excel file created successfully at: ${filepath}`);
    } catch (error) {
        console.error('Error occurred:', error);
    } finally {
        await client.end();
    }
}

exportPhonesToExcel();