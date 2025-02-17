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

async function exportCoursesToExcel() {
    console.log("function is running")
    const client = new Client(dbConfig);
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Courses');

    try {
        await client.connect();
        console.log("connected to database")

        const query = `
            SELECT 
                "Course"."title",
                "Course"."purchaseMode",
                "Course"."webPriceINR",
                "Course"."webPriceUSD",
                "Course"."slug",
                "Course"."language"
            FROM "Course" 
            WHERE "Course"."type" = 'NORMAL'
            AND "Course"."status" = 'PUBLISHED'
            ORDER BY "Course"."createdAt" DESC
        `;

        const result = await client.query(query);

        worksheet.columns = [
            { header: 'S.No', key: 'sno', width: 10 },
            { header: 'Title', key: 'title', width: 50 },
            { header: 'Type', key: 'purchaseMode', width: 15 },
            { header: 'INR Pricing', key: 'webPriceINR', width: 15 },
            { header: 'USD Pricing', key: 'webPriceUSD', width: 15 },
            { header: 'Language', key: 'language', width: 15 },
            { 
                header: 'Course Link', 
                key: 'courseLink', 
                width: 50,
                style: { font: { color: { argb: '0000FF' }, underline: true } }
            }
        ];

        result.rows.forEach((row, index) => {
            worksheet.addRow({
                sno: index + 1,
                title: row.title,
                purchaseMode: row.purchaseMode === 'PAID' ? 'Paid' : 'Free',
                webPriceINR: row.webPriceINR || 'N/A',
                webPriceUSD: row.webPriceUSD || 'N/A',
                language: row.language || 'N/A',
                courseLink: `https://euron.one/course/${row.slug}`
            });
        });

        worksheet.getRow(1).font = { bold: true };

        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const filename = `courses_report_${timestamp}.xlsx`;
        const filepath = path.join(__dirname, filename);

        await workbook.xlsx.writeFile(filepath);

        console.log(`Excel file created successfully at: ${filepath}`);
    } catch (error) {
        console.error('Error occurred:', error);
    } finally {
        await client.end();
    }
}

exportCoursesToExcel();