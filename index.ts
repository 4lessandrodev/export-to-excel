import Express from 'express';
import bodyParser from 'body-parser'
import { GenerateWorkbook } from './src/worksheet-generator';

const app = Express();

app.use(bodyParser.json({ type: 'application/*+json' }));

app.get('/', async (req, res) => {
	const workbook = await GenerateWorkbook();
	res.setHeader('Content-Length', workbook.length);
	res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	res.setHeader('Content-Disposition', 'attachment; filename=worksheet.xlsx');
	res.status(200).write(workbook, 'binary');
	res.end();
});

app.listen(3000, 'localhost', () => console.log('server running'));
