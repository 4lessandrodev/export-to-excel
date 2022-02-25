import Express from 'express';
import bodyParser from 'body-parser'
import { GenerateWorkbook } from './src/worksheet-generator';

const app = Express();

app.use(bodyParser.json({ type: 'application/*+json' }));

app.get('/', async (req, res) => {
	const { workbook, size } = await GenerateWorkbook();
	res.setHeader('Content-Length', size);
	res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	res.setHeader('Content-Disposition', 'attachment; filename=worksheet.xlsx');
	res.status(200);
	
	workbook.pipe(res);

});

app.listen(3000, 'localhost', () => console.log('server running'));
