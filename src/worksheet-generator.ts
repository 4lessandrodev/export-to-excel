import { Border, Borders, Style, Workbook } from 'exceljs';
import { randomUUID } from 'crypto';
import fake from 'faker';

interface Register {
	id: string;
	name: string;
	description: string;
	amount: string;
	accountNumber: string;
	createdAt: Date;
}

export const GenerateWorkbook = async (): Promise<Buffer> => {

	const registers: Register[] = [];

	for (let index = 0; index < 100; index++) {
		const reg: Register = {
			id: randomUUID(),
			name: fake.name.findName(),
			description: fake.lorem.paragraph(),
			amount: fake.finance.amount(100, 9000, 2, 'R$ '),
			accountNumber: fake.finance.account(),
			createdAt: new Date(),
		};

		registers.push(reg);
	}
	
	const workbook = new Workbook();
	
	const worksheet = workbook.addWorksheet('example');

	const rows = registers.map((reg) => Object.values(reg));

	const defaultBorderStyle: Partial<Border> = {
		style: 'thin',
		color: {
			argb: '0000796C'
		}
	};
	
	const defaultBordersStyle: Partial<Borders> = {
		bottom: defaultBorderStyle,
		left: defaultBorderStyle,
		right: defaultBorderStyle,
		top: defaultBorderStyle,
	}

	const defaultRowStyle: Partial<Style> = { alignment: { horizontal: 'center' }, border: defaultBordersStyle };
	const customRowStyle: Partial<Style> = {
		alignment: { wrapText: true, horizontal: 'left' }, border: defaultBordersStyle
	};

	worksheet.columns = [
		{ key: 'id', width: 42, style: defaultRowStyle },
		{ key: 'name', width: 32, style: defaultRowStyle },
		{ key: 'description', width: 72, style: customRowStyle },
		{ key: 'accountNumber', width: 21, style: defaultRowStyle },
		{ key: 'amount', width: 21, style: defaultRowStyle },
		{ key: 'createdAt', width: 21, style: defaultRowStyle}
	];
	worksheet.addTable({
		name: 'tableName',
		ref: 'A1',
		headerRow: true,
		style: {
		  theme: 'TableStyleDark3',
		  showRowStripes: true,
		  showColumnStripes: true
		},
		columns: [
			{ name: 'id' },
			{ name: 'name' },
			{ name: 'description' },
			{ name: 'accountNumber' },
			{ name: 'amount' },
			{ name: 'createdAt' }
		],
		rows: rows
	});

	// set header style
	worksheet.getRow(1).font = { bold: true, size: 14, name: 'Tahoma' }

	const result = workbook.xlsx.writeBuffer({ filename: 'worksheet.xlsx' });

	return result as Promise<Buffer>;
	
}
