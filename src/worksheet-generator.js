"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.GenerateWorkbook = void 0;
const exceljs_1 = require("exceljs");
const crypto_1 = require("crypto");
const faker_1 = __importDefault(require("faker"));
const GenerateWorkbook = () => __awaiter(void 0, void 0, void 0, function* () {
    const registers = [];
    for (let index = 0; index < 100; index++) {
        const reg = {
            id: (0, crypto_1.randomUUID)(),
            name: faker_1.default.name.findName(),
            description: faker_1.default.lorem.paragraph(),
            amount: faker_1.default.finance.amount(100, 9000, 2, 'R$ '),
            accountNumber: faker_1.default.finance.account(),
            createdAt: new Date(),
        };
        registers.push(reg);
    }
    const workbook = new exceljs_1.Workbook();
    const worksheet = workbook.addWorksheet('example');
    const rows = registers.map((reg) => Object.values(reg));
    const defaultBorderStyle = {
        style: 'thin',
        color: {
            argb: '0000796C'
        }
    };
    const defaultBordersStyle = {
        bottom: defaultBorderStyle,
        left: defaultBorderStyle,
        right: defaultBorderStyle,
        top: defaultBorderStyle,
    };
    const defaultRowStyle = { alignment: { horizontal: 'center' }, border: defaultBordersStyle };
    const customRowStyle = {
        alignment: { wrapText: true, horizontal: 'left' }, border: defaultBordersStyle
    };
    worksheet.columns = [
        { key: 'id', width: 42, style: defaultRowStyle },
        { key: 'name', width: 32, style: defaultRowStyle },
        { key: 'description', width: 72, style: customRowStyle },
        { key: 'accountNumber', width: 21, style: defaultRowStyle },
        { key: 'amount', width: 21, style: defaultRowStyle },
        { key: 'createdAt', width: 21, style: defaultRowStyle }
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
    worksheet.getRow(1).font = { bold: true, size: 14, name: 'Tahoma' };
    const result = workbook.xlsx.writeBuffer({ filename: 'worksheet.xlsx' });
    return result;
});
exports.GenerateWorkbook = GenerateWorkbook;
