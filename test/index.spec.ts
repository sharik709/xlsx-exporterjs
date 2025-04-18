import { ExcelWorkbook, Sheet } from '../src';

describe('index', () => {
    describe('ExcelWorkbook', () => {
        it('We can create new workbook object', () => {
            const result = new ExcelWorkbook();

            expect(result).toBeInstanceOf(ExcelWorkbook);
        });

        it('We can set a name for the workbook', () => {
            const name = 'TestBook';
            const wb = new ExcelWorkbook();
            wb.setName(name);
            expect(wb.getName()).toBe(name);
        });

        it('We can set Sheets', () => {
            const s = new Sheet();
            const wb = new ExcelWorkbook();
            wb.setSheets([s]);
            expect(wb.getSheets()).toEqual([s]);
        });

        it('throws if sheets is null or undefined', () => {
            const wb = new ExcelWorkbook();
            expect(() => wb.setSheets(undefined as any)).toThrowError(
                new Error('Sheets cannot be null or undefined')
            );
            expect(() => wb.setSheets(null as any)).toThrowError(
                new Error('Sheets cannot be null or undefined')
            );
        });

        it('throws if sheets is not an array', () => {
            const wb = new ExcelWorkbook();
            expect(() => wb.setSheets('notanarray' as any)).toThrowError(
                new Error('Sheets must be an array')
            );
        });

        it('throws if sheets is empty', () => {
            const wb = new ExcelWorkbook();
            expect(() => wb.setSheets([])).toThrowError(
                new Error('Sheets cannot be empty')
            );
        });
    });

    describe('Sheet', () => {
        it('tests if we can create sheet object', () => {
            const s = new Sheet();
            expect(s).toBeInstanceOf(Sheet);
        });

        it('tests if rows can be added to sheet', () => {
            const data = ['a', 'b', 'c'];
            const s = new Sheet();
            s.setRows(data);
            expect(s.getRows()).toEqual(data);
        });

        it('Set and get sheet name', () => {
            const name = 'testSheet';

            const s = new Sheet();
            s.setName(name);
            expect(s.getName()).toEqual(name);
        });

        it('appends and gets a row from sheet', () => {
            const row = ['a', 'b', 'c'];
            const s = new Sheet();
            s.appendRow(row);
            expect(s.getRow(0)).toEqual(row);
        });

        it('throws if index is not a number', () => {
            const s = new Sheet();
            expect(() => s.getRow('foo' as any)).toThrowError(
                new TypeError('Index must be a number')
            );
        });

        it('throws if index is out of bound (negative or too large)', () => {
            const s = new Sheet();
            // empty rows ⇒ any index is out of bound
            expect(() => s.getRow(-1)).toThrowError(
                new TypeError('Index out of bound')
            );
            expect(() => s.getRow(0)).toThrowError(
                new TypeError('Index out of bound')
            );
        });

        it('throws if row at index exists but is undefined/null', () => {
            const s = new Sheet();
            // force an undefined entry in the rows array
            s.setRows([undefined]);
            expect(() => s.getRow(0)).toThrowError(
                new TypeError('Index does not exists')
            );
        });
    });
});
