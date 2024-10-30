import { WorkBook } from 'xlsx';

export default function excell(array, XLSX: typeof import('xlsx'), workbook: WorkBook) {
    const newWorkSheet = XLSX.utils.aoa_to_sheet(array)




    const range = XLSX.utils.decode_range(newWorkSheet['!ref']!);
    for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            if (!newWorkSheet[cellAddress]) continue;

            newWorkSheet[cellAddress].s = {
                fill: {
                    fgColor: { rgb: 'FFFF00' }, // Yellow background
                },
                font: {
                    bold: true,
                    color: { rgb: '0000FF' }, // Blue text
                },
                alignment: {
                    horizontal: 'center',
                    vertical: 'center',
                },
            };
        }
    }

    const columnWidths = array[0].map((_, colIndex) =>
        ({ wch: Math.max(...array.map(row => (row[colIndex] ? row[colIndex].toString().length : 10))) })
    );
    newWorkSheet['!cols'] = columnWidths;

    if (workbook.SheetNames.includes('FilteredData')) {
        const sheetIndex = workbook.SheetNames.indexOf('FilteredData');
        workbook.SheetNames.splice(sheetIndex, 1);
        delete workbook.Sheets['FilteredData'];
    }








    XLSX.utils.book_append_sheet(workbook, newWorkSheet, 'Resuilt')
    const workbookBlob = new Blob([XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    const url = URL.createObjectURL(workbookBlob)
    const link = document.createElement('a')
    link.href = url
    link.download = 'excell.xlsx'
    document.body.appendChild(link)
    link.click();
    document.body.removeChild(link)
    console.log('end')
}
