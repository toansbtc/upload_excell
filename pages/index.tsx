import { Button } from 'bootstrap';
import React, { useEffect, useState } from 'react'
// import excell from './function/excell';
// import * as XLSX from 'xlsx'

let XLSX: typeof import('xlsx') | undefined = undefined;

if (typeof window !== 'undefined') {
  XLSX = require('xlsx');
}

export default function index() {
  const [file, setfile] = useState<File | null>(null);


  const getFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    e.preventDefault();
    if (e.target.files && e.target.files[0]) {
      setfile(e.target.files[0]);

    }

  }
  const transform = () => {
    if (file) {
      const reader = new FileReader();
      if (XLSX)
        reader.onload = async (e) => {
          const data = new Uint8Array(e.target.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'buffer' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const secondSheet = workbook.Sheets[workbook.SheetNames[1]];
          const thirdSheet = workbook.Sheets[workbook.SheetNames[2]];

          const jsonData_sheet1 = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
          const jsonData_sheet2 = XLSX.utils.sheet_to_json(secondSheet, { header: 1 });
          const jsonData_sheet3 = XLSX.utils.sheet_to_json(thirdSheet, { header: 1 });


          // const rows = jsonData_sheet2.slice(1) as any[][]
          // let data_Resuilt_Array: unknown[][] = new Array(rows.length).fill([]).map(() => new Array());



          const data_Array_Sheet1: any[] = []
          const data_Array_merge = new Map<string, [string, string]>()
          const data_Array_Sheet2: any[][] = [];
          const data_Array_Sheet3: any[][] = [];
          const data_Array_Sheet3_alone: any[] = [];
          const data_Array_Resuilt: any[][] = []
          const uniqueValues = new Set<string>();

          jsonData_sheet1.slice(8).forEach((data) => {
            if (!uniqueValues.has(data[10]) && data[10] != undefined) {
              uniqueValues.add(data[10])
              data_Array_Sheet1.push(data[10])
            }
          })
          uniqueValues.clear();

          // console.log(data_Array_Sheet1)


          jsonData_sheet2.slice(1).forEach((data) => {
            const key = `${data[1]}_${data[12]}`;
            if (!uniqueValues.has(key)) {
              uniqueValues.add(key);
              data_Array_Sheet1.forEach((data1) => {
                if (data[12] == data1)
                  data_Array_Sheet2.push([data[1], data[12]]);
              })
            }
          });
          uniqueValues.clear();
          // console.log(data_Array_Sheet2)



          // jsonData_sheet3.slice(1).forEach(data => {
          //   data_Array_Sheet2.forEach(data1 => {
          //     const key = `${data[0]}_${data[4]}_${data1[1]}`;
          //     if (data[2] == data1[0] && !uniqueValues.has(key)) {
          //       data_Array_Sheet3.push([data[4], data[0], data1[1]])
          //       uniqueValues.add(key)
          //     }
          //   })
          // })
          // uniqueValues.clear()

          data_Array_Sheet1.forEach(data => {
            data_Array_Sheet2.forEach(data1 => {
              if (data == data1[1])
                uniqueValues.add(data)
            })
          })
          data_Array_Sheet1.forEach((data) => {
            if (!uniqueValues.has(data))
              data_Array_Sheet3_alone.push(data)
          })

          uniqueValues.clear()
          // console.log(data_Array_Sheet3_alone)

          // jsonData_sheet3.slice(0, 10).forEach(data => {
          //   data_Array_Resuilt.push([data[0], data[1], data[2], data[2], data[4], data[5], data[6], data[7], data[8], data[9], data[10], data[11], data[12], data[13], data[14], data[15]])

          // })
          // console.log(data_Array_Resuilt)
          data_Array_Resuilt.push(['物料', '訂單數量 (GMEIN)', '訂單', '物料說明', '合併主檔訂單', '採購單號碼'])
          jsonData_sheet3.slice(1).forEach((data) => {
            data_Array_Sheet2.forEach(data1 => {

              if (!uniqueValues.has(data[12] + '_' + data[22])) {
                data_Array_Resuilt.push([data[2], data[6], data[12], data[15], data[22], ''])
                uniqueValues.add(data[12] + '_' + data[22])

              }

            })
          })

          jsonData_sheet3.slice(1).forEach((data) => {
            data_Array_Sheet2.forEach(data1 => {
              if (data[12] == data1[0])
                data_Array_Resuilt.push([data[2], data[6], data1[0], data[15], data[22], data1[1]])

            })
          })

          data_Array_Sheet1.forEach(data => {

            data_Array_Resuilt.slice(1).forEach(data1 => {
              if (data.toString().trim() == data1[5].toString().trim())
                if (data_Array_merge.has(data.toString().trim())) {
                  data_Array_merge.set(data.toString().trim(), [`${data_Array_merge.get(data.toString().trim())[0]}\n${data1[4]}`, `${data_Array_merge.get(data.toString().trim())[1]}\n${data1[0]}`])
                }
                else {
                  data_Array_merge.set(data.toString().trim(), [`${data1[4]}`, `${data1[0]}`])

                }

            })
          })


          const custom_data_array_sheet1: any[][] = []
          let i = 1
          jsonData_sheet1.forEach(data => {
            if (data[10] != undefined && typeof data[10] == 'number') {
              if (data_Array_merge.has(data[10].toString().trim())) {
                let data1 = ''
                let data2 = ''
                uniqueValues.clear()
                data_Array_merge.get(data[10].toString().trim())[0].split('\n').forEach(data => {

                  if (!uniqueValues.has(data)) {
                    uniqueValues.add(data)
                    data1 += data + '\n'
                  }

                })
                uniqueValues.clear()
                data_Array_merge.get(data[10].toString().trim())[1].split('\n').forEach(data => {
                  if (!uniqueValues.has(data)) {
                    uniqueValues.add(data)
                    data2 += data + '\n'
                  }

                })

                XLSX.utils.sheet_add_aoa(firstSheet, [[data1.trim()]], { origin: `B${i}` })
                XLSX.utils.sheet_add_aoa(firstSheet, [[data2.trim()]], { origin: `D${i}` })



              }

            }
            i++

          })

          for (let key in firstSheet) {
            if (firstSheet.hasOwnProperty(key) && key[0] !== '!') {
              if (typeof firstSheet[key].v === 'string')
                firstSheet[key].v = firstSheet[key].v.replace(/\n/g, '\n');
            }
          }

          // }
          // else
          //   custom_data_array_sheet1.push([data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8], data[9], data[10], data[11], data[12], data[13], data[14], data[15]])


          // console.log(custom_data_array_sheet1)

          //insert data in excell
          // await excell(data_Array_Resuilt, XLSX, workbook)



          // const newWorkSheet = XLSX.utils.aoa_to_sheet(data_Array_Resuilt)
          // XLSX.utils.book_append_sheet(workbook, newWorkSheet, 'Resuilt')
          // const workbookBlob = new Blob([XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

          // const url = URL.createObjectURL(workbookBlob)
          // const link = document.createElement('a')
          // link.href = url
          // link.download = 'excell.xlsx'
          // document.body.appendChild(link)
          // link.click();
          // document.body.removeChild(link)


        };
      reader.readAsArrayBuffer(file);
    }
  }
  return (
    <div className='container-fluid'>
      <form className='form-control'>
        <input type='file' onChange={getFile} accept=".xls,.xlsx" />
      </form>
      <button onClick={transform}>chuyen doi</button>
    </div>
  )
}
