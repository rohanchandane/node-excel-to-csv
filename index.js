const fs = require('fs');
const { stringify } = require('csv-stringify/sync');

const readXlsxFile = require('read-excel-file/node');
const { getSheetCount } = require('./readXlsx.js');

    async function main() {
        const filePath = './DummyData.xlsx';

        try {
            const sheetCount = await getSheetCount(filePath);

            for (let sheetNumber = 1; sheetNumber <= sheetCount; sheetNumber++ ) {

                readXlsxFile(filePath, { sheet: sheetNumber }).then((rows) => {
                
                    // getting all null values index from current sheet's titles
                    const nullIndexes = rows[0].reduce( (accumulator, currentItem, index)=> {
                        if (currentItem === null) {
                            accumulator.push(index);
                        }
                        return accumulator;
                    }, []);
                    const tableCount = nullIndexes.length + 1;
                    const headerLengthCount = rows[0].length;
                    const processedIndex = [-1, ...nullIndexes, headerLengthCount + 1];
                    
                    for (let i = 0; i < tableCount; i++) {
                        const startIndex = processedIndex[i] + 1;
                        const lastIndex = processedIndex[i + 1];
                        console.log("----- Table -----");
                        const table = rows.map(row => {
                            return row.slice(startIndex, lastIndex);                            
                        });
                        console.log(table);

                        const csvString = stringify([table]);    
                        fs.writeFile(sheetNumber.toString() + '_' + i + '.csv', csvString, (err) => {
                            if (err) throw err;
                            console.log('CSV file '+ sheetNumber.toString() + '_' + i + '.csv' +' created successfully!');
                        });
                    }
                  })

            }

        } catch (error) {
            console.log(error);
        }
    }

    main();

  
  