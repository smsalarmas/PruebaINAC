const testFolder = '../Excel/';
const fs = require('fs');
var Excel = require('exceljs');

var wb = new Excel.Workbook();
var path = require('path');

function simpleStringify (object){
    var simpleObject = {};
    for (var prop in object ){
        if (!object.hasOwnProperty(prop)){
            continue;
        }
        if (typeof(object[prop]) == 'object'){
            continue;
        }
        if (typeof(object[prop]) == 'function'){
            continue;
        }
        simpleObject[prop] = object[prop];
    }
    return JSON.stringify(simpleObject); // returns cleaned up JSON
};

function ConvertirLetra(Letra) {
    switch (Letra) {
        case 'A':
            return '1' 
            break;
        case 'B':
            return '2' 
            break;
        case 'C':
            return '3' 
            break;
        case 'D':
            return '4' 
            break;
        case 'E':
            return '5' 
            break;
        default:
            return Letra
            break;
    }
}

// Paso 1 Lista todos los archivos que tengo en Excel.
fs.readdir(testFolder, (err, files) => {
  files.forEach(file => {
   // console.log(file);
    // Paso 2 leer la hora de excel.
    var filePath = path.resolve(__dirname,file);

    wb.xlsx.readFile(filePath).then(function(){

    wb.worksheets
    wb.worksheets.forEach(function (element, index) {
            var sh = element;

            console.log(sh.rowCount);
            //Get all the rows data [1st and 2nd column]
            const Materia = sh.getRow(1).getCell(1).value.richText[0].text;
            console.log("Materia: " + Materia + ' / '+ file);
            if (!fs.existsSync(file.replace(".xlsx", ""))){
                fs.mkdirSync(file.replace(".xlsx", ""));
            }
            for (i = 2; i <= sh.rowCount; i++) { // sh.rowCount
                /* console.log(sh.getRow(i).getCell(1).value);
                console.log(sh.getRow(i).getCell(2).value.text); */
                try {
                    if (sh.getRow(i).getCell(2).value.richText[0].text != "Pregunta" && ConvertirLetra(sh.getRow(i).getCell(3).value.richText[0].text) != undefined) {
                        var Pregunta;
                        try {
                            Pregunta = {
                                "Pregunta": sh.getRow(i).getCell(2).value.richText[0].text,
                                "RespuestaCorrecta": ConvertirLetra(sh.getRow(i).getCell(3).value.richText[0].text),
                                "A": sh.getRow(i).getCell(4).value.richText[0].text,
                                "B": sh.getRow(i).getCell(5).value.richText[0].text,
                                "C": sh.getRow(i).getCell(6).value.richText[0].text,
                                "D": sh.getRow(i).getCell(7).value.richText[0].text,
                            }
                        } catch (error) {
                            try {
                                Pregunta = {
                                    "Pregunta": sh.getRow(i).getCell(2).value.richText[0].text,
                                    "RespuestaCorrecta": ConvertirLetra(sh.getRow(i).getCell(3).value.richText[0].text),
                                    "A": sh.getRow(i).getCell(4).value.richText[0].text,
                                    "B": sh.getRow(i).getCell(5).value.richText[0].text,
                                    "C": sh.getRow(i).getCell(6).value.richText[0].text,                                     
                                }
                            } catch (error) {
                                Pregunta = {
                                    "Pregunta": sh.getRow(i).getCell(2).value.richText[0].text,
                                    "RespuestaCorrecta": ConvertirLetra(sh.getRow(i).getCell(3).value.richText[0].text),
                                    "A": sh.getRow(i).getCell(4).value.richText[0].text,
                                    "B": sh.getRow(i).getCell(5).value.richText[0].text,                                                                         
                                }
                            }
                        }
                        
                        
                        
                        console.log(Pregunta)
                        const SalidaFile = file.replace(".xlsx", "") + '/_' + (index + 1) + '.json'
                        console.log(SalidaFile);
                        fs.appendFile(SalidaFile, JSON.stringify(Pregunta), function (err) {
                            if (err) throw err;
                            console.log('Saved!');
                        });
            
                    }
                } catch (error) {
                    // console.log(sh);
                    console.error(error);
                    console.log(file);
                    const txtWriteError = simpleStringify(sh);
                    fs.appendFile('Error.json', txtWriteError, function (err) {
                        if (err) throw err;
                        console.log('Saved!');
                    });
                    
                }
                
                
                // Debo Escribir esto en un archivo .json

            }
    });
    
});

  });
});

