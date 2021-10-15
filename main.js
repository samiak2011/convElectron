// Modules to control application life and create native browser window
const { app, BrowserWindow, Menu, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const convert = require('xml-js');
// https://www.brcline.com/blog/how-to-write-an-excel-file-in-nodejs
const Excel = require('exceljs');
const zlib = require('zlib');    
const unzip = zlib.createUnzip();  
const xml2js = require('xml2js');
const parser = new xml2js.Parser({ attrkey: "ATTR" });
const nativeImage = require('electron').nativeImage;

const template = [
  {label: "Export to Excel"}
 // {    label: 'File',    submenu: [       { role: 'quit' }    ]  }
];

function createWindow () {
  const iconPath = path.join(__dirname, "logo.png");
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    resizable: false,
    title: "Convert .gz file to Excel",
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
    icon: iconPath
  });

  const image = nativeImage.createFromPath(iconPath);
  mainWindow.setIcon(image);
  //console.log(image);
    // and load the index.html of the app.
  mainWindow.loadFile(path.join(__dirname, 'index.html'));

  // Open the DevTools.
  mainWindow.webContents.openDevTools();
//  console.log(dialog.showOpenDialog({ properties: ['openFile', 'multiSelections'] }))
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.whenReady().then(() => {
  createWindow();

  app.on('activate', function () {
    // On macOS it's common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    if (BrowserWindow.getAllWindows().length === 0) 
      createWindow();

  });
});

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit()
})

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and require them here.
const menu = Menu.buildFromTemplate(template);
Menu.setApplicationMenu(menu);


ipcMain.on('selectFile', (event, arg) => {  
  //console.log("select file called", arg); 
  // prints "ping"  
  event.reply('fileSelected', '/home/sam/Proj/nodejs/xmlparse/files/2021-10-03.029.2.xml.gz');
});

function testCall(arg)
{
  //console.log("test call: ", arg); 
}

ipcMain.on('convertFile', (event, arg) => {  
  //console.log("File to convert: ", arg); 
  
  retMsg = convertFileFunction(arg);

  //testCall(arg);

  event.reply('fileConverted', retMsg);
});


let totalCount = 0;

class TransData {
    constructor(trDate, trNO, trAmount, trCustomer, trCustCode, trProduct)
    {
        this.trDate = trDate;
        this.trNO = trNO;
        this.trAmount = trAmount;
        this.trCustomer = trCustomer;
        this.trCustCode = trCustCode;
        this.trProduct = trProduct;
    }

};


function parseNodetrLines(xmlNode)
{
    /**

trLine: [
    {
      ATTR: [Object],
      trlTaxes: [Array],
      trlFee: [Array],
      trlFlags: [Array],
      trlDept: [Array],
      trlNetwCode: [Array],
      trlQty: [Array],
      trlSign: [Array],
      trlSellUnit: [Array],
      trlUnitPrice: [Array],
      trlLineTot: [Array],
      trlDesc: [Array],
      trlUPC: [Array],
      trlModifier: [Array],
      trlUPCEntry: [Array]
    },
    {
      ATTR: [Object],
      trlTaxes: [Array],
      trlFee: [Array],
      trlFlags: [Array],
      trlDept: [Array],
      trlNetwCode: [Array],
      trlQty: [Array],
      trlSign: [Array],
      trlSellUnit: [Array],
      trlUnitPrice: [Array],
      trlLineTot: [Array],
      trlDesc: [Array],
      trlUPC: [Array],
      trlModifier: [Array],
      trlUPCEntry: [Array],
      trlMixMatches: [Array]
    },
    {
      ATTR: [Object],
      trlTaxes: [Array],
      trlFee: [Array],
      trlFlags: [Array],
      trlDept: [Array],
      trlNetwCode: [Array],
      trlQty: [Array],
      trlSign: [Array],
      trlSellUnit: [Array],
      trlUnitPrice: [Array],
      trlLineTot: [Array],
      trlDesc: [Array],
      trlUPC: [Array],
      trlModifier: [Array],
      trlUPCEntry: [Array],
      trlMixMatches: [Array]
    },
    {
      ATTR: [Object],
      trlTaxes: [Array],
      trlFlags: [Array],
      trlDept: [Array],
      trlNetwCode: [Array],
      trlQty: [Array],
      trlSign: [Array],
      trlSellUnit: [Array],
      trlUnitPrice: [Array],
      trlLineTot: [Array],
      trlDesc: [Array],
      trlUPC: [Array],
      trlModifier: [Array],
      trlUPCEntry: [Array]
    },
    {
      ATTR: [Object],
      trlTaxes: [Array],
      trlFlags: [Array],
      trlDept: [Array],
      trlNetwCode: [Array],
      trlQty: [Array],
      trlSign: [Array],
      trlUnitPrice: [Array],
      trlLineTot: [Array],
      trlFuel: [Array],
      trlDesc: [Array]
    }
  ]
    */
    ////console.log("Start lines");
    var prodName = "";
    var trPaylines = xmlNode['trLines'];
    for(i=0; i< trPaylines.length;i++)
    {
        for(j=0; j<trPaylines[i]['trLine'].length; j++)
        {
            if(trPaylines[i]['trLine'][j]['trlFuel']!= undefined)
            {
                for(k=0;k<trPaylines[i]['trLine'][j]['trlFuel'].length;k++)
                {
                    prodName += trPaylines[i]['trLine'][j]['trlFuel'][k]['fuelProd'][0]['_'];
                    ////console.log(i,j,k,prodName);
                }
            }
            }
        //.trlFuel);
    } 
    ////console.log("End lines");
    return prodName;
}

function parseNodetrPaylines(xmlNode)
{
    /**
     * {
      ATTR: [Object],
      trpPaycode: [Array],
      trpAmt: [Array],
      trpHouseAcct: [Array]
    }
     */
    var trPaylines = xmlNode['trPaylines'];
    let cusName = "";
    let cusCode = "";
    if(trPaylines != undefined)
    {
        var tttr = trPaylines[0].trPayline;
        for(i=0; i< tttr.length;i++)
        {
            if(tttr[i].trpHouseAcct != undefined)
            {
                for(j=0; j<tttr[i].trpHouseAcct.length; j++)
                    if(tttr[i].trpHouseAcct[j]['ATTR']!= undefined){
                        cusName = tttr[i].trpHouseAcct[j]['ATTR'].name;
                        cusCode = tttr[i].trpHouseAcct[j]['_'];
                        ////console.log(totalCount++, i, j, cusCode);//['ATTR'])
                    }
            }
        }
    }
    else{
        cusName = "";
    } 
    return { cusName, cusCode };
}


function parseNode(xmlNode)
{
    customer = parseNodetrPaylines(xmlNode);
    custName = customer.cusName;
    custCode = customer.cusCode;
    if(custName.length==0)
        return 0;
    prodName = parseNodetrLines(xmlNode);

    var data1 = new TransData(
        //date
        xmlNode['trHeader'][0].date[0], 
        //no
        xmlNode['trHeader'][0].termMsgSN[0]['_'],
        // Amount  
        xmlNode['trValue'][0].trTotWTax[0],
        // customer
        custName,
        // trCustCode
        custCode,
        // trProduct
        prodName
    );


    ////console.log("Start");
     //parseNodetrPaylines(xmlNode);

    //not needed
    //    //console.log(xmlNode)/;
    //parseNodetrLines(xmlNode);
    ////console.log("End");
    return data1;
}



function convertFileFunction(fileNamePassed)
{
    
    var inputFileName = fileNamePassed;
    var outFileName = fileNamePassed.replace(".gz","");
    var xlsFileName = outFileName.replace(".xml",".xlsx");
    //console.log("input file name: ", inputFileName);
    //console.log("output file name: ", outFileName);
    //console.log("excel file name: ", xlsFileName);
    const fileName = outFileName;

  const inp = fs.createReadStream(inputFileName);  
  const out = fs.createWriteStream(outFileName);  
      
  inp.pipe(unzip).pipe(out);  

  out.on("close", () => {
    const xmlFile = fs.readFileSync(fileName, 'utf8');

    parser.parseString(xmlFile, 
        function(error, result) {
        if(error === null) {
            AllTransData = [];
            let j=0;
            arrayObj = result['transSet']['trans']; 
            for (let i = 0; i < arrayObj.length; i++) {
                if(arrayObj[i].ATTR.type == 'sale')
                {
                    var yyyy = parseNode(arrayObj[i]);
                    if(yyyy != 0)
                    {
                        AllTransData[j++] = yyyy;
                        ////console.log("yy: ", yyyy);
                    }
                }
            }

            // //console.log("j= ", j);
            // //console.log("array = ", AllTransData.length);
            xyz = 0;
            let workbook = new Excel.Workbook();
            let worksheet = workbook.addWorksheet('Transactions');
            // trDate, trNO, trAmount, trCustomer, trCustCode, trProduct
            worksheet.columns = [
                {header: 'Date', key: 'trDate'},
                {header: 'NO', key: 'trNO'},
                {header: 'Amount', key: 'trAmount'},
                {header: 'Customer', key: 'trCustomer'},
                {header: 'Customer Code', key: 'trCustCode'},
                {header: 'Product', key: 'trProduct'}
              ];
            // force the columns to be at least as long as their header row.
            // Have to take this approach because ExcelJS doesn't have an autofit property.
            worksheet.columns.forEach(column => {
                column.width = column.header.length < 12 ? 12 : column.header.length;
            });
            
            // Make the header bold.
            // Note: in Excel the rows are 1 based, meaning the first row is 1 instead of 0.
            worksheet.getRow(1).font = {bold: true};

            AllTransData.forEach((element, index) => {
                const rowIndex = index + 2;
                worksheet.addRow({...element});
                //console.log(xyz, element);
                xyz++;
            });
            workbook.xlsx.writeFile(xlsFileName);
            return "Convert completed file: "+xlsFileName;
            //console.log("Finish: ", xlsFileName);
        }
        else {
          return "Failed to convert";
            //console.log(error);
        }
    });

});

return "Convert completed file: "+xlsFileName;

}
