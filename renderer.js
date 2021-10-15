// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// No Node.js APIs are available in this process because
// `nodeIntegration` is turned off. Use `preload.js` to
// selectively enable features needed in the rendering
// process.

const { ipcRenderer } = require('electron');

// var myArgs = process.argv.slice(2); 

// if(myArgs.length<=0)
// {
//     //console.log("Must pass the gz file name");
//     return(1);
// }

// var inputFileName = myArgs[0];
// var outFileName = myArgs[0].replace(".gz","");
// var xlsFileName = outFileName.replace(".xml",".xlsx");
// //console.log("input file name: ", inputFileName);
// //console.log("output file name: ", outFileName);
// //console.log("excel file name: ", xlsFileName);
// const fileName = outFileName;



const selectFileBtn = document.getElementById('selectFile');
const fileNameStr = document.getElementById('sourcefilename');
// renderer

// selectFileBtn.onclick = e => {  
//     //console.log("Select BTN pressed");
//     e.preventDefault();
//     ipcRenderer.send('selectFile');
// };

ipcRenderer.on('fileSelected', (e, command) => {
    //console.log("File selected: ", command);
    fileNameStr.innerText = command;
    ipcRenderer.send('convertFile', command);
    });

ipcRenderer.on('fileConverted', (e, command) => {
    console.log("fileConverted: ", command);
    fileNameStr.innerText = command;
    });
        

    document.getElementById("selectFile").addEventListener("click", myFunction);


    function myFunction(e) {
        e.preventDefault();
        //console.log("test call")
//        document.getElementById("sourcefilename").innerHTML = "Hello World";
    var fileToCon = document.getElementById("fileToConvert").innerHTML;
    console.log("convert BTN pressed:", fileToCon);
    if( fileToCon !='')
    {
        console.log("call convert:", fileToCon);
        ipcRenderer.send('convertFile', fileToCon);
    }

      }

      document.getElementById("myFile").addEventListener("change", handleOnChange);


      function handleOnChange(){
        var x = document.getElementById("myFile");
        var txt = "";
        if ('files' in x) {
          if (x.files.length == 0) {
            txt = "Select one  files.";
          } else {
            console.log(x.files);
            for (var i = 0; i < x.files.length; i++) {
//              txt += "<br><strong>" + (i+1) + ". file</strong><br>";
              var file = x.files[i];
              console.log(file);
              if ('name' in file) {
//                txt += "name: " + file.name + "<br>";
//                txt += "file: " + file.path + "<br>";
                document.getElementById("fileToConvert").innerHTML = file.path;
            }
            //   if ('size' in file) {
            //     txt += "size: " + file.size + " bytes <br>";
            //   }
            }
          }
        } 
        else {
          if (x.value == "") {
            txt += "Select one file.";
          } else {
            txt += "The files property is not supported by your browser!";
            txt  += "<br>The path of the selected file: " + x.value; // If the browser does not support the files property, it will return the path of the selected file instead. 
          }
        }
        
        document.getElementById("sourcefilename").innerHTML = txt;
      }
           