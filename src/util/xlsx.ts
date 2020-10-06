interface Window { 
   loadExcelData(): any; 
} 
window.loadExcelData = loadExcelData; 

function loadExcelData() { 
   window.console.log('excel data loaded'); 

   Excel.run(function (ctx) { 

      let myNumber: any; 
      let wbk: Excel.Workbook = ctx.workbook; 
      let wsh: Excel.Worksheet = wbk.worksheets.getActiveWorksheet(); 
      wsh.load("items/name"); 

      myNumber = [[Math.floor(Math.random() * 1000)]]; 
      wsh.getRange("B2").values = myNumber; 

      AddToListBox1(myNumber.toString()); 

      return ctx.sync() 
         .then(function () { 

            AddToListBox2(wsh); 
         }) 
   }); 
} 

function AddToListBox1( 
   someText: string) { 

   var element2 = document.getElementById("textareaID"); 
   element2.innerText = someText + repeatStringNumTimes(' ', 35 - someText.length) + element2.innerText; 
} 

function AddToListBox2( 
   _wsh: Excel.Worksheet) { 

   var thetext = _wsh.name; 
   var element2 = document.getElementById("textareaID"); 
   element2.innerText = thetext + repeatStringNumTimes(' ', 35 - thetext.length) + element2.innerText; 
} 

function repeatStringNumTimes( 
   _string: string, 
   _times: number): string { 

   let repeatedString: string = ""; 
   while (_times > 0) { 
      repeatedString += _string; 
      _times--; 
   } 
   return repeatedString; 
} 
