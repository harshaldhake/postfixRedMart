import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import * as _ from 'lodash';
import * as postfixCalculator from 'postfix-calculator';

type AOA = any[][];

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'app!';
 
  data: AOA = [];
  arrayData = [];
  globalArray = [];
  finalArray = [];
  mainCopyOfArray = [];

	wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
	fileName: string = 'SheetJS.xlsx';

    onFileChange(evt: any) {
      /* wire up file reader */
      const target: DataTransfer = <DataTransfer>(evt.target);
      if (target.files.length !== 1) throw new Error('Cannot use multiple files');
      const reader: FileReader = new FileReader();
      reader.onload = (e: any) => {
        /* read workbook */
        const bstr: string = e.target.result;
        const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});

        /* grab first sheet */
        const wsname: string = wb.SheetNames[0];
        const ws: XLSX.WorkSheet = wb.Sheets[wsname];

        /* save data */
        this.data = <AOA>(XLSX.utils.sheet_to_json(ws, {header: 1}));
    
        //Generate all the cells where only numbers present;
        for (var i = 0; i < this.data.length; i++)
        {
  
          if(this.data[i].length > 0)
          {
            this.arrayData.push(this.data[i]);
          }
        }

        this.arrayData = this.generateArrayData(this.arrayData);

        //Replace the cells values generated using function postFixValueGeneration
        this.arrayData = this.replaceValuesFromGlobal(this.globalArray,this.arrayData);

        this.arrayData = this.reIterate(this.globalArray,this.arrayData);
        console.log(this.arrayData);
        // This will print the final array in console
        this.finalArray = this.finalOutPut(this.arrayData);
        console.log(this.finalArray);
      };
      reader.readAsBinaryString(target.files[0]);

  }

   generateArrayData(mArrayData){
    this.globalArray = [];
    for (let j = 2; j < mArrayData.length; j++)
    {
      for (let m = 1; m < mArrayData[j].length; m++)
      {
        // this method will generate a array e. A1-> 3, A2-> 20 , A-> 40, B1-> 30, B2->50
        this.pushToGlobalArray(mArrayData[j][m], mArrayData[j][0], m);
      }
    }
    return mArrayData;
   }

   replaceValuesFromGlobal(vGlobalArray, vArrayData){
      for (let j = 2; j < vArrayData.length; j++)
      {
        for (let m = 1; m < vArrayData[j].length; m++)
        {
          let findNullElem =  _.join([vArrayData[j][0], m], '');
        
            for(let k = 0; k< vGlobalArray.length; k++ )
            {
              var arrKeys = Object.keys(vGlobalArray[k]);
              
              var arrValues = Object.values(vGlobalArray[k]);
              if( arrValues[0] !== null)
              {
                vArrayData[j][m] =  vArrayData[j][m].replace(arrKeys[0], arrValues[0]);  
              }
            }

          }
        }
      return vArrayData;
   }


   pushToGlobalArray(position, j, m){
      var v =  postfixCalculator(position);
        //Generate Json array of values which are generated.
          let jsonData = {};
          let rowChar =  _.join([j, m], '');
          jsonData[rowChar] = v;
          this.globalArray.push(jsonData);
   }


   reIterate(pGlobalArray, pArrayData){
    for(let t = 0; t< pGlobalArray.length; t++ )
      {
        var arrValues = Object.values(pGlobalArray[t]);
        if( arrValues[0] == null)
        {
          this.arrayData = this.generateArrayData(pArrayData);

          // This will get values from Global Array  and replace in the Json array which was generated from xlxs file.
          pArrayData = this.replaceValuesFromGlobal(this.globalArray,this.arrayData);
        }
      }
    return pArrayData;
   }

   int2Float(n, place) { return n.toFixed(place); }

   
   finalOutPut(arrayData){
    let jsonData;
    let height = arrayData.length -2;    
    for (let j = 2; j < arrayData.length; j++)
    {
      if(j == 2){
        let width = arrayData[j].length - 1 ;
        jsonData = width + ' '+ height  +  '               :' +  width + ' '+ height ;
        this.finalArray.push(jsonData);
      }
      
      for (let m = 1; m < arrayData[j].length; m++)
      {
        // this method call will generate a array with cell and respective value.
        this.pushToGlobalArray(arrayData[j][m], arrayData[j][0], m);
        var val =  postfixCalculator(arrayData[j][m]);
        //Generate Json array of values which are generated.
            jsonData =  arrayData[j][m]  +  '               :' +  this.int2Float(val,6);
            this.finalArray.push(jsonData);
      }
    }
    return this.finalArray;
   }

}