+(function (window, document) {

  'use strict';

  function sheetWriteData(e) {
    let sendData;
    let arrayRow = 0;
    let arrayColumn = 0;
    if(Array.isArray(e.data)){
      for(let i in e.data){
        if(Array.isArray(e.data[i])){
          arrayRow = arrayRow + 1;
          arrayColumn = e.data[i].length;
        }
      }
      sendData = e.data.join();
    }else{
      sendData = e.data;
    }
    let parameter = {
      sheetUrl: e.sheetUrl,
      sheetName: e.sheetName,
      type: e.type,
      data: sendData,
      range: e.range,
      arrayRow: arrayRow,
      arrayColumn: arrayColumn
    }
    $.ajax({
      type: "POST",
      data: parameter,
      url: "https://script.google.com/macros/s/AKfycbwBLLmy6R9TIQrzRmoJer7HVn1MXn20vfZ9d3WE9NY3LcIkm37j/exec",
      success: function (msg) {
        console.log(msg);
      }
    });
  }



  function sheetReadData(e, callback) {
    let parameter = {
      sheetUrl: e.sheetUrl,
      sheetName: e.sheetName
    }
    $.ajax({
      type: "POST",
      data: parameter,
      url: "https://script.google.com/macros/s/AKfycbyPI3r8HHEbLFFiXQvpRinUGrqSyYkfj5cvZNzshfsZEjIEEg4/exec",
      success: function (data) {
        callback(data);
      }
    });
  }



  function sheetFeature(e,action,type,num,addNum) {
    let parameter = {
      sheetUrl: e.sheetUrl,
      sheetName: e.sheetName,
      action: action,
      type: type,
      num: num,
      addNum: addNum,
    }
    $.ajax({
      type: "POST",
      data: parameter,
      url: "https://script.google.com/macros/s/AKfycbwTcoeTDtmubhyZIZ6_idtClzvRL5NH7xipOR0P4dBs2yvj1f2L/exec",
      success: function (msg) {
        console.log(msg);
      }
    });
  }



  window.sheetFeature = sheetFeature;
  window.sheetWriteData = sheetWriteData;
  window.sheetReadData = sheetReadData;
}(window, window.document));
