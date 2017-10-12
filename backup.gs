//var Cfg = {
//  getSheetName1 : ["S組織一覧","ユーザ組織対応表：元データ"],
//  stRow1 : [4,2], //元データ,取得データ
//  enCol1 : [2,9],
//  inputCol1 : [8,9,10],
//  getSheetName2 : ["S施設一覧","GSutie施設：元データ"],
//  stRow2 : [4,3],
//  enCol2:  [2,5],
//  inputCol2 : [8,9,10]
//};


//function onOpen(){
//  var ui = SpreadsheetApp.getUi();
//  ui.createMenu('CustomMenu')
//  .addItem('子組織格納(アド/カレ用組織)', 'fn_storesSosiki7')
//  .addItem('メンバー格納(アド/カレ用組織)', 'fn_storesMemberAddress')
//  .addItem('子組織格納(リソース組織)',   'fn_storesSosiki3')
//  .addItem('メンバー格納(リソース組織)',   'fn_storesMemberAddress2')
//  .addToUi();
//}




//
//
//
//function fn_storesSosiki3(){
//}
//  subfn_storesSosikiSub3();
//function fn_storesSosiki7(){
//  subfn_storesSosikiSub7();
//}
//
////組織ツリー構成処理
//function fn_storesMemberAddress() {
//  var inputValues = new Array();
//  var ss1 = SpreadsheetApp.getActive().getSheetByName(Cfg.getSheetName1[0]); //グループデータ
//  var ss2 = SpreadsheetApp.getActive().getSheetByName(Cfg.getSheetName1[1]); //メンバーデータ
//  
//  var enRow1 = ss1.getLastRow() - Cfg.stRow1[0] + 1;
//  var enRow2 = ss2.getLastRow() - Cfg.stRow1[1] + 1;
//  var stCol1 = 1;
//  var stCol2 = 1;
//  
//  var numGroup = ss1.getRange(Cfg.stRow1[0], stCol1, enRow1, Cfg.enCol1[0]).getValues();
//  var numMember = ss2.getRange(Cfg.stRow1[1], stCol2, enRow2, Cfg.enCol1[1]).getValues();
//  for ( var i in numGroup ) {
//    var targetGroup = String(numGroup[i][0]); //対象グループNo.
//    var memberList = new Array();
//    for ( var j in numMember ) {
//      var targetMember = String(numMember[j][8]); //対象メンバーNo.
//      if ( targetGroup == targetMember ){ //グループの記載Noとメンバー一覧の所属No一致かどうか
////        Logger.log(targetGroup + " / " +  targetMember);
//        memberList = memberList + numMember[j][1] + ","; //2列目メンバー追記
//      }
//    }
//    inputValues.push([memberList]);
//  }
////  Logger.log(inputValues);
//  ss1.getRange(Cfg.stRow1[0],Cfg.inputCol1[2],enRow1, 1).setValues(inputValues);
//}
//
////組織ツリー（施設用）構成処理
//function fn_storesMemberAddress2() {
//  var inputValues = new Array();
//  var ss1 = SpreadsheetApp.getActive().getSheetByName(Cfg.getSheetName2[0]); //グループデータ
//  var ss2 = SpreadsheetApp.getActive().getSheetByName(Cfg.getSheetName2[1]); //メンバーデータ
//
//  var enRow1 = ss1.getLastRow() - Cfg.stRow2[0] + 1;
//  var enRow2 = ss2.getLastRow() - Cfg.stRow2[1] + 1;
//  var stCol1 = 1;
//  var stCol2 = 26;
//
//  var numGroup = ss1.getRange(Cfg.stRow2[0], stCol1, enRow1, Cfg.enCol2[0]).getValues();
//  var numMember = ss2.getRange(Cfg.stRow2[1], stCol2, enRow2, Cfg.enCol2[1]).getValues();
//  for ( var i in numGroup ) {
//    var targetGroup = numGroup[i][0];
//    var memberList = new Array();
//    for ( var j in numMember ) {
//      var targetMember = String(numMember[j][0]); //対象メンバーNo.
//      if ( targetGroup == targetMember ){ //グループの記載Noとメンバー一覧の所属No一致かどうか
////        Logger.log(targetGroup + " / " +  targetMember);
//        memberList = memberList + numMember[j][4] + ","; //2列目メンバー追記
//      }
//    }
//    inputValues.push([memberList]);
//  }
//  ss1.getRange(Cfg.stRow2[0],Cfg.inputCol1[2],enRow1, 1).setValues(inputValues);
//}
//
//
////アドレス組織構造作成
//function subfn_storesSosikiSub7(sheetName) {
//  var inputRow = Cfg.stRow2[0];
//  var stWord = 5;
//  var endWord = 19;
//  var ss = SpreadsheetApp.getActive().getSheetByName(Cfg.getSheetName1[0]);
////  var ss = SpreadsheetApp.getActive().getSheetName();
////  if ( ss != sheetName ){
////    SpreadsheetApp.getActive().toast("info","対象外シートです。処理を中断します。",5);
////  }
//  var datas = ss.getRange(inputRow,2,ss.getLastRow() - inputRow + 1, 1).getValues();
//  
//  //#################################################################
//  var check, member = "";
//  var cnt = 0;
//  for ( var j = 0; j < datas.length; j++ ) {
//
//    var val = datas[j][0].slice(stWord,endWord);
//    
//    if ( val.slice(1,3) == "00" ){ //[1 [00 00 00 000 00 00]]
//      Logger.log("MATCH1-1: " + val);
//      var cnt2 = 0;
//      for ( var k in datas ) {
//        var val2 = datas[k][0].slice(stWord,endWord);
//        if ( val != val2 && //親ではない
//            val.slice(0,1) == val2.slice(0,1) && //親同じ
//          val2.slice(3,5) == "00" //子
//        ){
//          Logger.log("MATCH1-2: " + val2);
//          member = member + datas[k][0] + ",";
//          cnt2++;
//        }
//      }
//      ss.getRange(inputRow + cnt, Cfg.inputCol1[0]).setValue(member);
//      ss.getRange(inputRow + cnt, Cfg.inputCol1[1]).setValue("1 : " + cnt2);
//      member = "";
//      cnt++;
//      continue;
//    }
//    
//    if ( val.slice(1,3) != "00" && val.slice(3,5) == "00" ){ //1 [00 [00 00 000 00 00]]
//      Logger.log("MATCH2-1: " + val);
//      var cnt2 = 0;
//      for ( var k in datas ) {
//        var val2 = datas[k][0].slice(stWord,endWord);
//        if ( val != val2 && //親ではない
//            val.slice(0,3) == val2.slice(0,3) && //親同じ
//          val2.slice(5,7) == "00" //子
//        ){
//          Logger.log("MATCH2-2: " + val2);
//          member = member + datas[k][0] + ",";
//          cnt2++;
//        }
//      }
//      ss.getRange(inputRow + cnt, Cfg.inputCol1[0]).setValue(member);
//      ss.getRange(inputRow + cnt, Cfg.inputCol1[1]).setValue("2 : " + cnt2);
//      member = "";
//      cnt++;
//      continue;
//    }
//    
//    if ( val.slice(3,5) != "00" && val.slice(5,7) == "00" ){ //1 00 [00 [00 000 00 00]]
//      Logger.log("MATCH3-1: " + val);
//      var cnt2 = 0;
//      for ( var k in datas ) {
//        var val2 = datas[k][0].slice(stWord,endWord);
//        if ( val != val2 && //親ではない
//            val.slice(0,5) == val2.slice(0,5) && //親同じ
//          val2.slice(7,10) == "000" //子
//        ){
//          Logger.log("MATCH3-2: " + val2);
//          member = member + datas[k][0] + ",";
//          cnt2++;
//        }
//      }
//      ss.getRange(inputRow + cnt, Cfg.inputCol1[0]).setValue(member);
//      ss.getRange(inputRow + cnt, Cfg.inputCol1[1]).setValue("3 : " + cnt2);
//      member = "";
//      cnt++;
//      continue;
//    }
//    
//    if ( val.slice(5,7) != "00" && val.slice(7,10) == "000" ){ //1 00 00 [00 [000 00 00]]
//      Logger.log("MATCH4-1: " + val);
//      var cnt2 = 0;
//      for ( var k in datas ) {
//        var val2 = datas[k][0].slice(stWord,endWord);
//        if ( val != val2 && //親ではない
//            val.slice(0,7) == val2.slice(0,7) && //親同じ
//          val2.slice(10,12) == "00" //子
//        ){
//          Logger.log("MATCH4-2: " + val2);
//          member = member + datas[k][0] + ",";
//          cnt2++;
//        }
//      }
//      ss.getRange(inputRow + cnt, Cfg.inputCol1[0]).setValue(member);
//      ss.getRange(inputRow + cnt, Cfg.inputCol1[1]).setValue("4 : " + cnt2);
//      member = "";
//      cnt++;
//      continue;
//    }
//    
//    if ( val.slice(7,10) != "000" && val.slice(10,12) == "00" ){ //1 00 00 00 [000 [00 00]]
//      Logger.log("MATCH5-1: " + val);
//      var cnt2 = 0;
//      for ( var k in datas ) {
//        var val2 = datas[k][0].slice(stWord,endWord);
//        if ( val != val2 && //親ではない
//            val.slice(0,10) == val2.slice(0,10) && //親同じ
//          val2.slice(12,14) == "00" //子
//        ){
//          Logger.log("MATCH5-2: " + val2);
//          member = member + datas[k][0] + ",";
//          cnt2++;
//        }
//      }
//      ss.getRange(inputRow + cnt, Cfg.inputCol1[0]).setValue(member);
//      ss.getRange(inputRow + cnt, Cfg.inputCol1[1]).setValue("5 : " + cnt2);
//      member = "";
//      cnt++;
//      continue;
//    }
//    
//    if ( val.slice(10,12) != "00" && val.slice(12,14) == "00" ){ //1 00 00 00 000 [00 [00]]
//      Logger.log("MATCH6-1: " + val);
//      var cnt2 = 0;
//      for ( var k in datas ) {
//        var val2 = datas[k][0].slice(stWord,endWord);
//        if ( val != val2 && //親ではない
//            val.slice(0,12) == val2.slice(0,12) //親同じ
//        ){
//          Logger.log("MATCH6-2: " + val2);
//          member = member + datas[k][0] + ",";
//          cnt2++;
//        }
//      }
//      ss.getRange(inputRow + cnt, Cfg.inputCol1[0]).setValue(member);
//      ss.getRange(inputRow + cnt, Cfg.inputCol1[1]).setValue("6 : " + cnt2);
//      member = "";
//      cnt++;
//      continue;
//    }
//    
//    cnt++;
//  }
//}
////リソース組織構造作成
//function subfn_storesSosikiSub3(sheetName) {
//  var inputRow = Cfg.stRow1[0];
//  var stWord = 5;
//  var endWord = 13;
//  var ss = SpreadsheetApp.getActive().getSheetByName(Cfg.getSheetName2[0]);
////  var ss = SpreadsheetApp.getActive().getSheetName();
////  if ( ss != sheetName ){
////    SpreadsheetApp.getActive().toast("info","対象外シートです。処理を中断します。",5);
////  }
//  var datas = ss.getRange(inputRow,2,ss.getLastRow() - inputRow + 1, 1).getValues();
//  
//  //#################################################################
//  var check, member = "";
//  var cnt = 0;
//  for ( var j = 0; j < datas.length; j++ ) {
//    var val = datas[j][0].slice(stWord,endWord);
//    
//    if ( val.slice(2,4) == "00" ){ //[00 [00 00 00]]
//      Logger.log("MATCH1-1: " + val);
//      var cnt2 = 0;
//      for ( var k in datas ) {
//        var val2 = datas[k][0].slice(stWord,endWord);
//        if ( val != val2 && //親ではない
//            val.slice(0,2) == val2.slice(0,2) && //親同じ
//          val2.slice(4,6) == "00" //子
//        ){
//          Logger.log("MATCH1-2: " + val2);
//          member = member + datas[k][0] + ",";
//          cnt2++;
//        }
//      }
//      ss.getRange(inputRow + cnt, Cfg.inputCol2[0]).setValue(member);
//      ss.getRange(inputRow + cnt, Cfg.inputCol2[1]).setValue("1 : " + cnt2);
//      member = "";
//      cnt++;
//      continue;
//    }
//
//    if ( val.slice(2,4) != "00" && val.slice(4,6) == "00"  ){ //00 [00 [00 00]]
//      Logger.log("MATCH2-1: " + val);
//      var cnt2 = 0;
//      for ( var k in datas ) {
//        var val2 = datas[k][0].slice(stWord,endWord);
//        if ( val != val2 && //親ではない
//            val.slice(0,4) == val2.slice(0,4) && //親同じ
//          val2.slice(6) == "00" //子
//        ){
//          Logger.log("MATCH2-2: " + val2);
//          member = member + datas[k][0] + ",";
//          cnt2++;
//        }
//      }
//      ss.getRange(inputRow + cnt, Cfg.inputCol2[0]).setValue(member);
//      ss.getRange(inputRow + cnt, Cfg.inputCol2[1]).setValue("2 : " + cnt2);
//      member = "";
//      cnt++;
//      continue;
//    }    
//    if ( val.slice(4,6) != "00" && val.slice(6) == "00" ){ //00 00 [00 [00]]
//      Logger.log("MATCH3-1: " + val);
//      var cnt2 = 0;
//      for ( var k in datas ) {
//        var val2 = datas[k][0].slice(stWord,endWord);
//        if ( val != val2 && //親ではない
//            val.slice(0,6) == val2.slice(0,6) //親同じ
//        ){
//          Logger.log("MATCH3-2: " + val2);
//          member = member + datas[k][0] + ",";
//          cnt2++;
//        }
//      }
//      ss.getRange(inputRow + cnt, Cfg.inputCol2[0]).setValue(member);
//      ss.getRange(inputRow + cnt, Cfg.inputCol2[1]).setValue("3 : " + cnt2);
//      member = "";
//      cnt++;
//      continue;
//    }
//    
//    cnt++;
//  }
//}
