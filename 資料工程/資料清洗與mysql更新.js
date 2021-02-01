
/*<sql setting>*/
var connectionName = 'Instance_connection_name';
var rootPwd = 'root_password';
var user = 'user_name';
var userPwd = 'user_password';
var db = 'database_name';

var root = 'root';
var instanceUrl = 'jdbc:google:mysql://' + connectionName;
var dbUrl = instanceUrl + '/' + db;
/*</sql setting>*/

var urlRawData1 = '...'
var urlRawData2 = '...'

function main(){
  let tableName1 = '..'
  let tableName2 = '..'

  let table1 = getTable1().getUnique().removeExistedData(tableName1)
  let table2 = getTable2().getUnique().removeExistedData(tableName2)

  table1.writeIntoSql(tableName1)
  table2.writeIntoSql(tableName2)

}


function readFromSql(tableName) {

  let start = new Date();
  let stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  let results = stmt.executeQuery('SELECT * FROM ' + tableName);
  let numCols = results.getMetaData().getColumnCount();

  let output = []
  while (results.next()) {
    let row = [];
    for (var col = 0; col < numCols; col++) {
      row.push(results.getString(col + 1));
    }
    output.push(row)
  }

  results.close();
  stmt.close();

  var end = new Date();
  Logger.log('Time elapsed: %sms', end - start);
  return output
}

Array.prototype.getUnique = function(){
  /* @brief get unique value
   */
  return [...new Set(this.map(JSON.stringfy))].map(JSON.parse)
}

Array.prototype.removeExistedData = function(tableName){
  const existedData = readFromSql(tableName).map(JSON.stringify)
  return this.map(JSON.stringify).filter(row => !existedData.includes(row)).map(JSON.parse)
}

Array.prototype.writeIntoSql = function(tableName){
  conn.setAutoCommit(false);
  let start = new Date();
  
  let cols = this[0][0]
  let questionMarks = '?'
  this.slice(1, ).forEach(function(elem){
    cols = cols + ', ' + elem
    questionMarks = questionMarks + ', ?'
  })

  let stmt = conn.prepareStatement('INSERT INTO ' + tableName +
      + ' (' + cols + ') values (' + questionMarks + ')');

  this.slice(1,).forEach(function(row){
    cols.forEach(function(col, j){
      stmt.setString(j, row[j]);
    })
    stmt.addBatch();
  })

  var batch = stmt.executeBatch();
  conn.commit();
  conn.close();

  var end = new Date();
  Logger.log('Time elapsed: %sms for %s rows.', end - start, batch.length);
}

function boolHandler(input){
  let stringToZeroOne = {
    "是": 1,
    "否": 0,
    True: 1,
    False: 0    
  }
  return input ? stringToZeroOne[input] : ""
}

function dateHandler(input){ //把格式轉成 YYYY-MM-DD
  input = input.split(" ")[0]
  let y, m, d;
 if(!input){ //處理空值
   return "" 
 }else if(input.length == 5){ //處理格式 MM/DD
    y = "2020"
    input = input.split("/")
    m = input[0]
    d = input[1]
    return y + "-" + m + "-" + d
  }else if(input.match("[0-9]+/[0-9]+/[0-9]+")){
    input = input.split("/")
    y = input[0]
    m = input[1]
    d = input[2]
    return y + "-" + m + "-" + d
  }
}

function getUserId(name, card_id){
  const userTable = getTableFromSql("user_table")
  
  if(card_id){
    return userTable.filter(row => row[1] == name && row[2] == card_id)[0][0] 
  }else{
    return userTable.filter(row => row[1] == name )[0][0] 
  }
}


function getTable1(){
  const input = SpreadsheetApp.openByUrl(urlRawData1).getSheetByName('tableName1').getDataRange().getValues()

  let output = [['col1', 'col2', 'col3']]
  output = output.concat(
      input
      .slice(1,) //去除第一列欄位名稱
      .filter(row => !row[1]) //去除指定欄位無值之對象
      .map(function(row){      
        return [
          getUserId(row[1].toString().match("[\u4e00-\u9af5]+"), row[1].toString().match("[0-9]+")), //用中文名字, 身分證後四碼去查id
          boolHandler(row[2]), //處理布林值
          dateHandler(row[3])          
        ]
      })
  )
  return output
}


function getTable2(){
  const input = SpreadsheetApp.openByUrl(urlRawData2).getSheetByName('tableName2').getDataRange().getValues()

  let output = [['tag1', 'tag2', 'tag3']]
  let tag1 = new Set()
  let tag1 = new Set()
  let tag1 = new Set()

  input.slice(1,).forEach(function(row){
    row.slice(3,).filter(elem => !elem.toString().match(REF) || !elem.toString().match(NA)).forEach(function(elem){
      tag1 = elem.split("_")[0]
      tag2 = elem.split("_")[1]
      let tag3 = elem.split("_")[2]
      cache.push(tag1, tag2, tag3)
    })
  })
  tag1 = [...new Set(tag1)]
  tag2 = [...new Set(tag2)]
  tag3 = [...new Set(tag3)]

  tag1.forEach(tag1_ => 
    tag2.forEach(tag2_ =>
      tag3.forEach(tag3_ =>
        output.push([tag1_, tag2_, tag3_])
        )
      ))

}

//table3, 4, ...