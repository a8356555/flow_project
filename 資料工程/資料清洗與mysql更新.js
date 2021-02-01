
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


var Table = function(url, tableName){
  /* @brief 自訂型別
   * @attr 網址, sql表名, 及資料本身
   */
  this.url = url
  this.name = tableName
  this.data = ''
}

Table.prototype.getUnique = function(){
  /* @brief get unique value
   */
  return [...new Set(this.data.map(JSON.stringfy))].map(JSON.parse)
}

Table.prototype.removeExistedData = function(){
  /* @brief 資料比對, 去除已存在mysql之資料
   */
  const existedData = readFromSql(this.name).map(JSON.stringify)
  return this.data.map(JSON.stringify).filter(row => !existedData.includes(row)).map(JSON.parse)
}

Table.prototype.writeIntoSql = function(){
  /* @brief 寫入mysql
   */
  conn.setAutoCommit(false);
  let start = new Date();
  
  let cols = this.data[0][0]
  let questionMarks = '?'
  this.data.slice(1, ).forEach(function(elem){
    cols = cols + ', ' + elem
    questionMarks = questionMarks + ', ?'
  })

  let stmt = conn.prepareStatement('INSERT INTO ' + this.name +
      + ' (' + cols + ') values (' + questionMarks + ')');

  this.data.slice(1,).forEach(function(row){
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


function boolHandler(input){
  /* @brief 處理布林值
   * @param input 輸入為string
   * @return 輸出為0/1
   */
  let stringToZeroOne = {
    "是": 1,
    "否": 0,
    True: 1,
    False: 0    
  }
  return input ? stringToZeroOne[input] : ""
}

function dateHandler(input){ 
  /* @brief 處理格式把格式轉成 YYYY-MM-DD
   * @param input 輸入格式可能為MM/DD YYYY/MM/DD
   */
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
  /* @brief 從Sql比對出user id
   * @param name, card_id 用名字、身分證後四碼比對
   */
  const userTable = readFromSql("user_table")
  
  if(card_id){
    return userTable.filter(row => row[1] == name && row[2] == card_id)[0][0] 
  }else{
    return userTable.filter(row => row[1] == name )[0][0] 
  }
}


function main(){
  let urlRawData1 = '...'
  let urlRawData2 = '...'
  let table1 = new Table(urlRawData1, 'tableName1')
  let table2 = new Table(urlRawData2, 'tableName2')

  table1.getTable = function(){
    const input = SpreadsheetApp.openByUrl(this.url).getSheetByName(this.name).getDataRange().getValues()

    this.data = [['col1', 'col2', 'col3']]
    this.data = this.data.concat(
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
  }
  
  
  table2.getTable = function(){
    const input = SpreadsheetApp.openByUrl(this.url).getSheetByName(this.name).getDataRange().getValues()

    this.data = [['tag1', 'tag2', 'tag3']]
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
          this.data.push([tag1_, tag2_, tag3_])
          )
        ))
  }
  
  
  table1.getTable().getUnique().removeExistedData().writeIntoSql()
  table2.getTable().getUnique().removeExistedData().writeIntoSql()
}


//table3, 4, ...