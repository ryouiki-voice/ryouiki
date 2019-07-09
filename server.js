const GoogleSpreadsheet = require('google-spreadsheet');
const http = require('http');
const fs = require('fs');

const SHEET_CRE = JSON.parse(fs.readFileSync('spread_sheet.json'));
//自分のGoogleSpreadSheetのIDに置き換える
const SHEET_ID = ' 1hu-7JH1idWFk2lER5J7ypzqUr8MpMR6hyF1l_Lcq1lE/edit#gid=0 ';

// fullfillment webhookが来た時にどのような応答を返すかを定義する
function response(con,replyCallback){
    //console.log(JSON.stringify(con.body.queryResult.intent));
    let action = con.body.queryResult.action;  //アクション名
    let param = con.body.queryResult.parameters; //パラメータ
    let userSpeech = con.body.queryResult.queryText; //ユーザーの発話
    
    //ここで応答を編集する
    
    if(action==='get-garbage'){
        getGabage(SHEET_ID,SHEET_CRE,param.Weekday)
        .then((value)=>{
            let msg = param.Weekday+"は"+value+"ゴミの日です";
            speech = makeSimpleResponse(msg,msg);
            replyCallback(speech);
        })
        .catch((err)=>{
            speech = makeSimpleResponse("エラーが発生しました",makeErrorMessage(err,con.body));
            replyCallback(speech);
        });
    }
    if(action==='set-garbage'){
        setGabage(SHEET_ID,SHEET_CRE,param.Weekday,param.Garbage)
        .then(()=>{
            let msg = param.Weekday+"に"+param.Garbage+"ゴミを登録しました";
            speech = makeSimpleResponse(msg,msg);
            replyCallback(speech);
        })
        .catch((err)=>{
            speech = makeSimpleResponse("エラーが発生しました",makeErrorMessage(err,con.body));
            replyCallback(speech);
        });
    }
}

// httpserverを起動してDialogFlowからWebhookを受け取れるようにする
http.createServer((req, res) => {
    Promise.resolve({
        req:req,
        res:res
    })
    .then((con)=>getBody(con))
    .then((con)=>response(con,function(s){
        res.writeHead(200, {'Content-Type': 'application/json; charset=utf-8'});
        res.end(JSON.stringify(s));
    }));
    
}).listen(process.env.PORT || 3000);

function getGabage(sheetId,credentials,weekday){
    return new Promise((resolve,reject)=>{
        authentication(sheetId,credentials)
        .then((gss)=>getSheets(gss))
        .then((sheets)=>{
            sheet = sheets.worksheets[0];
            return findCell(sheet,1,weekday); 
        })
        .then((cell)=>getCell(sheet,cell['row'],2))
        .then((cell)=>{
            resolve(cell['value']);
        })
        .catch((err)=>{
            reject(err);
        });

    });
}

function setGabage(sheetId,credentials,weekday,val){
    return new Promise((resolve,reject)=>{
        authentication(sheetId,credentials)
        .then((gss)=>getSheets(gss))
        .then((sheets)=>{
            sheet = sheets.worksheets[0];
            return findCell(sheet,1,weekday); 
        })
        .then((cell)=>getCell(sheet,cell['row'],2))
        .then((cell)=>{
            cell.setValue(val,()=>{
                resolve();
            });
        })
        .catch((err)=>{
            reject(err);
        });

    });
}

function authentication(sheetId,credentials){
    return new Promise((resolve,reject)=>{
        let gss = new GoogleSpreadsheet(sheetId);
        gss.useServiceAccountAuth(credentials, (err) => {
            if (!err) {
                resolve(gss);
            } else {
                reject(err);
            }
        });
    });
}

function getSheets(gss){
    return new Promise((resolve, reject) => {
        gss.getInfo((err, data) => {
            if (!err) {
                resolve(data);
            } else {
                reject(err);
            }
        });
    });
}

function findCell(sheet,col_index,value){
    return new Promise((resolve,reject)=>{
        new Promise((resolve,reject)=>{
            sheet.getCells({},(err,cells)=>{
                if(!err){
                    resolve(cells);
                }else{
                    reject(err);
                }
            })
        })
        .then((cells)=>{
            let isFind = false;
            cells.forEach((cell)=>{
                if(cell['col']==col_index){
                    if(cell['value']===value){
                        resolve(cell);
                        isFind = true;
                    }
                }
            });
            if(!isFind){
                reject('cannot find cell');
            }
        })
        .catch((err)=>{
            reject(err);
        });
        
    });
}

function getCell(sheet,row_index,col_index){
    return new Promise((resolve,reject)=>{
        sheet.getCells({},(err,cells)=>{
            if(err){
                reject(err);
            }
            let isFind = false;
            cells.forEach((cell)=>{
                if(cell['col']==col_index&&cell['row']==row_index){
                    resolve(cell);
                }
            });
            if(!isFind){
                reject('cannot find cell');
            }
        });
    });
}


function makeSimpleResponse(speech,displayText){
    return {
        fulfillmentText:speech
    }
}


function makeErrorMessage(err,body){
    return "エラーが発生しました "+"["+err+"] "+JSON.stringify(body.result,null," ");
}

function getBody(con){
    return new Promise((resolve,reject)=>{
        let body = '';
        con.req.on('data', function (data) {
            body += data;
        });
        con.req.on('end', function () {
            if(body!==''){
                con.body = JSON.parse(body);
            }else{
                con.body = {};
            }
            resolve(con);
        });
    });
}
