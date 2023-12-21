const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');
var { preServerconn } = require('../../config/db/rapsql');

exports.StoneDetailView = async(req,res) => {
    
    jwt.verify(req.token,_tokenSecret, async (err, authData) => {
       if (err) {
           res.sendStatus(401);
       } else {
           const TokenData = await authData;

           try{
            const pool = await preServerconn();
            let request = await pool.request()

               request.input('STONEID',sql.VarChar(2056),req.body.STONEID);
               request.input('CertNo',sql.VarChar(2056),'');
               request.input('UserId',sql.VarChar(16),'');

               request = await request.execute('CRM_StoneDetailView');
           
               if(request.recordset){
                   res.json({success:1,data:request.recordset})
               }else{
                   res.json({success:0,data:"Not Found"})
               }

           }catch(err){
               res.json({success:0,data:err})
           }
       }
   });
}

exports.PartyEmailIdList = async(req,res) => {
    
    jwt.verify(req.token,_tokenSecret, async (err, authData) => {
       if (err) {
           res.sendStatus(401);
       } else {
           const TokenData = await authData;

           try{
            const pool = await preServerconn();
            let request = await pool.request()

               request.input('UserId',sql.VarChar(16),TokenData.UserId);
               request.input('UserCat',sql.VarChar(16),TokenData.Cat);

               request = await request.execute('SP_PartyEmailIdList');
           
               if(request.recordset){
                   res.json({success:1,data:request.recordset})
               }else{
                   res.json({success:0,data:"Not Found"})
               }

           }catch(err){
               res.json({success:0,data:err})
           }
       }
   });
}

exports.PacketOfferView = async(req,res) => {
    
    jwt.verify(req.token,_tokenSecret, async (err, authData) => {
       if (err) {
           res.sendStatus(401);
       } else {
           const TokenData = await authData;

           try{
            const pool = await preServerconn();
            let request = await pool.request()

               request.input('UserId',sql.VarChar(16),TokenData.UserId);
               request.input('StoneId',sql.Int,parseInt(req.body.StoneId));


               if(req.body.ISFANCY == 'true'){
                request.input('ISFANCY', sql.Bit, 1);
                }else if(req.body.ISFANCY == 'false'){
                    request.input('ISFANCY', sql.Bit, 0);
                }else {
                    request.input('ISFANCY', sql.Bit, req.body.ISFANCY);
                }

                request = await request.execute('SP_TrnPacketOfferView');

               if(request.recordset){
                   res.json({success:1,data:request.recordset})
               }else{
                   res.json({success:0,data:"Not Found"})
               }

           }catch(err){
               res.json({success:0,data:err})
           }
       }
   });
}


exports.OfferEntrySave = async(req,res) => {
    
    jwt.verify(req.token,_tokenSecret, async (err, authData) => {
       if (err) {
           res.sendStatus(401);
       } else {
           const TokenData = await authData;

           try{
            const pool = await preServerconn();
            let request = await pool.request()


               request.input('StoneId',sql.VarChar(50),req.body.StoneId);
               request.input('PId',sql.VarChar(50),req.body.PId);
               request.input('P_Code',sql.VarChar(16),req.body.P_Code);
               request.input('PartyName',sql.VarChar(100),req.body.PartyName);
               request.input('UserId',sql.VarChar(50),TokenData.UserId);
               request.input('OfferPer',sql.Decimal(10,2),parseFloat(req.body.OfferPer));
               request.input('OfferRate',sql.Decimal(10,2),parseFloat(req.body.OfferRate));
               request.input('GRap',sql.Decimal(10,2),parseFloat(req.body.GRap));
               request.input('GPer',sql.Decimal(10,2),parseFloat(req.body.GPer));
               request.input('GRate',sql.Decimal(10,2),parseFloat(req.body.GRate));
               request.input('OfferRemark',sql.VarChar(256),req.body.OfferRemark);
               request.input('E_Machine',sql.VarChar(50),TokenData.UserId);
               request.input('SP_CODE',sql.VarChar(50),TokenData.UserId);
               request.input('CR_CODE',sql.Int,parseInt(req.body.CR_CODE));

               request = await request.execute('USP_OfferEntrySave');


                res.json({success:1,data:""})

           }catch(err){
            console.log(err)
               res.json({success:0,data:err})
           }
       }
   });
}



exports.DownloadDoc = async(req,res) => {

   let https = require('https')
   https.get(req.body.url, function(response) {
    data = [];
    response.on('data', function(chunk) {
        data.push(chunk);
    });
    response.on('end', function() {
        data = Buffer.concat(data); // do something with data
        let DfileName = req.body.url.substr(req.body.url.lastIndexOf('/') + 1)
        res.writeHead(200, {
            'Content-Length': Buffer.byteLength(data),
            'Content-Type': 'application/octet-stream',
            'Content-disposition': 'attachment;filename='+DfileName}).end(data);
    }); 
 });
 data = [];
}