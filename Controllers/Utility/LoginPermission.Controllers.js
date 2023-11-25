const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');


exports.CompPerMastIPFill = async(req,res) => {
    
    jwt.verify(req.token,_tokenSecret, async (err, authData) => {
       if (err) {
           res.sendStatus(401);
       } else {
           const TokenData = await authData;

           try{
               var request  = new sql.Request();
                
               if(req.body.USERID){ request.input('USERID',sql.VarChar(32),req.body.USERID) }
               if(req.body.F_DATE){ request.input('F_DATE',sql.DateTime2,new Date(req.body.F_DATE) )}
               if(req.body.T_DATE){ request.input('T_DATE',sql.DateTime2,new Date(req.body.T_DATE) )}
                
               request = await request.execute('USP_CompPerMastIPFill');
            
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

exports.CompPerMastIPSave = async(req,res) => {
    
    jwt.verify(req.token,_tokenSecret, async (err, authData) => {
       if (err) {
           res.sendStatus(401);
       } else {
           const TokenData = await authData;
           
           try{
               var request  = new sql.Request();
               request.input('ANO',sql.Int,parseInt(req.body.ANO))
               request.input('USERID',sql.VarChar(32),req.body.USERID)
               request.input('IPADDRESS',sql.VarChar(128),req.body.IPADDRESS)
               request.input('MACADDRESS',sql.VarChar(128),req.body.MACADDRESS)
               request.input('ISACTIVE',sql.Bit,req.body.ISACTIVE)
               request.input('REMARK',sql.VarChar(128),req.body.REMARK)
               request.input('APPROVEUSERID',sql.VarChar(32),TokenData.UserId)

               request = await request.execute('USP_CompPerMastIPSave');
               
               res.json({success:1,data:''})

           }catch(err){
            res.json({success:0,data: err})
           }
       }
   });
}


exports.CompPerMastIPDelete = async(req,res) => {
    
    jwt.verify(req.token,_tokenSecret, async (err, authData) => {
       if (err) {
           res.sendStatus(401);
       } else {
           const TokenData = await authData;
           
           try{
               var request  = new sql.Request();

               request.input('ANO',sql.Int,parseInt(req.body.ANO))
               
               request = await request.execute('USP_CompPerMastIPDelete');
               
               res.json({success:1,data:''})

           }catch(err){
            res.json({success:0,data: err})
           }
       }
   });
}
