var CryptoJS = require("crypto-js");
const jwt = require("jsonwebtoken");
const sql = require("mssql");
var nodemailer = require('nodemailer');
var fs = require('fs');
let ejs = require("ejs");
const path = require("path")

const { _tokenSecret } = require("./../../Config/token/TokenConfig.json");
const { sendHtml } = require('./../Common/nodemailer.Controllers');

exports.LoginAuthentication = async (req, res) => {

  try {
    var request = new sql.Request();
    request.input('USERID', sql.VarChar(32), req.body.USERID)

    request = await request.execute('USP_UserMastFill');
    if (request.recordset) {
      if (request.recordset.length == 1) {
        let MACHINE_IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
        // MACHINE_IP = MACHINE_IP.split('::ffff:')[1]
        const USER_IP = request.recordset[0].IP
        var bytes = CryptoJS.AES.decrypt(request.recordset[0].U_PASS, process.env.PASS_KEY);
        var originalText = bytes.toString(CryptoJS.enc.Utf8);
        // console.log("pSS",originalText)
        // console.log('PASSWORD: ', originalText)

        let userid = ''
        userid = request.recordset[0].USERID
        // console.log(MACHINE_IP)
        if (request.recordset[0].RUN == true && userid.toUpperCase() != "ADMIN" && MACHINE_IP != request.recordset[0].USERIP) {

          // console.log(request.recordset)
          res.json({ success: 4, data: "User already logged in on IP address: " + request.recordset[0].USERIP })
          return;
        }

        if (USER_IP && MACHINE_IP !== USER_IP) {
          res.json({ success: 6, data: `${request.recordset[0].USERID} can not login from this IP address.` })
        } else if (request.recordset[0].LINK == req.body.LINK) {
          if (!request.recordset[0].ISACCESS) {

            res.json({ success: 6, data: `${request.recordset[0].USERID} can not login from this url` })
          } else if (originalText == req.body.PASS) {
            if (req.body.BASEURL != 'api.tender.peacocktech.in') {
              jwt.sign({
                UserId: request.recordset[0].USERID,
                USER_FULLNAME: request.recordset[0].USER_FULLNAME,
                U_CAT: request.recordset[0].U_CAT,
                PROC_CODE: request.recordset[0].PROC_CODE,
                DEPT_CODE: request.recordset[0].DEPT_CODE,
                PNT: request.recordset[0].PNT,
                EMAIL: request.recordset[0].EMAIL
              }, _tokenSecret, { expiresIn: "12h" }, (err, token) => {
                res.json({ success: 1, data: token })
              });
            } else {
              var CheckUuidReq = new sql.Request();

              CheckUuidReq.input('USERID', sql.VarChar(32), req.body.USERID)
              CheckUuidReq.input('IPADDRESS', sql.VarChar(64), MACHINE_IP)
              CheckUuidReq.input('MACADDRESS', sql.VarChar(128), req.body.uuid)
              CheckUuidReq.input('COMPUTERNAME', sql.VarChar(512), req.body.COMPUTERNAME)

              CheckUuidReq = await CheckUuidReq.execute('usp_UserLoginPermissionCheckIP');

              if (CheckUuidReq.recordset) {
                if (CheckUuidReq.recordset[0].ISACTIVE == true) {
                  jwt.sign({
                    UserId: request.recordset[0].USERID,
                    USER_FULLNAME: request.recordset[0].USER_FULLNAME,
                    U_CAT: request.recordset[0].U_CAT,
                    PROC_CODE: request.recordset[0].PROC_CODE,
                    DEPT_CODE: request.recordset[0].DEPT_CODE,
                    PNT: request.recordset[0].PNT,
                    EMAIL: request.recordset[0].EMAIL
                  }, _tokenSecret, { expiresIn: "12h" }, (err, token) => {
                    res.json({ success: 1, data: token })
                  });
                } else {
                  res.json({ success: 10, data: "Access denied." })
                }
              }
            }
          } else {
            res.json({ success: 3, data: "Password wrong" })
          }

        }
        else if (originalText == req.body.PASS) {
          if (req.body.BASEURL != 'api.tender.peacocktech.in') {
            jwt.sign({
              UserId: request.recordset[0].USERID,
              USER_FULLNAME: request.recordset[0].USER_FULLNAME,
              U_CAT: request.recordset[0].U_CAT,
              PROC_CODE: request.recordset[0].PROC_CODE,
              DEPT_CODE: request.recordset[0].DEPT_CODE,
              PNT: request.recordset[0].PNT,
              EMAIL: request.recordset[0].EMAIL,
            }, _tokenSecret, { expiresIn: "12h" }, (err, token) => {
              res.json({ success: 1, data: token })
            });
          }else{
            var CheckUuidReq = new sql.Request();
                CheckUuidReq.input('USERID', sql.VarChar(32), req.body.USERID)
                CheckUuidReq.input('IPADDRESS', sql.VarChar(64), MACHINE_IP)
                CheckUuidReq.input('MACADDRESS', sql.VarChar(128), req.body.uuid)
                CheckUuidReq.input('COMPUTERNAME', sql.VarChar(512), req.body.COMPUTERNAME)
  
                CheckUuidReq = await CheckUuidReq.execute('usp_UserLoginPermissionCheckIP');
  
                if (CheckUuidReq.recordset) {
                  if (CheckUuidReq.recordset[0].ISACTIVE == true) {
                    jwt.sign({
                      UserId: request.recordset[0].USERID,
                      USER_FULLNAME: request.recordset[0].USER_FULLNAME,
                      U_CAT: request.recordset[0].U_CAT,
                      PROC_CODE: request.recordset[0].PROC_CODE,
                      DEPT_CODE: request.recordset[0].DEPT_CODE,
                      PNT: request.recordset[0].PNT,
                      EMAIL: request.recordset[0].EMAIL,
                    }, _tokenSecret, { expiresIn: "12h" }, (err, token) => {
                      res.json({ success: 1, data: token })
                    });
                  } else {
                    res.json({ success: 10, data: "Access denied." })
                  }
                }
          }
        } else {
          res.json({ success: 3, data: "Password wrong" })
        }
      } else {
        res.json({ success: 5, data: "User not found" })
      }
    } else {
      res.json({ success: 2, data: "Not Found" })
    }
  } catch (err) {
    console.log(err);
    res.json({ success: 0, data: err })
  }
}

exports.UserFrmOpePer = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('USER_NAME', sql.VarChar(10), req.body.USER_NAME)
        request.input('FORM_NAME', sql.VarChar(50), req.body.FORM_NAME)

        request = await request.execute('usp_UserFrmOpePer');
        if (request.recordset) {
          res.json({ success: 1, data: request.recordset })
        } else {
          res.json({ success: 0, data: "Not Found" })
        }
      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}

exports.UserFrmPer = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('USER_NAME', sql.VarChar(10), req.body.USER_NAME)

        request = await request.execute('usp_UserFrmPer');
        if (request.recordset) {
          res.json({ success: 1, data: request.recordsets })
        } else {
          res.json({ success: 0, data: "Not Found" })
        }
      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}
exports.GetEmail = async (req, res) => {

      try {
        var request = new sql.Request();
        request.input('IUSER', sql.VarChar(10), req.body.IUSER)

        request = await request.execute('USP_GetEmail');
        if (request.recordset) {
          res.json({ success: 1, data: request.recordsets })
        } else {
          res.json({ success: 0, data: "Not Found" })
        }
      } catch (err) {
        console.log(err)
        res.json({ success: 0, data: err })
      }
    }

const crypt = (salt, text) => {
  const textToChars = (text) => text.split("").map((c) => c.charCodeAt(0));
  const byteHex = (n) => ("0" + Number(n).toString(16)).substr(-2);
  const applySaltToChar = (code) => textToChars(salt).reduce((a, b) => a ^ b, code);

  return text
    .split("")
    .map(textToChars)
    .map(applySaltToChar)
    .map(byteHex)
    .join("");
};

exports.EmailSendOTP = async(req,res) => {
  try{
      var val = Math.floor(100000 + Math.random() * 900000);
      let TemplateOTP = { OTP:val, year:new Date().getFullYear() , USERID:req.body.USERID};
      let OTPRes = false 
      for(let i =0;i<req.body.EMAIL.length;i++){
        OTPRes = await sendHtml(req.body.EMAIL[i]['EMAIL'], path.join(__dirname, "../../Public/MailTemplate/CRMOTP.ejs") ,TemplateOTP,'Tendar - OTP '+val )
      }
      console.log(val)
      res.json({success:OTPRes,data: crypt(process.env.PASS_KEY, val.toString()) })
  
  }catch(err){
    console.log(err)
      res.json({success:0,data:err})
  }
}


