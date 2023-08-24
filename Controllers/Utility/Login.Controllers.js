var CryptoJS = require("crypto-js");
const jwt = require("jsonwebtoken");
const sql = require("mssql");

const { _tokenSecret } = require("./../../Config/token/TokenConfig.json");

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
          // console.log(`FOUND ${USER_IP}.`)
          // console.log(`COMAPRED ${USER_IP} AND ${MACHINE_IP}.`)
          res.json({ success: 6, data: `${request.recordset[0].USERID} can not login from this IP address.` })
        } else if (request.recordset[0].LINK == req.body.LINK) {
          if (!request.recordset[0].ISACCESS) {

            res.json({ success: 6, data: `${request.recordset[0].USERID} can not login from this url` })
          } else if (originalText == req.body.PASS) {
            jwt.sign({
              UserId: request.recordset[0].USERID,
              USER_FULLNAME: request.recordset[0].USER_FULLNAME,
              U_CAT: request.recordset[0].U_CAT,
              PROC_CODE: request.recordset[0].PROC_CODE,
              DEPT_CODE: request.recordset[0].DEPT_CODE,
              PNT: request.recordset[0].PNT
            }, _tokenSecret, { expiresIn: "12h" }, (err, token) => {
              res.json({ success: 1, data: token })
            });
          } else {
            res.json({ success: 3, data: "Password wrong" })
          }

        }
        else if (originalText == req.body.PASS) {
          jwt.sign({
            UserId: request.recordset[0].USERID,
            USER_FULLNAME: request.recordset[0].USER_FULLNAME,
            U_CAT: request.recordset[0].U_CAT,
            PROC_CODE: request.recordset[0].PROC_CODE,
            DEPT_CODE: request.recordset[0].DEPT_CODE,
            PNT: request.recordset[0].PNT
          }, _tokenSecret, { expiresIn: "12h" }, (err, token) => {
            res.json({ success: 1, data: token })
          });
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
