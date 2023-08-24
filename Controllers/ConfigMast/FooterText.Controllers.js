const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

var { msgServerconn } = require('../../config/db/rapsql');

exports.FootMsgFill = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        const pool = await msgServerconn();
        let request = await pool.request()



        request = await request.execute('USP_FootMsgFill');

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

exports.FootMsgSave = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        const pool = await msgServerconn();
        let request = await pool.request()

        request.input('DEPT_CODE', sql.VarChar(10), req.body.DEPT_CODE)
        request.input('MSG', sql.NVarChar(700), req.body.MSG)

        request = await request.execute('USP_FootMsgSave');

        res.json({ success: 1, data: '' })

      } catch (err) {
        console.log(err)
        res.json({ success: 0, data: err })
      }
    }
  });
}

