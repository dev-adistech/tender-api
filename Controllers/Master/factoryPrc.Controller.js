const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.FacPrcFill = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)

        request = await request.execute('USP_FacPrcMastFill');

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

exports.facPrcSave = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)
        request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
        request.input('PRC_CODE', sql.VarChar(5), req.body.PRC_CODE)
        request.input('PRC_NAME', sql.VarChar(20), req.body.PRC_NAME)
        request.input('PRVPRC', sql.VarChar(100), req.body.PRVPRC)
        request.input('ORD', sql.Int, parseInt(req.body.ORD))

        request = await request.execute('USP_FacPrcMastSave');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}
