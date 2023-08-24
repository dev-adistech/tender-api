const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.AssMastFill = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request = await request.execute('USP_AssMastFill');
        // console.log(request.recordset)

        if (request.recordset) {
          res.json({ success: 1, data: request.recordset })
        } else {
          res.json({ success: 0, data: "Not Found" })
        }

      } catch (err) {
        console.log(err)
        res.json({ success: 0, data: err })
      }
    }
  });
}

exports.AssMastSave = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('A_CODE', sql.Int, parseInt(req.body.A_CODE))
        request.input('A_NAME', sql.VarChar(15), req.body.A_NAME)
        request.input('TYPE', sql.VarChar(2), req.body.TYPE)
        request.input('ORD', sql.Int, parseInt(req.body.ORD))

        request = await request.execute('USP_AssMastSave');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}

exports.AssMastDelete = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        if (req.body.A_CODE) { request.input('A_CODE', sql.Int, parseInt(req.body.A_CODE)) }

        request = await request.execute('USP_AssMastDelete');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}
