const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.MakRateMastSave = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('FPER', sql.Decimal(10, 2), req.body.FPER)
        request.input('TPER', sql.Decimal(10, 2), req.body.TPER)
        request.input('SZ_CODE', sql.Int, parseInt(req.body.SZ_CODE))
        request.input('RATE', sql.Decimal(10, 2), req.body.RATE)
        request.input('TYPE', sql.VarChar(2), req.body.TYPE)

        request = await request.execute('usp_FrmMakeRateSave');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}


exports.MakRateMastFill = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();
        request.input('TYP_CODE', sql.VarChar(10), req.body.TYP_CODE)

        request = await request.execute('usp_FrmMakRateFill');

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


