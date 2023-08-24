const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.GrdParaMastFill = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
        request.input('CUT_CODE', sql.Int, parseInt(req.body.CUT_CODE))

        request = await request.execute('USP_GrdParamMastFill');

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

exports.GrdParaMastSave = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
        request.input('CUT_CODE', sql.Int, parseInt(req.body.CUT_CODE))
        request.input('FRATIO_CODE', sql.Numeric(10, 3), req.body.FRATIO_CODE)
        request.input('TRATIO_CODE', sql.Numeric(10, 3), req.body.TRATIO_CODE)
        request.input('FTDEPTH_CODE', sql.Numeric(10, 3), req.body.FTDEPTH_CODE)
        request.input('TTDEPTH_CODE', sql.Numeric(10, 3), req.body.TTDEPTH_CODE)
        request.input('FTABLE', sql.Numeric(10, 3), req.body.FTABLE)
        request.input('TTABLE', sql.Numeric(10, 3), req.body.TTABLE)
        request.input('FGIRDLE', sql.Numeric(10, 3), req.body.FGIRDLE)
        request.input('TGIRDLE', sql.Numeric(10, 3), req.body.TGIRDLE)

        request = await request.execute('USP_GrdParamMastSave');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}

exports.GrdParaMastDelete = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
        request.input('CUT_CODE', sql.Int, parseInt(req.body.CUT_CODE))

        request = await request.execute('USP_GrdParamMastDelete');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}
