const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.LabourMastDelete = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('SRNO', sql.Int, req.body.SRNO)
        request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
        request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE)
        request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)

        request = await request.execute('USP_LabourMastDelete');

        res.json({ success: 1, data: request.recordset })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}

exports.LabourMastFill = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
        request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE)
        request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)


        request = await request.execute('USP_LabourMastFill');

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

exports.LabourMastSave = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('SRNO', sql.Int, req.body.SRNO)
        request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
        request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE)
        request.input('PRC_TYPE', sql.VarChar(5), req.body.PRC_TYPE)
        request.input('F_CARAT', sql.Numeric(10, 3), req.body.F_CARAT)
        request.input('T_CARAT', sql.Numeric(10, 3), req.body.T_CARAT)
        request.input('RATE', sql.Int, req.body.RATE)
        request.input('TYPE', sql.VarChar(2), req.body.TYPE)

        request = await request.execute('USP_LabourMastSave');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}



