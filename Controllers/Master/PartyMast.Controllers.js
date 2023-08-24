const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.PartyMastFill = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
        request.input('PROC_CODE', sql.Int, parseInt(req.body.PROC_CODE))

        request = await request.execute('USP_PartyMastFill');

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

exports.PartyMastSave = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
        request.input('PROC_CODE', sql.Int, parseInt(req.body.PROC_CODE))
        request.input('COMP_CODE', sql.VarChar(5), req.body.COMP_CODE)
        request.input('COMP_NAME', sql.VarChar(50), req.body.COMP_NAME)
        request.input('OWNER_NAME', sql.VarChar(50), req.body.OWNER_NAME)
        request.input('ADDRESS', sql.VarChar(500), req.body.ADDRESS)
        request.input('MOB1', sql.VarChar(10), req.body.MOB1)
        request.input('MOB2', sql.VarChar(10), req.body.MOB2)
        request.input('GSTNO', sql.VarChar(50), req.body.GSTNO)
        request.input('HSNCODE', sql.VarChar(50), req.body.HSNCODE)

        request = await request.execute('USP_PartyMastSave');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}

exports.PartyMastDelete = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
        request.input('PROC_CODE', sql.Int, parseInt(req.body.PROC_CODE))
        request.input('COMP_CODE', sql.VarChar(5), req.body.COMP_CODE)

        request = await request.execute('USP_PartyMastDelete');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}
