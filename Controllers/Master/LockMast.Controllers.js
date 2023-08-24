const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');
exports.LockMastFill = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('LOCK_TYPE', sql.VarChar(2), req.body.LOCK_TYPE);

        request = await request.execute('USP_LockMastFill');

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
exports.LockMastDelete = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('LOCK_TYPE', sql.VarChar(2), req.body.LOCK_TYPE);
        request.input('SRNO', sql.Int, req.body.SRNO);
        request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE);

        request = await request.execute('USP_LockMastDelete');

        res.json({ success: 1, data: request.recordset })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}
exports.LockMastSave = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('LOCK_TYPE', sql.VarChar(2), req.body.LOCK_TYPE);
        request.input('SRNO', sql.Int, parseInt(req.body.SRNO));
        request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE);
        request.input('PROC_NAME', sql.VarChar(50), req.body.PROC_NAME);
        request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE);
        request.input('HRS', sql.Numeric(10, 2), req.body.HRS);
        request.input('ORD', sql.Int, parseInt(req.body.ORD));

        request = await request.execute('USP_LockMastSave');

        res.json({ success: 1, data: request.recordset })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}
