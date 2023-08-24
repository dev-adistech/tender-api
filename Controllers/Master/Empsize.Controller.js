const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.EmpsizeFill = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request = await request.execute('usp_FrmEmpSizeFill');

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

exports.Empsizeupdate = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
        request.input('PROC_CODE', sql.Int, parseInt(req.body.PROC_CODE))
        request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
        request.input('S1', sql.Bit, req.body.S1)
        request.input('S2', sql.Bit, req.body.S2)
        request.input('S3', sql.Bit, req.body.S3)
        request.input('S4', sql.Bit, req.body.S4)
        request.input('S5', sql.Bit, req.body.S5)
        request.input('S6', sql.Bit, req.body.S6)
        request.input('S7', sql.Bit, req.body.S7)
        request.input('S8', sql.Bit, req.body.S8)
        request.input('S9', sql.Bit, req.body.S9)

        // console.log(req.body)
        request = await request.execute('usp_FrmEmpSizeUpdate');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}
