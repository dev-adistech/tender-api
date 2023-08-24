const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.LooseLabMastFill = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request = await request.execute('USP_FrmLooseLabFill');

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


exports.LooseLabMastSave = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('C_CODE', sql.VarChar(200), req.body.C_CODE)
        request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
        request.input('F_SIZE', sql.Numeric(10, 3), req.body.F_SIZE)
        request.input('T_SIZE', sql.Numeric(10, 3), req.body.T_SIZE)
        request.input('RATE', sql.Numeric(10, 3), req.body.RATE)
        request.input('RAPTYPE', sql.VarChar(5), req.body.RAPTYPE)

        // console.log(req.body)
        request = await request.execute('USP_FrmLooseLabSave');

        res.json({ success: 1, data: '' })

      } catch (err) {
        console.log(err)
        res.json({ success: 0, data: err })
      }
    }
  });
}

exports.LooseLabMastDelete = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
        request.input('F_SIZE', sql.Numeric(10, 3), req.body.F_SIZE)
        request.input('T_SIZE', sql.Numeric(10, 3), req.body.T_SIZE)

        request = await request.execute('USP_FrmLooseLabDelete');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}
