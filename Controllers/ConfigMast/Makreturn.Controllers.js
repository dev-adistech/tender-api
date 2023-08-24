const sql = require("mssql");
const jwt = require("jsonwebtoken");

var { _tokenSecret } = require("../../Config/token/TokenConfig.json");
const { request } = require("express");

exports.makreturnFill = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();
        let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;

        request.input("L_CODE", sql.VarChar(10), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("TAG", sql.VarChar(5), req.body.TAG);
        request.input("IUSER", sql.VarChar(10), req.body.IUSER);
        request.input("ICOMP", sql.VarChar(30), IP);


        request = await request.execute('USP_MakReturn')

        res.json({ success: 1, data: request.recordsets })

      } catch (err) {
        console.log(err);
        res.json({ success: 0, data: err });
      }
    }
  });
};
