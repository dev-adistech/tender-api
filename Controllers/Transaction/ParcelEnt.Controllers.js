const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.TendarParcelEnt = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)

                request = await request.execute('USP_TendarParcelEnt');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordsets })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.TendarResParcelSave = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
      if (err) {
        res.sendStatus(401);
      } else {
        const TokenData = await authData;
  
        try {
          var request = new sql.Request();
  
          let IP = req.headers["x-forwarded-for"] || req.connection.remoteAddress;
  
          request.input("COMP_CODE", sql.VarChar(10), req.body.COMP_CODE);
          request.input("DETID", sql.Int, parseInt(req.body.DETID));
          request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
          request.input("I_CARAT", sql.Numeric(10,3), req.body.I_CARAT);
          request.input("RESRVE", sql.Numeric(10,2), req.body.RESRVE);
          request.input("PERCTS", sql.Numeric(10,2), req.body.PERCTS);
          request.input("SRW", sql.VarChar(50), req.body.SRW);
          request.input("FL_CODE", sql.Int, parseInt(req.body.FL_CODE));
          request.input("FBID", sql.Numeric(10,2), req.body.FBID);
          request.input("T_CODE", sql.VarChar(5), req.body.T_CODE);
          request.input("LS", sql.Bit, req.body.LS);
          request.input("FFLAT1", sql.VarChar(10), req.body.FFLAT1);
          request.input("FFLAT2", sql.VarChar(10), req.body.FFLAT2);
          request.input("FMED", sql.VarChar(10), req.body.FMED);
          request.input("FHIGH", sql.VarChar(10), req.body.FHIGH);
          request.input("RFLAT1", sql.VarChar(10), req.body.RFLAT1);
          request.input("RFLAT2", sql.VarChar(10), req.body.RFLAT2);
          request.input("RMED", sql.VarChar(10), req.body.RMED);
          request.input("RHIGH", sql.VarChar(10), req.body.RHIGH);
          request.input("MFLFLAT1", sql.VarChar(10), req.body.MFLFLAT1);
          request.input("MFLFLAT2", sql.VarChar(10), req.body.MFLFLAT2);
          request.input("MFLMED", sql.VarChar(10), req.body.MFLMED);
          request.input("MFLHIGH", sql.VarChar(10), req.body.MFLHIGH);
          request.input("FLNFLAT1", sql.VarChar(10), req.body.FLNFLAT1);
          request.input("FLNFLAT2", sql.VarChar(10), req.body.FLNFLAT2);
          request.input("FLNMED", sql.VarChar(10), req.body.FLNMED);
          request.input("FLNHIGH", sql.VarChar(10), req.body.FLNHIGH);
          request.input("CFLAT1", sql.VarChar(100), req.body.CFLAT1);
          request.input("CFLAT2", sql.VarChar(100), req.body.CFLAT2);
          request.input("CMED", sql.VarChar(100), req.body.CMED);
          request.input("CHIGH", sql.VarChar(100), req.body.CHIGH);
          request.input("DNC_CODE", sql.SmallInt, parseInt(req.body.DNC_CODE));
          request.input("I1C_CODE", sql.SmallInt, parseInt(req.body.I1C_CODE));
          request.input("I2C_CODE", sql.SmallInt, parseInt(req.body.I2C_CODE));
          request.input("I3C_CODE", sql.SmallInt, parseInt(req.body.I3C_CODE));
          request.input("RC_CODE", sql.SmallInt, parseInt(req.body.RC_CODE));
          request.input("R1C_CODE", sql.SmallInt, parseInt(req.body.R1C_CODE));
          request.input("R2C_CODE", sql.SmallInt, parseInt(req.body.R2C_CODE));
          request.input("FC_CODE", sql.SmallInt, parseInt(req.body.FC_CODE));
          request.input("F1C_CODE", sql.SmallInt, parseInt(req.body.F1C_CODE));
          request.input("F2C_CODE", sql.SmallInt, parseInt(req.body.F2C_CODE));
          request.input("PUSER", sql.VarChar(20), req.body.PUSER);
          request.input("PCOMP", sql.VarChar(30), IP);
          request.input("TEN_NAME",sql.VarChar(50),req.body.TEN_NAME)
          request.input("ADIS",sql.Numeric(10,2),req.body.ADIS)
          request.input("FAMT",sql.Numeric(10,2),req.body.FAMT)
          request.input("UUSER1", sql.VarChar(20), req.body.UUSER1);
          request.input("UUSER2", sql.VarChar(20), req.body.UUSER2);
          request.input("UUSER3", sql.VarChar(20), req.body.UUSER3);
          request.input("ISBV", sql.Bit, req.body.ISBV);
          request.input("BVCOMMENT", sql.VarChar(500), req.body.BVCOMMENT);
  
          request = await request.execute("USP_TendarResParcelSave");
          res.json({ success: 1, data: "" });
        } catch (err) {
          res.json({ success: 0, data: err });
        }
      }
    });
  };

exports.TendarMastDisSave = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
      if (err) {
        res.sendStatus(401);
      } else {
        const TokenData = await authData;
  
        try {
          var request = new sql.Request();
  
          let IP = req.headers["x-forwarded-for"] || req.connection.remoteAddress;
  
          request.input("COMP_CODE", sql.VarChar(10), req.body.COMP_CODE);
          request.input("DETID", sql.Int, parseInt(req.body.DETID));
          request.input("ADIS", sql.Numeric(10,2), req.body.ADIS);
          request.input("TRESRVE", sql.Numeric(10,2), req.body.TRESRVE);
          
  
          request = await request.execute("USP_TendarMastDisSave");
          res.json({ success: 1, data: "" });
        } catch (err) {
          res.json({ success: 0, data: err });
        }
      }
    });
  };