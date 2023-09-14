const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.TendarPrdDetDisp = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))

                request = await request.execute('USP_TendarPrdDetDisp');

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

exports.FindRap = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
      if (err) {
        res.sendStatus(401);
      } else {
        const TokenData = await authData;
        try {
          let request = new sql.Request();
  
          request.input("S_CODE", sql.VarChar(5), req.body.S_CODE);
          request.input("Q_CODE", sql.Int, parseInt(req.body.Q_CODE));
          request.input("C_CODE", sql.Int, parseInt(req.body.C_CODE));
          request.input("CARAT", sql.Decimal(10, 3), parseFloat(req.body.CARAT));
          request.input("CUT_CODE", sql.Int, parseInt(req.body.CUT_CODE));
          request.input("FL_CODE", sql.Int, parseInt(req.body.FL_CODE));
          request.input("IN_CODE", sql.Int, parseInt(req.body.IN_CODE));
          request.input("RTYPE", sql.VarChar(5), req.body.RTYPE);
          request.input("RAPTYPE", sql.VarChar(5), req.body.RAPTYPE);
          request.input("ML_CODE", sql.Int, req.body.ML_CODE);
          request.input("SH_CODE", sql.Int, req.body.SH_CODE);
          request.input("REF_CODE", sql.Int, req.body.REF_CODE);
          
          
          request = await request.execute("USP_FindRap");
  
          if (request.recordsets) {
            res.json({ success: 1, data: request.recordsets });
          } else {
            res.json({ success: 0, data: "Not Found" });
          }
        } catch (err) {
          console.log(err)
          res.json({ success: 0, data: err });
        }
      }
    });
  };

  exports.TendarPrdDetSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
            let PerArr = req.body;
            let status = true;
            let ErrorPer = []
            for (let i = 0; i < PerArr.length; i++) {
                try {
                    var request = new sql.Request();

                    request.input('COMP_CODE', sql.VarChar(10), PerArr[i].COMP_CODE)
                    request.input('DETID', sql.Int, parseInt(PerArr[i].DETID))
                    request.input('T_DATE', sql.DateTime2, new Date(PerArr[i].T_DATE))
                    request.input('SRNO', sql.Int, parseInt(PerArr[i].SRNO))
                    request.input('PLANNO', sql.Int, parseInt(PerArr[i].PLANNO))
                    request.input('PTAG', sql.VarChar(2), PerArr[i].PTAG)
                    request.input('I_CARAT', sql.Numeric(10,3), PerArr[i].I_CARAT)
                    request.input('S_CODE', sql.VarChar(5), PerArr[i].S_CODE)
                    request.input('C_CODE', sql.Int, parseInt(PerArr[i].C_CODE))
                    request.input('Q_CODE', sql.Int, parseInt(PerArr[i].Q_CODE))
                    request.input('CARAT', sql.Numeric(10,3), PerArr[i].CARAT)
                    request.input('CT_CODE', sql.Int, parseInt(PerArr[i].CT_CODE))
                    request.input('FL_CODE', sql.Int, parseInt(PerArr[i].FL_CODE))
                    request.input('LB_CODE', sql.VarChar(2), PerArr[i].LB_CODE)
                    request.input('IN_CODE', sql.Int, parseInt(PerArr[i].IN_CODE))
                    request.input('ADNO', sql.Int, parseInt(PerArr[i].ADNO))
                    request.input('PLNSEL', sql.Bit, PerArr[i].PLNSEL)
                    request.input('RTYPE', sql.VarChar(2), PerArr[i].RTYPE)
                    request.input('ORAP', sql.Numeric(10,3), PerArr[i].ORAP)
                    request.input('RATE', sql.Numeric(10,3), PerArr[i].RATE)
                    request.input('IUSER', sql.VarChar(20), PerArr[i].IUSER)
                    if(PerArr[i].IUSER){
                    request.input('ICOMP', sql.VarChar(30), IP)
                    }else{
                      request.input('ICOMP', sql.VarChar(30), '')
                    }
                    request.input('OTAG', sql.VarChar(2), PerArr[i].OTAG)
                    request.input('ML_CODE', sql.Int, parseInt(PerArr[i].ML_CODE))
                    request.input('DEP_CODE', sql.Numeric(10,2), PerArr[i].DEP_CODE)
                    request.input('RAT_CODE', sql.Numeric(10,2), PerArr[i].RAT_CODE)
                    request.input('GRD_CODE', sql.Numeric(10,2), PerArr[i].GRD_CODE)
                    request.input('MPER', sql.Numeric(10,2), PerArr[i].MPER)
                    request.input('SH_CODE', sql.Int, PerArr[i].SH_CODE)
                    request.input('REF_CODE', sql.Int, PerArr[i].REF_CODE)
                    request.input('RAPTYPE', sql.VarChar(5), PerArr[i].RAPTYPE)

                    request = await request.execute('USP_TendarPrdDetSave');
                } catch (err) {
                    status = false;
                    ErrorPer.push(PerArr[i])
                }
            }

            res.json({ success: status, data: ErrorPer })
        }
    });
}

exports.TendarVidUpload = async (req, res) => {
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
        request.input("SECURE_URL", sql.VarChar(500), req.body.SECURE_URL);
        request.input("URL", sql.VarChar(500), req.body.URL);
        request.input("CLOUDID", sql.VarChar(500), req.body.CLOUDID);
        request.input("PUBLICID", sql.VarChar(500), req.body.PUBLICID);
        request.input("I_TYPE", sql.VarChar(25), req.body.I_TYPE);

        request = await request.execute("USP_TendarVidUpload");

        res.json({ success: 1, data: "" });
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.TendarVidDisp = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("COMP_CODE", sql.VarChar(10), req.body.COMP_CODE);
        request.input("DETID", sql.Int, parseInt(req.body.DETID));
        request.input("FSRNO", sql.Int, parseInt(req.body.FSRNO));
        request.input("TSRNO", sql.Int, parseInt(req.body.TSRNO));
        request.input("I_TYPE", sql.VarChar(25), req.body.I_TYPE);

        request = await request.execute("USP_TendarVidDisp");

        if (request.recordset) {
          res.json({ success: 1, data: request.recordset });
        } else {
          res.json({ success: 0, data: "Not Found" });
        }
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.TendarVidDelete = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("COMP_CODE", sql.VarChar(10), req.body.COMP_CODE);
        request.input("DETID", sql.Int, parseInt(req.body.DETID));
        request.input("FSRNO", sql.Int, parseInt(req.body.FSRNO));
        request.input("TSRNO", sql.Int, parseInt(req.body.TSRNO));
        request.input("I_TYPE", sql.VarChar(25), req.body.I_TYPE);

        request = await request.execute("USP_TendarVidDelete");

        res.json({ success: 1, data: request.recordset });
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.TendarPrdDetDock = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("COMP_CODE", sql.VarChar(10), req.body.COMP_CODE);
        request.input("DETID", sql.Int, parseInt(req.body.DETID));

        request = await request.execute("USP_TendarPrdDetDock");

        if (request.recordset) {
          res.json({ success: 1, data: request.recordset });
        } else {
          res.json({ success: 0, data: "Not Found" });
        }
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.TendarResSave = async (req, res) => {
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

        request = await request.execute("USP_TendarResSave");

        res.json({ success: 1, data: "" });
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};
exports.TendarApprove = async (req, res) => {
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
        request.input("TEN_NAME", sql.VarChar(50), req.body.TEN_NAME);
        request.input("ISAPPROVE", sql.Bit, req.body.ISAPPROVE);
        request.input("AUSER", sql.VarChar(20), req.body.AUSER);
        request.input("ACOMP", sql.VarChar(30), IP);
       

        request = await request.execute("USP_TendarApprove");

        res.json({ success: 1, data: "" });
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.TendarVidUploadDisp = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
      if (err) {
          res.sendStatus(401);
      } else {
          const TokenData = await authData;

          try {
              var request = new sql.Request();

              request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
              request.input('DETID', sql.Int, parseInt(req.body.DETID))
              request.input('SRNO', sql.Int, parseInt(req.body.SRNO))

              request = await request.execute('USP_TendarVidUploadDisp');

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