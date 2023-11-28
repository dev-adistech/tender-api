const sql = require("mssql");
const jwt = require("jsonwebtoken");

var { _tokenSecret } = require("../../Config/token/TokenConfig.json");

var { RapServerConn } = require("../../config/db/rapsql");

exports.RapCalcDisp = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(5), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("TAG", sql.VarChar(2), req.body.TAG);
        request.input("EMP_CODE", sql.VarChar(5), req.body.EMP_CODE);
        request.input("IUSER", sql.VarChar(10), req.body.IUSER);

        request = await request.execute("USP_RapCalDisp");
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

exports.PrdEmpFill = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(10), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("TAG", sql.VarChar(5), req.body.TAG);
        request.input("IUSER", sql.VarChar(10), req.body.IUSER);
        request.input("OEMP_CODE", sql.VarChar(10), req.body.OEMP_CODE);

        request = await request.execute("USP_PrdEmpFill");
        if (request.recordsets) {
          res.json({ success: 1, data: request.recordsets });
        } else {
          res.json({ success: 0, data: "Not Found" });
        }
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.RapCalcCrtFill = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(10), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("TAG", sql.VarChar(5), req.body.TAG);

        request = await request.execute("usp_RapCalcCrtFill");
        if (request.recordsets) {
          res.json({ success: 1, data: request.recordsets });
        } else {
          res.json({ success: 0, data: "Not Found" });
        }
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.FindRap = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;
      // console.log("findrap", req.body)
      try {
        // const pool = new sql.Request();
        let request = new sql.Request();

        request.input("S_CODE", sql.VarChar(5), req.body.S_CODE);
        request.input("Q_CODE", sql.Int, parseInt(req.body.Q_CODE));
        request.input("C_CODE", sql.Int, parseInt(req.body.C_CODE));
        request.input("CARAT", sql.Decimal(10, 3), parseFloat(req.body.CARAT));
        request.input("CUT_CODE", sql.Int, parseInt(req.body.CUT_CODE));
        request.input("POL_CODE", sql.Int, parseInt(req.body.POL_CODE));
        request.input("SYM_CODE", sql.Int, parseInt(req.body.SYM_CODE));
        request.input("FL_CODE", sql.Int, parseInt(req.body.FL_CODE));
        request.input("IN_CODE", sql.Int, parseInt(req.body.IN_CODE));
        request.input("SH_CODE", sql.Int, parseInt(req.body.SH_CODE));
        request.input("TABLE", sql.Int, parseInt(req.body.TABLE));
        request.input("TABLE_BLACK", sql.Int, parseInt(req.body.TABLE_BLACK));
        request.input("TABLE_OPEN", sql.Int, parseInt(req.body.TABLE_OPEN));
        request.input("SIDE", sql.Int, parseInt(req.body.SIDE));
        request.input("SIDE_BLACK", sql.Int, parseInt(req.body.SIDE_BLACK));
        request.input("SIDE_OPEN", sql.Int, parseInt(req.body.SIDE_OPEN));
        request.input("CROWN_OPEN", sql.Int, parseInt(req.body.CROWN_OPEN));
        request.input("GIRDLE_OPEN", sql.Int, parseInt(req.body.GIRDLE_OPEN));
        request.input("PAV_OPEN", sql.Int, parseInt(req.body.PAV_OPEN));
        request.input("CULET", sql.Int, parseInt(req.body.CULET));
        request.input("EXTFACET", sql.Int, parseInt(req.body.EXTFACET));
        request.input("EYECLAEN", sql.Int, parseInt(req.body.EYECLAEN));
        request.input("GRAINING", sql.Int, parseInt(req.body.GRAINING));
        request.input("LUSTER", sql.Int, parseInt(req.body.LUSTER));
        request.input("MILKY", sql.Int, parseInt(req.body.MILKY));
        request.input("NATURAL", sql.Int, parseInt(req.body.NATURAL));
        request.input("REDSPOT", sql.Int, parseInt(req.body.REDSPOT));
        request.input("V_CODE", sql.Int, parseInt(req.body.V_CODE));
        request.input("RAPTYPE", sql.VarChar(5), req.body.RAPTYPE);
        request.input("DIA", sql.Int, parseInt(req.body.DIA));
        request.input("DEPTH", sql.Int, parseInt(req.body.DEPTH));
        request.input("RATIO", sql.Int, parseInt(req.body.RATIO));
        request.input("TAB", sql.Int, parseInt(req.body.TAB));
        request.input("RTYPE", sql.VarChar(5), req.body.RTYPE);
        request.input("COL_REM", sql.VarChar(2), req.body.COL_REM);
        request.input("QUA_REM", sql.VarChar(2), req.body.QUA_REM);
        request.input("ODIA", sql.Numeric(10, 2), req.body.ODIA);
        request.input("ORATIO", sql.Numeric(10, 2), req.body.ORATIO);
        request.input("ODEPTH", sql.Numeric(10, 2), req.body.ODEPTH);
        request.input("OTAB", sql.Numeric(10, 2), req.body.OTAB);
        request.input("ISMFG", sql.Bit, req.body.ISMFG);

        if (req.body.MPER) {
          request.input("ISMPER", sql.Bit, req.body.ISMPER);
        } else {
          request.input("ISMPER", sql.Bit, false);
        }
        if (req.body.MPER) {
          request.input("MPER", sql.Decimal(10, 2), parseFloat(req.body.MPER));
        } else {
          request.input("MPER", sql.Decimal(10, 2), parseFloat(0));
        }
        request.input("ISORD", sql.Bit, req.body.ISORD);

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

exports.FindRapType = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        // const pool = new sql.Request();
        let request = new sql.Request();

        request.input("S_CODE", sql.VarChar(5), req.body.S_CODE);
        request.input("Q_CODE", sql.Int, parseInt(req.body.Q_CODE));
        request.input("C_CODE", sql.Int, parseInt(req.body.C_CODE));
        request.input("CARAT", sql.Decimal(10, 3), parseFloat(req.body.CARAT));
        request.input("CUT_CODE", sql.Int, parseInt(req.body.CUT_CODE));
        request.input("POL_CODE", sql.Int, parseInt(req.body.POL_CODE));
        request.input("SYM_CODE", sql.Int, parseInt(req.body.SYM_CODE));
        request.input("FL_CODE", sql.Int, parseInt(req.body.FL_CODE));
        request.input("IN_CODE", sql.Int, parseInt(req.body.IN_CODE));
        request.input("SH_CODE", sql.Int, parseInt(req.body.SH_CODE));
        request.input("TABLE", sql.Int, parseInt(req.body.TABLE));
        request.input("TABLE_BLACK", sql.Int, parseInt(req.body.TABLE_BLACK));
        request.input("TABLE_OPEN", sql.Int, parseInt(req.body.TABLE_OPEN));
        request.input("SIDE", sql.Int, parseInt(req.body.SIDE));
        request.input("SIDE_BLACK", sql.Int, parseInt(req.body.SIDE_BLACK));
        request.input("SIDE_OPEN", sql.Int, parseInt(req.body.SIDE_OPEN));
        request.input("CROWN_OPEN", sql.Int, parseInt(req.body.CROWN_OPEN));
        request.input("GIRDLE_OPEN", sql.Int, parseInt(req.body.GIRDLE_OPEN));
        request.input("PAV_OPEN", sql.Int, parseInt(req.body.PAV_OPEN));
        request.input("CULET", sql.Int, parseInt(req.body.CULET));
        request.input("EXTFACET", sql.Int, parseInt(req.body.EXTFACET));
        request.input("EYECLAEN", sql.Int, parseInt(req.body.EYECLAEN));
        request.input("GRAINING", sql.Int, parseInt(req.body.GRAINING));
        request.input("LUSTER", sql.Int, parseInt(req.body.LUSTER));
        request.input("MILKY", sql.Int, parseInt(req.body.MILKY));
        request.input("NATURAL", sql.Int, parseInt(req.body.NATURAL));
        request.input("REDSPOT", sql.Int, parseInt(req.body.REDSPOT));
        request.input("V_CODE", sql.Int, parseInt(req.body.V_CODE));
        request.input("DIA", sql.Int, parseInt(req.body.DIA));
        request.input("DEPTH", sql.Int, parseInt(req.body.DEPTH));
        request.input("RATIO", sql.Int, parseInt(req.body.RATIO));
        request.input("TAB", sql.Int, parseInt(req.body.TAB));
        request.input("RAPTYPE", sql.VarChar(5), req.body.RAPTYPE);
        request.input("ISORD", sql.Bit, req.body.ISORD);
        request.input("RTYPE", sql.VarChar(5), req.body.RTYPE);
        request.input("COL_REM", sql.VarChar(2), req.body.COL_REM);
        request.input("QUA_REM", sql.VarChar(2), req.body.QUA_REM);
        request.input("ISMFG", sql.Bit, req.body.ISMFG);


        request = await request.execute("ufn_FindRapType");

        if (request.output) {
          res.json({ success: 1, data: request.output[""] });
        } else {
          res.json({ success: 0, data: "Not Found" });
        }
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.RapSave = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;
      let RapArr = req.body;
      // console.log('@####', RapArr[0])
      let status = true;
      let ErrorRap = [];
      for (let i = 0; i < RapArr.length; i++) {
        try {
          var request = new sql.Request();
          let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;

          request.input("L_CODE", sql.VarChar(500), RapArr[i].L_CODE);
          request.input("SRNO", sql.Int, parseInt(RapArr[i].SRNO));
          request.input("TAG", sql.VarChar(5), RapArr[i].TAG);
          request.input("PRDTYPE", sql.VarChar(5), RapArr[i].PRDTYPE);
          request.input("PLANNO", sql.Int, parseInt(RapArr[i].PLANNO));
          request.input("EMP_CODE", sql.VarChar(5), RapArr[i].EMP_CODE);
          request.input("PTAG", sql.VarChar(5), RapArr[i].PTAG);
          request.input(
            "I_CARAT",
            sql.Decimal(10, 3),
            parseFloat(RapArr[i].I_CARAT)
          );
          request.input("PRDS_CODE", sql.VarChar(5), RapArr[i].PRDS_CODE);
          request.input("PRDQ_CODE", sql.Int, parseInt(RapArr[i].PRDQ_CODE));
          request.input("PRDC_CODE", sql.Int, parseInt(RapArr[i].PRDC_CODE));
          request.input(
            "PRDCARAT",
            sql.Decimal(10, 3),
            parseFloat(RapArr[i].PRDCARAT)
          );
          request.input("PRDMC_CODE", sql.Int, parseInt(RapArr[i].PRDMC_CODE));
          request.input(
            "PRDCUT_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDCUT_CODE)
          );
          request.input(
            "PRDPOL_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDPOL_CODE)
          );
          request.input(
            "PRDSYM_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDSYM_CODE)
          );
          request.input("PRDFL_CODE", sql.Int, parseInt(RapArr[i].PRDFL_CODE));
          request.input(
            "PRDMFL_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDMFL_CODE)
          );
          request.input("PRDIN_CODE", sql.Int, parseInt(RapArr[i].PRDIN_CODE));
          request.input("PRDSH_CODE", sql.Int, parseInt(RapArr[i].PRDSH_CODE));
          request.input("PRDTABLE", sql.Int, parseInt(RapArr[i].PRDTABLE));
          request.input(
            "PRDTABLE_BLACK",
            sql.Int,
            parseInt(RapArr[i].PRDTABLE_BLACK)
          );
          request.input(
            "PRDTABLE_OPEN",
            sql.Int,
            parseInt(RapArr[i].PRDTABLE_OPEN)
          );
          request.input("PRDSIDE", sql.Int, parseInt(RapArr[i].PRDSIDE));
          request.input(
            "PRDSIDE_BLACK",
            sql.Int,
            parseInt(RapArr[i].PRDSIDE_BLACK)
          );
          request.input(
            "PRDSIDE_OPEN",
            sql.Int,
            parseInt(RapArr[i].PRDSIDE_OPEN)
          );
          request.input(
            "PRDCROWN_OPEN",
            sql.Int,
            parseInt(RapArr[i].PRDCROWN_OPEN)
          );
          request.input(
            "PRDGIRDLE_OPEN",
            sql.Int,
            parseInt(RapArr[i].PRDGIRDLE_OPEN)
          );
          request.input(
            "PRDPAV_OPEN",
            sql.Int,
            parseInt(RapArr[i].PRDPAV_OPEN)
          );
          request.input("PRDCULET", sql.Int, parseInt(RapArr[i].PRDCULET));
          request.input(
            "PRDEXTFACET",
            sql.Int,
            parseInt(RapArr[i].PRDEXTFACET)
          );
          request.input(
            "PRDEYECLEAN",
            sql.Int,
            parseInt(RapArr[i].PRDEYECLEAN)
          );
          request.input(
            "PRDGRAINING",
            sql.Int,
            parseInt(RapArr[i].PRDGRAINING)
          );
          request.input("PRDLUSTER", sql.Int, parseInt(RapArr[i].PRDLUSTER));
          request.input("PRDMILKY", sql.Int, parseInt(RapArr[i].PRDMILKY));
          request.input("PRDNATURAL", sql.Int, parseInt(RapArr[i].PRDNATURAL));
          request.input("PRDREDSPOT", sql.Int, parseInt(RapArr[i].PRDREDSPOT));
          request.input(
            "PRDDIA_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDDIA_CODE)
          );
          request.input(
            "PRDRAT_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDRAT_CODE)
          );
          request.input(
            "PRDDEPTH_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDDEPTH_CODE)
          );
          request.input(
            "PRDTAB_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDTAB_CODE)
          );
          request.input("RATE", sql.Int, parseInt(RapArr[i].RATE));
          request.input("CPRATE", sql.Int, parseInt(RapArr[i].CPRATE));
          request.input("CMRATE", sql.Int, parseInt(RapArr[i].CMRATE));
          request.input("QPRATE", sql.Int, parseInt(RapArr[i].QPRATE));
          request.input("QMRATE", sql.Int, parseInt(RapArr[i].QMRATE));
          request.input("PLNSEL", sql.Bit, RapArr[i].PLNSEL);
          request.input("IUSER", sql.VarChar(10), RapArr[i].IUSER);
          request.input("ICOMP", sql.VarChar(30), IP);
          request.input("TYP", sql.VarChar(5), RapArr[i].TYP);
          request.input("ISSEL", sql.Bit, RapArr[i].ISSEL);
          request.input("V_CODE", sql.Int, parseInt(RapArr[i].V_CODE));
          request.input("RAPTYPE", sql.VarChar(5), RapArr[i].RAPTYPE);
          request.input("R_CODE", sql.Int, parseInt(RapArr[i].R_CODE));
          request.input("OTAG", sql.VarChar(5), RapArr[i].OTAG);
          request.input("RELOCK", sql.Bit, RapArr[i].RELOCK);
          request.input("ORDNO", sql.Int, parseInt(RapArr[i].ORDNO));
          request.input("ISCOMMON", sql.Bit, RapArr[i].ISCOMMON);
          request.input('PRDDIA', sql.Numeric(10, 2), RapArr[i].ODIA);
          request.input('PRDDEPTH', sql.Numeric(10, 2), RapArr[i].ODEPTH);
          request.input('PRDRAT', sql.Numeric(10, 2), RapArr[i].ORATIO);
          request.input('PRDTAB', sql.Numeric(10, 2), RapArr[i].OTAB);
          request.input('PRDGIRDLE', sql.Numeric(10, 3), RapArr[i].PRDGIRDLE);
          request.input('MC1', sql.VarChar(50), RapArr[i].MC1);
          request.input('MC2', sql.VarChar(50), RapArr[i].MC2);
          request.input('MC3', sql.VarChar(50), RapArr[i].MC3);
          request.input('TENC_CODE', sql.Int, parseInt(RapArr[i].TENC_CODE));
          request.input('YAHUC_CODE', sql.Int, parseInt(RapArr[i].YAHUC_CODE));
          request.input('SPEC_CODE', sql.Int, parseInt(RapArr[i].SPEC_CODE));
          request.input('NDC_CODE', sql.Int, parseInt(RapArr[i].NDC_CODE));
          request.input('YAHUFL_CODE', sql.Int, parseInt(RapArr[i].YAHUFL_CODE));
          request.input('TENFL_CODE', sql.Int, parseInt(RapArr[i].TENFL_CODE));
          request.input('NDFL_CODE', sql.Int, parseInt(RapArr[i].NDFL_CODE));
          request.input('TENQ_CODE', sql.Int, parseInt(RapArr[i].TENQ_CODE));
          request.input('NDQ_CODE', sql.Int, parseInt(RapArr[i].NDQ_CODE));
          request.input('MBOXQ_CODE', sql.Int, parseInt(RapArr[i].MBOXQ_CODE));
          request.input('PRDWHINC', sql.Int, parseInt(RapArr[i].PRDWHINC));

          request = await request.execute("USP_RapCalPrdSave");

          // console.log("ddddddd",request.recordset)
        } catch (err) {
          console.log(err)
          status = false;
          ErrorRap.push(err);
        }
      }

      res.json({ success: status, data: ErrorRap });
    }
  });
};

exports.PrdTagUpd = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(5), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("TAG", sql.VarChar(2), req.body.TAG);

        request = await request.execute("USP_RapCalPrdTagUpd");

        res.json({ success: 1, data: "" });
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.RapCalcSaveValidation = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(5), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("TAG", sql.VarChar(5), req.body.TAG);
        request.input("EMP_CODE", sql.VarChar(5), req.body.EMP_CODE);
        request.input("PLANNO", sql.Int, parseInt(req.body.PLANNO));
        request.input("IUSER", sql.VarChar(10), req.body.IUSER);

        request = await request.execute("ufn_RapCalcSaveValidation");
        if (request.output) {
          res.json({ success: 1, data: request.output[""] });
        } else {
          res.json({ success: 0, data: "Not Found" });
        }
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.RapCalcDoubleClick = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401)
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('L_CODE', sql.VarChar(7999), req.body.L_CODE);
        request.input('SRNO', sql.Int, parseInt(req.body.SRNO));
        request.input('TAG', sql.VarChar(5), req.body.TAG);

        request = await request.execute('USP_RapCalcDoubleClick')
        // console.log(request)
        res.json({ success: 1, data: request.recordset })

      } catch (err) {
        console.log(err)
        res.json({ success: 0, data: err })
      }
    }
  })
}

exports.RapCalcSelectValidation = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(5), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("PTAG", sql.VarChar(5), req.body.PTAG);
        request.input("EMP_CODE", sql.VarChar(5), req.body.EMP_CODE);
        request.input("PLANNO", sql.Int, parseInt(req.body.PLANNO));
        request.input("IUSER", sql.VarChar(10), req.body.IUSER);

        request = await request.execute("uFn_RapCalcSelectValidation");
        if (request.output) {
          res.json({ success: 1, data: request.output[""] });
        } else {
          res.json({ success: 0, data: "Not Found" });
        }
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.RapPrint = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(10), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("TAG", sql.VarChar(5), req.body.TAG);
        request.input("EMP_CODE", sql.VarChar(5), req.body.EMP_CODE);

        request = await request.execute("RP_RapPrint");

        if (request.recordsets) {
          res.json({ success: 1, data: request.recordsets });
        } else {
          res.json({ success: 0, data: "Not Found" });
        }
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.RapPrintOP = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(10), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("TAG", sql.VarChar(5), req.body.TAG);
        request.input("EMP_CODE", sql.VarChar(5), req.body.EMP_CODE);

        request = await request.execute("RP_RapPrintOP");

        if (request.recordsets) {
          res.json({ success: 1, data: request.recordsets });
        } else {
          res.json({ success: 0, data: "Not Found" });
        }
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.RapCalPrdDel = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        let IP = req.headers["x-forwarded-for"] || req.connection.remoteAddress;

        request.input("L_CODE", sql.VarChar(10), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("TAG", sql.VarChar(5), req.body.TAG);
        request.input("OTAG", sql.VarChar(5), req.body.OTAG);
        request.input("PRDTYPE", sql.VarChar(5), req.body.PRDTYPE);
        request.input("PLANNO", sql.Int, parseInt(req.body.PLANNO));
        request.input("EMP_CODE", sql.VarChar(10), req.body.EMP_CODE);
        request.input("PTAG", sql.VarChar(5), req.body.PTAG);
        request.input("IUSER", sql.VarChar(10), req.body.IUSER);
        request.input("PROC_CODE", sql.Int, parseInt(req.body.PROC_CODE));
        request.input("IP", sql.VarChar(25), IP);

        request = await request.execute("USP_RapCalPrdDel");

        res.json({ success: 1, data: "" });
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.TendarRapCalDisp = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(50), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("TAG", sql.VarChar(2), req.body.TAG);

        request = await request.execute("USP_TendarRapCalDisp");

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

exports.TendarRapCalPrdSave = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;
      let RapArr = req.body;
      let status = true;
      let ErrorRap = [];
      for (let i = 0; i < RapArr.length; i++) {
        try {
          var request = new sql.Request();
          let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
          
          request.input("L_CODE", sql.VarChar(50), RapArr[i].L_CODE);
          request.input("SRNO", sql.Int, parseInt(RapArr[i].SRNO));
          request.input("TAG", sql.VarChar(5), RapArr[i].TAG);
          request.input("PRDTYPE", sql.VarChar(5), RapArr[i].PRDTYPE);
          request.input("PLANNO", sql.Int, parseInt(RapArr[i].PLANNO));
          request.input("EMP_CODE", sql.VarChar(5), RapArr[i].EMP_CODE);
          request.input("PTAG", sql.VarChar(5), RapArr[i].PTAG);
          request.input("ADNO", sql.Int, parseInt(RapArr[i].ADNO));
          request.input(
            "I_CARAT",
            sql.Decimal(10, 3),
            parseFloat(RapArr[i].I_CARAT)
          );
          request.input("DIA", sql.Decimal(10, 3), parseFloat(RapArr[i].DIA));
          request.input("PRDS_CODE", sql.VarChar(5), RapArr[i].PRDS_CODE);
          request.input("PRDQ_CODE", sql.Int, parseInt(RapArr[i].PRDQ_CODE));
          request.input("PRDC_CODE", sql.Int, parseInt(RapArr[i].PRDC_CODE));
          request.input("I_PCS", sql.Int, parseInt(RapArr[i].I_PCS));
          request.input('R_CARAT', sql.Numeric(10, 3), RapArr[i].R_CARAT)
          request.input(
            "PRDCARAT",
            sql.Decimal(10, 3),
            parseFloat(RapArr[i].PRDCARAT)
          );
          request.input("PRDMC_CODE", sql.Int, parseInt(RapArr[i].PRDMC_CODE));
          request.input(
            "PRDCUT_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDCUT_CODE)
          );
          request.input(
            "PRDPOL_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDPOL_CODE)
          );
          request.input(
            "PRDSYM_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDSYM_CODE)
          );
          request.input("PRDFL_CODE", sql.Int, parseInt(RapArr[i].PRDFL_CODE));
          request.input(
            "PRDMFL_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDMFL_CODE)
          );
          request.input("PRDIN_CODE", sql.Int, parseInt(RapArr[i].PRDIN_CODE));
          request.input("PRDSH_CODE", sql.Int, parseInt(RapArr[i].PRDSH_CODE));
          request.input("PRDTABLE", sql.Int, parseInt(RapArr[i].PRDTABLE));
          request.input(
            "PRDTABLE_BLACK",
            sql.Int,
            parseInt(RapArr[i].PRDTABLE_BLACK)
          );
          request.input(
            "PRDTABLE_OPEN",
            sql.Int,
            parseInt(RapArr[i].PRDTABLE_OPEN)
          );
          request.input("PRDSIDE", sql.Int, parseInt(RapArr[i].PRDSIDE));
          request.input(
            "PRDSIDE_BLACK",
            sql.Int,
            parseInt(RapArr[i].PRDSIDE_BLACK)
          );
          request.input(
            "PRDSIDE_OPEN",
            sql.Int,
            parseInt(RapArr[i].PRDSIDE_OPEN)
          );
          request.input(
            "PRDCROWN_OPEN",
            sql.Int,
            parseInt(RapArr[i].PRDCROWN_OPEN)
          );
          request.input(
            "PRDGIRDLE_OPEN",
            sql.Int,
            parseInt(RapArr[i].PRDGIRDLE_OPEN)
          );
          request.input(
            "PRDPAV_OPEN",
            sql.Int,
            parseInt(RapArr[i].PRDPAV_OPEN)
          );
          request.input("PRDCULET", sql.Int, parseInt(RapArr[i].PRDCULET));
          request.input(
            "PRDEXTFACET",
            sql.Int,
            parseInt(RapArr[i].PRDEXTFACET)
          );
          request.input(
            "PRDEYECLEAN",
            sql.Int,
            parseInt(RapArr[i].PRDEYECLEAN)
          );
          request.input(
            "PRDGRAINING",
            sql.Int,
            parseInt(RapArr[i].PRDGRAINING)
          );
          request.input("PRDLUSTER", sql.Int, parseInt(RapArr[i].PRDLUSTER));
          request.input("PRDMILKY", sql.Int, parseInt(RapArr[i].PRDMILKY));
          request.input("PRDNATURAL", sql.Int, parseInt(RapArr[i].PRDNATURAL));
          request.input("PRDREDSPOT", sql.Int, parseInt(RapArr[i].PRDREDSPOT));
          request.input(
            "PRDDIA_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDDIA_CODE)
          );
          request.input(
            "PRDRAT_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDRAT_CODE)
          );
          request.input(
            "PRDDEPTH_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDDEPTH_CODE)
          );
          request.input(
            "PRDTAB_CODE",
            sql.Int,
            parseInt(RapArr[i].PRDTAB_CODE)
          );
          request.input("RATE", sql.Int, parseInt(RapArr[i].RATE));
          request.input("CPRATE", sql.Int, parseInt(RapArr[i].CPRATE));
          request.input("CMRATE", sql.Int, parseInt(RapArr[i].CMRATE));
          request.input("QPRATE", sql.Int, parseInt(RapArr[i].QPRATE));
          request.input("QMRATE", sql.Int, parseInt(RapArr[i].QMRATE));
          request.input("PLNSEL", sql.Bit, RapArr[i].PLNSEL);
          request.input("FPLNSEL", sql.Bit, RapArr[i].FPLNSEL);
          request.input("IUSER", sql.VarChar(10), RapArr[i].IUSER);
          request.input("ICOMP", sql.VarChar(30), IP);
          request.input("TYP", sql.VarChar(5), RapArr[i].TYP);
          request.input("ISSEL", sql.Bit, RapArr[i].ISSEL);
          request.input("RAPTYPE", sql.VarChar(5), RapArr[i].RAPTYPE);
          request.input("V_CODE", sql.Int, parseInt(RapArr[i].V_CODE));
          request.input("DEP", sql.Decimal(10, 3), parseFloat(RapArr[i].DEP));
          request.input("RAT", sql.Decimal(10, 3), parseFloat(RapArr[i].RAT));
          request.input("COMENT", sql.VarChar(100), RapArr[i].COMENT);
          request.input("MC1", sql.VarChar(50), RapArr[i].MC1);
          request.input("MC2", sql.VarChar(50), RapArr[i].MC2);
          request.input("MC3", sql.VarChar(50), RapArr[i].MC3);
          request.input("DN", sql.VarChar(50), RapArr[i].DN);
          request.input("E1", sql.VarChar(50), RapArr[i].E1);
          request.input("E2", sql.VarChar(50), RapArr[i].E2);
          request.input("E3", sql.VarChar(50), RapArr[i].E3);
          request.input("T_CODE", sql.Int, parseInt(RapArr[i].T_CODE));
          request.input("L_NAME", sql.VarChar(50), RapArr[i].L_NAME);
          request.input("COL_REM", sql.VarChar(2), RapArr[i].COL_REM);
          request.input("QUA_REM", sql.VarChar(2), RapArr[i].QUA_REM);
          request.input('ISIGI', sql.Bit, RapArr[i].ISIGI)

          // console.log(req.body)
          request = await request.execute("USP_TendarRapCalPrdSave");
        } catch (err) {
          status = false;
          ErrorRap.push(err.message);
        }
      }

      res.json({ success: status, data: ErrorRap });
    }
  });
};

exports.TendarRapCalPrdDelete = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(50), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("TAG", sql.VarChar(2), req.body.TAG);

        request = await request.execute("USP_TendarRapCalPrdDelete");

        res.json({ success: 1, data: "" });
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.TendarExcel = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(50), req.body.L_CODE);
        request.input("FSRNO", sql.Int, parseInt(req.body.FSRNO));
        request.input("TSRNO", sql.Int, parseInt(req.body.TSRNO));
        if (req.body.F_DATE != "") {
          request.input("F_DATE", sql.DateTime2, new Date(req.body.F_DATE));
        }
        if (req.body.T_DATE != "") {
          request.input("T_DATE", sql.DateTime2, new Date(req.body.T_DATE));
        }
        request.input("TYPE", sql.VarChar(10), req.body.TYPE);
        request.input("F_CARAT", sql.Numeric(10, 3), req.body.F_CARAT);
        request.input("T_CARAT", sql.Numeric(10, 3), req.body.T_CARAT);

        // console.log(req.body)
        request = await request.execute("USP_TendarExcel");

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

exports.TendarExcelDownload = async (req, res) => {
  try {
    var request = new sql.Request();

    request.input("L_CODE", sql.VarChar(50), req.body.L_CODE);
    if (req.body.FSRNO) {
      request.input("FSRNO", sql.Int, parseInt(req.body.FSRNO));
    } else {
      request.input("FSRNO", sql.Int, 0);
    }
    if (req.body.TSRNO) {
      request.input("TSRNO", sql.Int, parseInt(req.body.TSRNO));
    } else {
      request.input("TSRNO", sql.Int, 0);
    }
    if (req.body.F_DATE) {
      request.input("F_DATE", sql.DateTime2, new Date(req.body.F_DATE));
    } else {
      request.input("F_DATE", sql.DateTime2, null);
    }
    if (req.body.T_DATE) {
      request.input("T_DATE", sql.DateTime2, new Date(req.body.T_DATE));
    } else {
      request.input("T_DATE", sql.DateTime2, null);
    }
    request.input("TYPE", sql.VarChar(10), "EXCEL");

    request = await request.execute("USP_TendarExcel");
    // console.log('BODY: ', req.body);
    // console.log('LENGTH: ', request.recordsets.length, 'RECORDSET: ', request.recordsets[13]);
    // console.log('LENGTH: ', request.recordsets.length, 'RECORDSET: ', request.recordsets[0]);

    const Excel = require("exceljs");
    const workbook = new Excel.Workbook();

    // console.log(JSON.parse(req.body.SETTING))
    const Setting = JSON.parse(req.body.SETTING);
    const TenderExcle = [];
    const TenderColor = [];

    Setting.map((it) => {
      if (it.SKEY == "TENDAR_EXCEL") {
        TenderExcle.push(it);
      } else if (it.SKEY == "TENDAR_COLOR") {
        TenderColor.push(it);
      }
    });

    // console.log("TenderExcel", TenderExcle);
    // console.log("TenderColor", TenderColor);

    const TENCOLSVALUE = TenderColor[0].SVALUE.split(",");
    const TENDEXSVALUE = TenderExcle[0].SVALUE.split(",");

    let colorcheck = TENCOLSVALUE;

    // console.log("col", TENCOLSVALUE);
    // console.log("EX", TENDEXSVALUE);

    //* Sheet 1
    const worksheet1 = workbook.addWorksheet("Summary");
    const colorCell1 = ["A", "B", "C", "D", "E", "F"];
    worksheet1.columns = [
      { key: "LOT" },
      { key: "SRNO" },
      { key: "R.WEIGHT" },
      { key: "AVG" },
      { key: "DN(-+)" },
      { key: "BID" },
    ];

    worksheet1.getRow(1).values = [
      "LOT",
      "SRNO",
      "R.WEIGHT",
      "R.CTS",
      "DN(-+)",
      "BID",
    ];
    for (let i = 0; i < colorCell1.length; i++) {
      worksheet1.getCell(colorCell1[i] + "1").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "cca366" },
      };
    }
    if (request.recordsets[13].length != 0) {
      request.recordsets[13].forEach(function (record, index) {
        // ID1 = index + 1;
        ID2 = index + 2;
        let Bid =
          "IF(E" +
          ID2 +
          '= "", "", D' +
          ID2 +
          "+ ((D" +
          ID2 +
          "* E" +
          ID2 +
          ") / 100))";
        let dd = `=VLOOKUP(B${ID2},Details!AN2:AO${request.recordsets[0].length + 1
          },2,TRUE)`;
        let avg = `=VLOOKUP(B${ID2},Details!AN2:AP${request.recordsets[0].length + 1
          },3,TRUE)`;

        worksheet1.addRow({
          LOT: record["LOT"],
          SRNO: record["SRNO"],
          "R.WEIGHT": record["R.WEIGHT"] ? record["R.WEIGHT"].toFixed(2) : "",
          AVG: { formula: avg },
          "DN(-+)": { formula: dd },
          BID: { formula: Bid },
        });
      });
    } else {
      console.log('Data Not Found')
      // alert("Data Not Found");
    }

    worksheet1.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      row.eachCell(function (cell, colNumber) {
        cell.alignment = {
          vertical: "middle",
          horizontal: "center",
        };

        for (var i = 1; i < 7; i++) {
          row.getCell(i).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }
      });
    });

    //* Sheet 2
    const worksheet2 = workbook.addWorksheet("Details");
    const colorCell = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W",
      "X",
      "Y",
      "Z",
      "AA",
      "AB",
      "AC",
      "AD",
      "AE",
      "AF",
      "AG",
      "AG",
      "AH",
      "AI",
      "AJ",
      "AK",
      "AL",
      "AM",
      "AN",
      "AO",
      "AP",
      "AQ",
      "AR",
      "AS",
      "AT",
      "AU",
      "AV",
      "AW",
      "AX",
      "AY",
      "AZ",
      "BA",
      "BB",
      "BC",
      "BD",
      "BE",
    ];

    // Main table data
    worksheet2.columns = [
      { key: "LOT", width: "13.5" },
      { key: "SRNO" },
      { key: "IUSER" },
      { key: "PLANNO" },
      { key: "R.WEIGHT" },
      { key: "SHAPE" },
      { key: "CUT" },
      { key: "SIZE" },
      { key: "COLOR" },
      { key: "CLARITY" },
      { key: "FLO" },
      { key: "INC" },
      { key: "SHADE" },
      { key: "MILKY" },
      { key: "RAP" },
      { key: "DIS" },
      { key: "MUM" },
      { key: "SRT" },
      { key: "P.CTS" },
      { key: "AMOUNT" },
      { key: "TOTAL" },
      { key: "AVG" },
      { key: "R.P(%)" },
      { key: "DEP" },
      { key: "RAT" },
      { key: "DIAMETER" },
      { key: "NEWPLN" },
      { key: "COMENT" },
      { key: "MC1" },
      { key: "MC2" },
      { key: "MC3" },
      { key: "DN" },
      { key: "E1" },
      { key: "E2" },
      { key: "E3" },
      { key: "TENSION" },
      { key: "DN(-+)" },
      { key: "BID" },
      { key: "LINK" },
      { key: "NSRNO" },
      { key: "DD" },
      { key: "NRCTS" },
      { key: 'IPCS' },
      { key: 'R_CARAT' },
      { key: 'MCOL' },
      { key: 'TYP' },
      { key: 'IGICut' },
      { key: 'IGICol' },
      { key: 'IGIQua' },
      { key: 'IGIFlo' },
      { key: 'IGISHADE' },
      { key: 'IGIMILKY' },
      { key: 'IGIWHINC' },
      { key: 'IGIORAP' },
      { key: 'IGIDIS' },
      { key: 'IGIRATE' },
      { key: 'IGIAMT' },
    ];


    worksheet2.getRow(1).values = [
      "Tender Name",
      "SRNO",
      "IUSER",
      "PLANNO",
      "R.WEIGHT",
      "SHAPE",
      "CUT",
      "SIZE",
      "COLOR",
      "CLARITY",
      "FLO",
      "INC",
      "SHADE",
      "MILKY",
      "RAP",
      "DIS",
      "MUM",
      "SRT",
      "P.CTS",
      "AMOUNT",
      "TOTAL",
      "R.CTS",
      "R.P(%)",
      "DEPTH",
      "RATIO",
      "DIAMETER",
      "NEWPLN",
      "COMENT",
      "MC1",
      "MC2",
      "MC3",
      "DN",
      "E1",
      "E2",
      "E3",
      "TENSION",
      "DN(-+)",
      "BID",
      "LINK",
      "NSRNO",
      "DD",
      "NRCTS",
      "IPCS",
      "R.Crt",
      "M.COL",
      "TYP",
      "IGICUT",
      "IGICOL",
      "IGIQUA",
      "IGIFLO",
      "IGISHADE",
      "IGIMILKY",
      "IGIWHINC",
      "IGIORAP",
      "IGIDIS",
      "IGIRATE",
      "IGIAMT",  
    ];

    for (let i = 0; i < colorCell.length; i++) {
      worksheet2.getCell(colorCell[i] + "1").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "cca366" },
      };
    }

    let unqSrNoArray = [
      ...new Set(
        request.recordsets[0]
          .map((item) => {
            return item.SRNO;
          })
          .filter((fItem) => fItem)
      ),
    ];

    let srnoStartIndex = unqSrNoArray.map((item) => {
      return {
        SRNO: item,
        SRNO_INDEX:
          request.recordsets[0].findIndex((value) => {
            return value.SRNO === item;
          }) + 2,
      };
    });

    // let a = worksheet2.getRow(2)
    // let CountMatchTotal = 0
    // let PreviousTot = ''
    // let TempTotCountArr = []
    request.recordsets[0].forEach(function (record, index) {
      ID1 = index + 1;
      ID2 = index + 2;

      let PerCrt =
        "IF(P" + ID2 + '=0,"",O' + ID2 + "-(O" + ID2 + "*(P" + ID2 + "/100)))";
      // let PerCrt = '=IFERROR((O' + ID2 + '-(O' + ID2 + '*(P' + ID2 + '/100))), "")'
      let Amt = "IF(P" + ID2 + '=0,"",FLOOR(S' + ID2 + "*H" + ID2 + "*AQ" + ID2 + ",1))";
      // =IF(P2 = 0, "", Q2 * G2)
      // let Amt = '=IFERROR((O' + ID2 + '-(O' + ID2 + '*(P' + ID2 + '/100))), "")'
      let Tot =
        "IF(IF(U" +
        ID2 +
        "=U" +
        ID1 +
        ',"",SUMIF(U:U,U' +
        ID2 +
        ',R:R))=0,"",IF(U' +
        ID2 +
        "=U" +
        ID1 +
        ',"",SUMIF(U:U,U' +
        ID2 +
        ",R:R)))";
      // let Tot = 'SUM(R:R)'
      let Avg =
        "FLOOR(U" +
        ID2 +
        "/E" +
        (record["SRNO"]
          ? srnoStartIndex.find((item) => item.SRNO === record["SRNO"])
            .SRNO_INDEX
          : ID2) +
        ",1)";
      // let Bid = IF(AC5 = "", "", T5 - ((T5 * AC5) / 100))
      let Bid =
        "IF(AK" +
        ID2 +
        '= "", "", V' +
        ID2 +
        "+((V" +
        ID2 +
        "* AK" +
        ID2 +
        ") / 100))";

      let AvgIGI = `=IFERROR(BB${ID2}-(AA${ID2}*BC${ID2}/100), "")`
      let mlIGI = `=IFERROR((BD${ID2}*H${ID2}),"")`
      // let Bid = 'T' + ID2 + '-((T' + ID2 + '*AC' + ID2 + ')/100)'
      // console.log(Tot);
      // let ct = record['CARAT'] ? record['CARAT'].toFixed(2) : 0;
      // let rp = record['GRAP'] ? record['GRAP'].toFixed(2) : 0;
      // let rt = record['GRATE'] ? record['GRATE'].toFixed(2) : 0;
      // let tct = totalCts.toFixed(2);

      //DB na data for loop fare 6e
      worksheet2.addRow({
        // NO: index + 1,
        LOT: record["LOT"],
        SRNO: record["SRNO"],
        PLANNO: record["PLANNO"],
        "R.WEIGHT": record["R.WEIGHT"] ? record["R.WEIGHT"] : "",
        SHAPE: record["SHAP"],
        CUT: record["CUT"],
        SIZE: record["SIZE"] ? record["SIZE"] : "",
        COLOR: record["COLOR"],
        CLARITY: record["CLARITY"],
        FLO: record["FLO"],
        INC: record["INC"],
        SHADE: record["SHADE"],
        MILKY: record["MILKY"],
        RAP: record["RAP"],
        DIS: record["DIS"] ? Math.floor(parseFloat(record["DIS"])) : "",
        MUM: record["MUM"],
        SRT: record["SRT"],
        // 'P.CTS': parseFloat(record['P.CTS']),
        "P.CTS": record["LOT"] ? { formula: PerCrt } : "",
        // AMOUNT: parseFloat(record['AMOUNT']),
        AMOUNT: record["LOT"] ? { formula: Amt } : "",
        // TOTAL: parseFloat(record['TOTAL']),
        // TOTAL: record['LOT'] ? { formula: Tot } : '',
        TOTAL: "",
        // AVG: record['LOT'] ? { formula: Avg } : '',
        AVG: record["LOT"] ? { formula: Avg } : "",
        IUSER: record["IUSER"],
        DIAMETER: record["DIA"],
        DEP: record["DEP"] ? parseFloat(record["DEP"]) : "",
        RAT: record["RAT"] ? parseFloat(record["RAT"]) : "",
        "R.P(%)": record["R.P(%)"] ? parseFloat(record["R.P(%)"]) : "",
        NEWPLN: record["NEWPLN"],
        COMENT: record["COMENT"],
        MC1: record["MC1"],
        MC2: record["MC2"],
        MC3: record["MC3"],
        DN: record["DN"],
        E1: record["E1"],
        E2: record["E2"],
        E3: record["E3"],
        TENSION: record["TENSION"],
        "DN(-+)": record["DN(-+)"],
        BID: record["LOT"] ? { formula: Bid } : "",
        // LINK: record['LINK'],
        NSRNO: record["NSRNO"],
        IPCS: record["IPCS"],
        R_CARAT: record["R_CARAT"],
        MCOL: record["MCOL"],
        TYP: record["TYP"],
        IGICol: record["IGIC_NAME"],
        IGIQua: record["IGIQ_NAME"],
        IGICut: record["IGICUT"],
        IGIFlo: record["IGIFL_NAME"],
        IGIORAP: record["IGIORAP"],
        IGIDIS: record["IGIDIS"],
        IGIRATE: { formula: AvgIGI},
        IGIAMT: { formula: mlIGI},
        IGISHADE: record["IGISHADE"],
        IGIMILKY: record["IGIMILKY"],
        IGIWHINC: record["WHINC"]
        // NRCTS: { formula: nrcts},
      });

      if (record["LINK"]) {
        worksheet2.getCell("AM" + ID2).value = {
          text: "Video",
          hyperlink: record["LINK"] ? record["LINK"] : "",
        };
        worksheet2.getCell("AM" + ID2).font = {
          color: { argb: "1d4ed8" },
          underline: true,
        };
      } else {
        worksheet2.getCell("AM" + ID2).value = "";
      }
      const indexForColor = index;

      colorcheck = TENCOLSVALUE.map((it) => {
        return {
          code: it,
          color: false,
        };
      });
      // console.log("colorche",worksheet2.columns)

      // Color cell as per condition
      if (record["SHAP"] != "ROUND" && record["SHAP"] != "") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        TENCOLSVALUE.map((t) => {
          if ("SHAPE" == t) {
            Temp = true;
          }
        });
        // for(let i = 0; i < TENCOLSVALUE.length; i++){
        //   if(TENCOLSVALUE[i] == "SHAP" ){

        //     Temp = true;
        //   }
        // }
        // })
        if (Temp) {
          worksheet2.getCell("F" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffff00" },
          };
        } else {
          worksheet2.getCell("F" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
      }

      if (parseFloat(record["R.P(%)"]) > 42) {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("R.P(%)" == t) {
        //     Temp = true;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if (TENCOLSVALUE[i] == "R.P(%)") {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("W" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF7D7D" },
          };
        } else {
          worksheet2.getCell("W" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
      }

      if (record["CUT"] == "EX") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("CUT" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if (TENCOLSVALUE[i] == "CUT") {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("G" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF" },
          };
        } else {
          worksheet2.getCell("G" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
      } else if (record["CUT"] == "VG") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("CUT" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // TENCOLSVALUE.map((t) => {
        //   if ("CUT" == t) {
        //   } else {
        //     Temp = false;
        //   }
        // });
        // })
        if (Temp) {
          worksheet2.getCell("G" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "92D050" },
          };
        } else {
          worksheet2.getCell("G" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('G' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: '92D050' },
        // };
      } else if (record["CUT"] == "GD" || record["CUT"] == "FR") {
        let Temp = false;
        // TENCOLSVALUE.map((t) => {
        //   if ("CUT" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("CUT" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        if (Temp) {
          worksheet2.getCell("G" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF0000" },
          };
        } else {
          worksheet2.getCell("G" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('G' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'FF0000' },
        // };
      }

      if (record["FLO"] == "NO") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("FLO" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("FLO" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "dce6f1" },
          };
        } else {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('K' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'dce6f1' },
        // };
      } else if (record["FLO"] == "FA") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("FLO" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("FLO" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        } else {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('K' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'c5d9f1' },
        // };
      } else if (record["FLO"] == "ME") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("FLO" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("FLO" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "95b3d7" },
          };
        } else {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('K' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: '95b3d7' },
        // };
      } else if (record["FLO"] == "ST") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("FLO" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("FLO" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "538dd5" },
          };
        } else {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('K' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: '538dd5' },
        // };
      } else if (record["FLO"] == "VE") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("FLO" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("FLO" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "366092" },
          };
        } else {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('K' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: '366092' },
        // };
      }

      if (record["SHADE"] == "VL-BRN") {
        let Temp = false;
        // TENCOLSVALUE.map((t) => {
        //   if ("SHADE" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("SHADE" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        if (Temp) {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "f2dcdb" },
          };
        } else {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('M' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'f2dcdb' },
        // };
      } else if (record["SHADE"] == "L-BRN") {
        let Temp = false;
        // TENCOLSVALUE.map((t) => {
        //   if ("SHADE" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("SHADE" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        if (Temp) {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "da9694" },
          };
        } else {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('M' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'da9694' },
        // };
      } else if (record["SHADE"] == "D-BRN") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("SHADE" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("SHADE" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "963634" },
          };
        } else {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('M' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: '963634' },
        // };
      } else if (record["SHADE"] == "MIX-T") {
        let Temp = false;
        // TENCOLSVALUE.map((t) => {
        //   if ("SHADE" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("SHADE" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        if (Temp) {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "e26b0a" },
          };
        } else {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('M' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'e26b0a' },
        // };
      }

      if (record["MILKY"] == "L-MILKY") {
        let Temp = false;
        // TENCOLSVALUE.map((t) => {
        //   if ("MILKY" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("MILKY" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        if (Temp) {
          worksheet2.getCell("N" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c4bd97" },
          };
        } else {
          worksheet2.getCell("N" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('N' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'c4bd97' },
        // };
      } else if (record["MILKY"] == "H-MILKY") {
        let Temp = false;
        // TENCOLSVALUE.map((t) => {
        //   if ("MILKY" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("MILKY" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        if (Temp) {
          worksheet2.getCell("N" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "948a54" },
          };
        } else {
          worksheet2.getCell("N" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('N' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: '948a54' },
        // };
      }

      // console.log(record['NEWPLN'])
      // if (record['NEWPLN'] == '') {
      //     TempTotCountArr.push(CountMatchTotal)
      //     PreviousTot = ''
      //     CountMatchTotal = 0
      // }
      // if (PreviousTot != record['NEWPLN'] && record['NEWPLN'] != '') {
      //     PreviousTot = record['NEWPLN']
      //     CountMatchTotal++
      // } else {
      //     if (PreviousTot == record['NEWPLN']) {
      //         CountMatchTotal++
      //     }
      // }
    });
    // console.log('CountMatchTotal: ', CountMatchTotal)
    // console.log('PreviousTot: ', PreviousTot)
    // console.log('TempTotCountArr: ', TempTotCountArr)
    // let totalMergeIndex = 2;
    // for (const i in TempTotCountArr) {
    //     if (i == 0) {
    //         let START = totalMergeIndex;
    //         let END = totalMergeIndex + (TempTotCountArr[i] - 1);
    //         totalMergeIndex = totalMergeIndex + TempTotCountArr[i]
    //         worksheet2.mergeCells('S' + START + ':S' + END);
    //         worksheet2.mergeCells('T' + START + ':T' + END);
    //         worksheet2.mergeCells('AC' + START + ':AC' + END);
    //         worksheet2.mergeCells('AD' + START + ':AD' + END);
    //         // console.log(START, END)
    //     } else {
    //         let START = totalMergeIndex + 1;
    //         let END = totalMergeIndex + (TempTotCountArr[i] - 1);
    //         totalMergeIndex = totalMergeIndex + TempTotCountArr[i]
    //         worksheet2.mergeCells('S' + START + ':S' + END);
    //         worksheet2.mergeCells('T' + START + ':T' + END);
    //         worksheet2.mergeCells('AC' + START + ':AC' + END);
    //         worksheet2.mergeCells('AD' + START + ':AD' + END);
    //         // console.log(START, END)
    //     }
    // }

    prevValue = "";
    repeatPlanNo = 0;
    totalArray = [];

    const result = request.recordsets[0].map((item, index) => {
      if (item.NEWPLN == "") item.NEWPLN = null;
      // console.log(prevValue, item.NEWPLN)
      if (index !== 0) {
        if (item.NEWPLN != prevValue) {
          totalArray.push(repeatPlanNo);
          repeatPlanNo = 0;
        } else {
          repeatPlanNo++;
        }
      }
      prevValue = item.NEWPLN;
    });

    let mergeIndex = 2;
    for (const i in totalArray) {
      let start = parseInt(mergeIndex);
      let end = parseInt(mergeIndex + totalArray[i]);
      mergeIndex = mergeIndex + totalArray[i] + 1;
      worksheet2.mergeCells("U" + start + ":U" + end);
      worksheet2.mergeCells("V" + start + ":V" + end);
      worksheet2.mergeCells("AK" + start + ":AK" + end);
      worksheet2.mergeCells("AL" + start + ":AL" + end);
      let totalFormula = "SUM(T" + start + ":T" + end + ")";
      worksheet2.getCell("U" + start).value = { formula: totalFormula };
    }

    // console.log(worksheet2.columns.map((t) => console.log(t)));
    // console.log(worksheet2.getCell('F4'))

    prevValue = "";
    totalArray = [];
    repeatSrNo = 0;
    request.recordsets[0].map((item, index) => {
      if (item.SRNO == "") item.SRNO = null;
      // console.log(prevValue, item.NEWPLN)
      if (index !== 0) {
        if (item.SRNO != prevValue) {
          totalArray.push(repeatSrNo);
          repeatSrNo = 0;
        } else {
          repeatSrNo++;
        }
      }
      prevValue = item.SRNO;
    });

    mergeIndex = 2;
    for (const i in totalArray) {
      let start = parseInt(mergeIndex);
      let end = parseInt(mergeIndex + totalArray[i]);
      mergeIndex = mergeIndex + totalArray[i] + 1;
      worksheet2.mergeCells("E" + start + ":E" + end);
      worksheet2.mergeCells("AN" + start + ":AN" + end);
      worksheet2.mergeCells("AO" + start + ":AO" + end);
      worksheet2.mergeCells("AP" + start + ":AP" + end);
      let ddformula =
        "=IF(AK" + start + '="","",' + "MAX(AK" + start + ":AK" + end + "))";
      let maxvalue =
        "IF(V" + start + '="","",' + "MAX(V" + start + ":V" + end + "))";
      worksheet2.getCell("AO" + start).value = { formula: ddformula };
      worksheet2.getCell("AP" + start).value = { formula: maxvalue };
    }

    worksheet2.getColumn(27).hidden = true;
    worksheet2.getColumn(40).hidden = true;
    worksheet2.getColumn(41).hidden = true;
    worksheet2.getColumn(42).hidden = true;

    // function getColumnIndex(columnName) {
    //     var columnIndex = 0;
    //     var chars = columnName.split('');
    //     for (var i = 0; i < chars.length; i++) {
    //       columnIndex = columnIndex * 26 + (chars[i].charCodeAt(0) - 'A'.charCodeAt(0) + 1);
    //     }
    //     return columnIndex;
    //   }

    worksheet2.columns.map((it, c) => {
      TENDEXSVALUE.map((i) => {
        if (it._key == i) {
          // console.log(it._number)
          worksheet2.getColumn(it._number).hidden = true;
        }
      });
    });

    worksheet2.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      row.eachCell(function (cell, colNumber) {
        cell.alignment = {
          vertical: "middle",
          horizontal: "center",
        };

        for (var i = 1; i < 58; i++) {
          row.getCell(i).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }
      });
    });

    worksheet2.getColumn(7).numFmt = "0.00";

    //total table data
    let totalTableIndex = request.recordsets[0].length + 4;
    // console.log(totalTableIndex)
    // background color and fill border
    for (let i = 0; i < 9; i++) {
      worksheet2.getCell(colorCell[8 + i] + totalTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "cca366" },
      };
      worksheet2.getCell(colorCell[8 + i] + totalTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(
        colorCell[8 + i] + (totalTableIndex + 1).toString()
      ).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }

    worksheet2.addTable({
      name: "tableTotal",
      ref: "I" + totalTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "LOT" },
        { name: "RoughCrt" },
        { name: "RoughPcs" },
        { name: "PolishCrt" },
        { name: "Amtount" },
        { name: "AvgORap" },
        { name: "P.CTS" },
        { name: "AvgDis%" },
        { name: "R.CTS" },
      ],
      rows: request.recordsets[1].map((item) => Object.values(item)),
    });
    let TotalRow = request.recordsets[2].map((item) => Object.values(item));

    //shape table data
    let shapeTableIndex = totalTableIndex + request.recordsets[1].length + 3;
    // console.log('shapeTableIndex', shapeTableIndex);
    let shapeTableMergeIndex = shapeTableIndex - 1;
    let shapeRowData = request.recordsets[3].map((item) => Object.values(item));
    shapeRowData.push(TotalRow[0]);
    worksheet2.mergeCells(
      "B" + shapeTableMergeIndex + ":F" + shapeTableMergeIndex
    );
    worksheet2.getCell("B" + shapeTableMergeIndex).value = "Shape Proposal";
    worksheet2.getCell("B" + shapeTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("B" + shapeTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("B" + (shapeTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("B" + (shapeTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[1 + i] + shapeTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(colorCell[1 + i] + shapeTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[3].length + 1; j++) {
        worksheet2.getCell(
          colorCell[1 + i] + (shapeTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[3].length + 1) {
          worksheet2.getCell(
            colorCell[1 + i] + (shapeTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }

    worksheet2.addTable({
      name: "shapeProposal",
      ref: "B" + shapeTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Shape Detail" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: shapeRowData,
    });

    //size table data
    let sizeTableIndex = shapeTableIndex + request.recordsets[3].length + 5;
    let sizeTableMergeIndex = sizeTableIndex - 1;
    let sizeRowData = request.recordsets[4].map((item) => Object.values(item));
    sizeRowData.push(TotalRow[0]);
    worksheet2.mergeCells(
      "B" + sizeTableMergeIndex + ":F" + sizeTableMergeIndex
    );
    worksheet2.getCell("B" + sizeTableMergeIndex).value = "Size Proposal";
    worksheet2.getCell("B" + sizeTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("B" + sizeTableMergeIndex).font = {
      bold: true,
    };
    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("B" + (sizeTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("B" + (sizeTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[1 + i] + sizeTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(colorCell[1 + i] + sizeTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[4].length + 1; j++) {
        worksheet2.getCell(
          colorCell[1 + i] + (sizeTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[4].length + 1) {
          worksheet2.getCell(
            colorCell[1 + i] + (sizeTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }

    worksheet2.addTable({
      name: "sizeProposal",
      ref: "B" + sizeTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Size" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: sizeRowData,
    });

    //color table data
    let colorTableIndex = totalTableIndex + request.recordsets[1].length + 3;
    let colorTableMergeIndex = colorTableIndex - 1;
    let colorRowData = request.recordsets[5].map((item) => Object.values(item));
    colorRowData.push(TotalRow[0]);
    worksheet2.mergeCells(
      "H" + colorTableMergeIndex + ":L" + colorTableMergeIndex
    );
    worksheet2.getCell("H" + colorTableMergeIndex).value = "Color Proposal";
    worksheet2.getCell("H" + colorTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("H" + colorTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("H" + (colorTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("H" + (colorTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[7 + i] + colorTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(colorCell[7 + i] + colorTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[5].length + 1; j++) {
        worksheet2.getCell(
          colorCell[7 + i] + (colorTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[5].length + 1) {
          worksheet2.getCell(
            colorCell[7 + i] + (colorTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "colorProposal",
      ref: "H" + colorTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Color" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: colorRowData,
    });

    //quality table data
    let qualityTableIndex = colorTableIndex + request.recordsets[5].length + 5;
    let qualityTableMergeIndex = qualityTableIndex - 1;
    let qualityRowData = request.recordsets[6].map((item) =>
      Object.values(item)
    );
    qualityRowData.push(TotalRow[0]);
    worksheet2.mergeCells(
      "H" + qualityTableMergeIndex + ":L" + qualityTableMergeIndex
    );
    worksheet2.getCell("H" + qualityTableMergeIndex).value = "Quality Proposal";
    worksheet2.getCell("H" + qualityTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("H" + qualityTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("H" + (qualityTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("H" + (qualityTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[7 + i] + qualityTableIndex.toString()).fill =
      {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(
        colorCell[7 + i] + qualityTableIndex.toString()
      ).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[6].length + 1; j++) {
        worksheet2.getCell(
          colorCell[7 + i] + (qualityTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[6].length + 1) {
          worksheet2.getCell(
            colorCell[7 + i] + (qualityTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "qualityProposal",
      ref: "H" + qualityTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Quality" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: qualityRowData,
    });

    //cut table data
    let cutTableIndex = totalTableIndex + request.recordsets[1].length + 3;
    let cutTableMergeIndex = cutTableIndex - 1;
    let cutRowData = request.recordsets[8].map((item) => Object.values(item));
    cutRowData.push(TotalRow[0]);
    worksheet2.mergeCells("N" + cutTableMergeIndex + ":R" + cutTableMergeIndex);
    worksheet2.getCell("N" + cutTableMergeIndex).value = "Cut Proposal";
    worksheet2.getCell("N" + cutTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("N" + cutTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("N" + (cutTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("N" + (cutTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[13 + i] + cutTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(colorCell[13 + i] + cutTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[8].length + 1; j++) {
        worksheet2.getCell(
          colorCell[13 + i] + (cutTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[8].length + 1) {
          worksheet2.getCell(
            colorCell[13 + i] + (cutTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "cutProposal",
      ref: "N" + cutTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Cut" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: cutRowData,
    });

    //polish table data
    let polishTableIndex = cutTableIndex + request.recordsets[8].length + 5;
    let polishTableMergeIndex = polishTableIndex - 1;
    let polishRowData = request.recordsets[9].map((item) =>
      Object.values(item)
    );
    polishRowData.push(TotalRow[0]);
    worksheet2.mergeCells(
      "N" + polishTableMergeIndex + ":R" + polishTableMergeIndex
    );
    worksheet2.getCell("N" + polishTableMergeIndex).value = "Polish Proposal";
    worksheet2.getCell("N" + polishTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("N" + polishTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("N" + (polishTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("N" + (polishTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[13 + i] + polishTableIndex.toString()).fill =
      {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(
        colorCell[13 + i] + polishTableIndex.toString()
      ).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[8].length + 1; j++) {
        worksheet2.getCell(
          colorCell[13 + i] + (polishTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[8].length + 1) {
          worksheet2.getCell(
            colorCell[13 + i] + (polishTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "PolishProposal",
      ref: "N" + polishTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Polish" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: polishRowData,
    });

    //sym table data
    let symTableIndex = polishTableIndex + request.recordsets[9].length + 5;
    let symTableMergeIndex = symTableIndex - 1;
    let symRowData = request.recordsets[10].map((item) => Object.values(item));
    symRowData.push(TotalRow[0]);
    worksheet2.mergeCells("N" + symTableMergeIndex + ":R" + symTableMergeIndex);
    worksheet2.getCell("N" + symTableMergeIndex).value = "Symmetry Proposal";
    worksheet2.getCell("N" + symTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("N" + symTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("N" + (symTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("N" + (symTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[13 + i] + symTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(colorCell[13 + i] + symTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[10].length + 1; j++) {
        worksheet2.getCell(
          colorCell[13 + i] + (symTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[10].length + 1) {
          worksheet2.getCell(
            colorCell[13 + i] + (symTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "symProposal",
      ref: "N" + symTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Symmetry" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: symRowData,
    });

    //floro table data
    let floroTableIndex = symTableIndex + request.recordsets[10].length + 5;
    let floroTableMergeIndex = floroTableIndex - 1;
    let floroRowData = request.recordsets[11].map((item) =>
      Object.values(item)
    );
    floroRowData.push(TotalRow[0]);
    worksheet2.mergeCells(
      "N" + floroTableMergeIndex + ":R" + floroTableMergeIndex
    );
    worksheet2.getCell("N" + floroTableMergeIndex).value = "Floro Proposal";
    worksheet2.getCell("N" + floroTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("N" + floroTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("N" + (floroTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("N" + (floroTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[13 + i] + floroTableIndex.toString()).fill =
      {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(
        colorCell[13 + i] + floroTableIndex.toString()
      ).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[11].length + 1; j++) {
        worksheet2.getCell(
          colorCell[13 + i] + (floroTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[11].length + 1) {
          worksheet2.getCell(
            colorCell[13 + i] + (floroTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "floroProposal",
      ref: "N" + floroTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Floro" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: floroRowData,
    });

    //inc table data
    let incTableIndex = floroTableIndex + request.recordsets[11].length + 5;
    let incTableMergeIndex = incTableIndex - 1;
    let incRowData = request.recordsets[12].map((item) => Object.values(item));
    incRowData.push(TotalRow[0]);
    worksheet2.mergeCells("N" + incTableMergeIndex + ":R" + incTableMergeIndex);
    worksheet2.getCell("N" + incTableMergeIndex).value = "Inc Proposal";
    worksheet2.getCell("N" + incTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("N" + incTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("N" + (incTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("N" + (incTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[13 + i] + incTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(colorCell[13 + i] + incTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[12].length + 1; j++) {
        worksheet2.getCell(
          colorCell[13 + i] + (incTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[12].length + 1) {
          worksheet2.getCell(
            colorCell[13 + i] + (incTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "incProposal",
      ref: "N" + incTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Inc" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: incRowData,
    });

    workbook.xlsx.writeBuffer().then(function (buffer) {
      let xlsData = Buffer.concat([buffer]);
      res
        .writeHead(200, {
          "Content-Length": Buffer.byteLength(xlsData),
          "Content-Type": "application/vnd.ms-excel",
          "Content-disposition": "attachment;filename=TENDAR.xlsx",
        })
        .end(xlsData);
    });
  } catch (err) {
    console.log(err);
    res.json({ success: 0, data: err });
  }
};
exports.TendarExcelDownloadSKD = async (req, res) => {
  try {
    var request = new sql.Request();

    request.input("L_CODE", sql.VarChar(50), req.body.L_CODE);
    if (req.body.FSRNO) {
      request.input("FSRNO", sql.Int, parseInt(req.body.FSRNO));
    } else {
      request.input("FSRNO", sql.Int, 0);
    }
    if (req.body.TSRNO) {
      request.input("TSRNO", sql.Int, parseInt(req.body.TSRNO));
    } else {
      request.input("TSRNO", sql.Int, 0);
    }
    if (req.body.F_DATE) {
      request.input("F_DATE", sql.DateTime2, new Date(req.body.F_DATE));
    } else {
      request.input("F_DATE", sql.DateTime2, null);
    }
    if (req.body.T_DATE) {
      request.input("T_DATE", sql.DateTime2, new Date(req.body.T_DATE));
    } else {
      request.input("T_DATE", sql.DateTime2, null);
    }
    request.input("TYPE", sql.VarChar(10), "EXCEL");

    request = await request.execute("USP_TendarExcel");
    // console.log('BODY: ', req.body);
    // console.log('LENGTH: ', request.recordsets.length, 'RECORDSET: ', request.recordsets[13]);
    // console.log('LENGTH: ', request.recordsets.length, 'RECORDSET: ', request.recordsets[0]);

    const Excel = require("exceljs");
    const workbook = new Excel.Workbook();

    // console.log(JSON.parse(req.body.SETTING))
    const Setting = JSON.parse(req.body.SETTING);
    const TenderExcle = [];
    const TenderColor = [];

    Setting.map((it) => {
      if (it.SKEY == "TENDAR_EXCEL") {
        TenderExcle.push(it);
      } else if (it.SKEY == "TENDAR_COLOR") {
        TenderColor.push(it);
      }
    });

    // console.log("TenderExcel", TenderExcle);
    // console.log("TenderColor", TenderColor);

    const TENCOLSVALUE = TenderColor[0].SVALUE.split(",");
    const TENDEXSVALUE = TenderExcle[0].SVALUE.split(",");

    let colorcheck = TENCOLSVALUE;

    // console.log("col", TENCOLSVALUE);
    // console.log("EX", TENDEXSVALUE);

    //* Sheet 1
    const worksheet1 = workbook.addWorksheet("Summary");
    const colorCell1 = ["A", "B", "C", "D", "E", "F"];
    worksheet1.columns = [
      { key: "LOT" },
      { key: "SRNO" },
      { key: "R.WEIGHT" },
      { key: "AVG" },
      { key: "DN(-+)" },
      { key: "BID" },
    ];

    worksheet1.getRow(1).values = [
      "LOT",
      "SRNO",
      "R.WEIGHT",
      "R.CTS",
      "DN(-+)",
      "BID",
    ];
    for (let i = 0; i < colorCell1.length; i++) {
      worksheet1.getCell(colorCell1[i] + "1").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "cca366" },
      };
    }
    if (request.recordsets[13].length != 0) {
      request.recordsets[13].forEach(function (record, index) {
        // ID1 = index + 1;
        ID2 = index + 2;
        let Bid =
          "IF(E" +
          ID2 +
          '= "", "", D' +
          ID2 +
          "+ ((D" +
          ID2 +
          "* E" +
          ID2 +
          ") / 100))";
        let dd = `=VLOOKUP(B${ID2},Details!AN2:AO${request.recordsets[0].length + 1
          },2,TRUE)`;
        let avg = `=VLOOKUP(B${ID2},Details!AN2:AP${request.recordsets[0].length + 1
          },3,TRUE)`;

        worksheet1.addRow({
          LOT: record["LOT"],
          SRNO: record["SRNO"],
          "R.WEIGHT": record["R.WEIGHT"] ? record["R.WEIGHT"].toFixed(2) : "",
          AVG: { formula: avg },
          "DN(-+)": { formula: dd },
          BID: { formula: Bid },
        });
      });
    } else {
      console.log('Data Not Found')
      // alert("Data Not Found");
    }

    worksheet1.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      row.eachCell(function (cell, colNumber) {
        cell.alignment = {
          vertical: "middle",
          horizontal: "center",
        };

        for (var i = 1; i < 7; i++) {
          row.getCell(i).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }
      });
    });

    //* Sheet 2
    const worksheet2 = workbook.addWorksheet("Details");
    const colorCell = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W",
      "X",
      "Y",
      "Z",
      "AA",
      "AB",
      "AC",
      "AD",
      "AE",
      "AF",
      "AG",
      "AG",
      "AH",
      "AI",
      "AJ",
      "AK",
      "AL",
      "AM",
      "AN",
      "AO",
      "AP",
      "AQ",
      "AR",
      "AS",
      "AT",
      // "AU",
      // "AV",
      // "AW",
      // "AX",
      // "AY",
      // "AZ",
      // "BA",
      // "BB",
      // "BC",
      // "BD",
      // "BE",
    ];

    // Main table data
    worksheet2.columns = [
      { key: "LOT", width: "13.5" },
      { key: "SRNO" },
      { key: "IUSER" },
      { key: "PLANNO" },
      { key: "R.WEIGHT" },
      { key: "SHAPE" },
      { key: "CUT" },
      { key: "SIZE" },
      { key: "COLOR" },
      { key: "CLARITY" },
      { key: "FLO" },
      { key: "INC" },
      { key: "SHADE" },
      { key: "MILKY" },
      { key: "RAP" },
      { key: "DIS" },
      { key: "MUM" },
      { key: "SRT" },
      { key: "P.CTS" },
      { key: "AMOUNT" },
      { key: "TOTAL" },
      { key: "AVG" },
      { key: "R.P(%)" },
      { key: "DEP" },
      { key: "RAT" },
      { key: "DIAMETER" },
      { key: "NEWPLN" },
      { key: "COMENT" },
      { key: "MC1" },
      { key: "MC2" },
      { key: "MC3" },
      { key: "DN" },
      { key: "E1" },
      { key: "E2" },
      { key: "E3" },
      { key: "TENSION" },
      { key: "DN(-+)" },
      { key: "BID" },
      { key: "LINK" },
      { key: "NSRNO" },
      { key: "DD" },
      { key: "NRCTS" },
      { key: 'IPCS' },
      { key: 'R_CARAT' },
      { key: 'MCOL' },
      { key: 'TYP' },
      // { key: 'IGICut' },
      // { key: 'IGICol' },
      // { key: 'IGIQua' },
      // { key: 'IGIFlo' },
      // { key: 'IGISHADE' },
      // { key: 'IGIMILKY' },
      // { key: 'IGIWHINC' },
      // { key: 'IGIORAP' },
      // { key: 'IGIDIS' },
      // { key: 'IGIRATE' },
      // { key: 'IGIAMT' },
    ];


    worksheet2.getRow(1).values = [
      "Tender Name",
      "SRNO",
      "IUSER",
      "PLANNO",
      "R.WEIGHT",
      "SHAPE",
      "CUT",
      "SIZE",
      "COLOR",
      "CLARITY",
      "FLO",
      "INC",
      "SHADE",
      "MILKY",
      "RAP",
      "DIS",
      "MUM",
      "SRT",
      "P.CTS",
      "AMOUNT",
      "TOTAL",
      "R.CTS",
      "R.P(%)",
      "DEPTH",
      "RATIO",
      "DIAMETER",
      "NEWPLN",
      "COMENT",
      "MC1",
      "MC2",
      "MC3",
      "DN",
      "E1",
      "E2",
      "E3",
      "TENSION",
      "DN(-+)",
      "BID",
      "LINK",
      "NSRNO",
      "DD",
      "NRCTS",
      "IPCS",
      "R.Crt",
      "M.COL",
      "TYP",
      // "IGICUT",
      // "IGICOL",
      // "IGIQUA",
      // "IGIFLO",
      // "IGISHADE",
      // "IGIMILKY",
      // "IGIWHINC",
      // "IGIORAP",
      // "IGIDIS",
      // "IGIRATE",
      // "IGIAMT",  
    ];

    for (let i = 0; i < colorCell.length; i++) {
      worksheet2.getCell(colorCell[i] + "1").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "cca366" },
      };
    }

    let unqSrNoArray = [
      ...new Set(
        request.recordsets[0]
          .map((item) => {
            return item.SRNO;
          })
          .filter((fItem) => fItem)
      ),
    ];

    let srnoStartIndex = unqSrNoArray.map((item) => {
      return {
        SRNO: item,
        SRNO_INDEX:
          request.recordsets[0].findIndex((value) => {
            return value.SRNO === item;
          }) + 2,
      };
    });

    // let a = worksheet2.getRow(2)
    // let CountMatchTotal = 0
    // let PreviousTot = ''
    // let TempTotCountArr = []
    request.recordsets[0].forEach(function (record, index) {
      ID1 = index + 1;
      ID2 = index + 2;

      let PerCrt =
        "IF(P" + ID2 + '=0,"",O' + ID2 + "-(O" + ID2 + "*(P" + ID2 + "/100)))";
      // let PerCrt = '=IFERROR((O' + ID2 + '-(O' + ID2 + '*(P' + ID2 + '/100))), "")'
      let Amt = "IF(P" + ID2 + '=0,"",FLOOR(S' + ID2 + "*H" + ID2 + "*AQ" + ID2 + ",1))";
      // =IF(P2 = 0, "", Q2 * G2)
      // let Amt = '=IFERROR((O' + ID2 + '-(O' + ID2 + '*(P' + ID2 + '/100))), "")'
      let Tot =
        "IF(IF(U" +
        ID2 +
        "=U" +
        ID1 +
        ',"",SUMIF(U:U,U' +
        ID2 +
        ',R:R))=0,"",IF(U' +
        ID2 +
        "=U" +
        ID1 +
        ',"",SUMIF(U:U,U' +
        ID2 +
        ",R:R)))";
      // let Tot = 'SUM(R:R)'
      let Avg =
        "FLOOR(U" +
        ID2 +
        "/E" +
        (record["SRNO"]
          ? srnoStartIndex.find((item) => item.SRNO === record["SRNO"])
            .SRNO_INDEX
          : ID2) +
        ",1)";
      // let Bid = IF(AC5 = "", "", T5 - ((T5 * AC5) / 100))
      let Bid =
        "IF(AK" +
        ID2 +
        '= "", "", V' +
        ID2 +
        "+((V" +
        ID2 +
        "* AK" +
        ID2 +
        ") / 100))";

      let AvgIGI = `=IFERROR(BB${ID2}-(AA${ID2}*BC${ID2}/100), "")`
      let mlIGI = `=IFERROR((BD${ID2}*H${ID2}),"")`
      // let Bid = 'T' + ID2 + '-((T' + ID2 + '*AC' + ID2 + ')/100)'
      // console.log(Tot);
      // let ct = record['CARAT'] ? record['CARAT'].toFixed(2) : 0;
      // let rp = record['GRAP'] ? record['GRAP'].toFixed(2) : 0;
      // let rt = record['GRATE'] ? record['GRATE'].toFixed(2) : 0;
      // let tct = totalCts.toFixed(2);

      //DB na data for loop fare 6e
      worksheet2.addRow({
        // NO: index + 1,
        LOT: record["LOT"],
        SRNO: record["SRNO"],
        PLANNO: record["PLANNO"],
        "R.WEIGHT": record["R.WEIGHT"] ? record["R.WEIGHT"] : "",
        SHAPE: record["SHAP"],
        CUT: record["CUT"],
        SIZE: record["SIZE"] ? record["SIZE"] : "",
        COLOR: record["COLOR"],
        CLARITY: record["CLARITY"],
        FLO: record["FLO"],
        INC: record["INC"],
        SHADE: record["SHADE"],
        MILKY: record["MILKY"],
        RAP: record["RAP"],
        DIS: record["DIS"] ? Math.floor(parseFloat(record["DIS"])) : "",
        MUM: record["MUM"],
        SRT: record["SRT"],
        // 'P.CTS': parseFloat(record['P.CTS']),
        "P.CTS": record["LOT"] ? { formula: PerCrt } : "",
        // AMOUNT: parseFloat(record['AMOUNT']),
        AMOUNT: record["LOT"] ? { formula: Amt } : "",
        // TOTAL: parseFloat(record['TOTAL']),
        // TOTAL: record['LOT'] ? { formula: Tot } : '',
        TOTAL: "",
        // AVG: record['LOT'] ? { formula: Avg } : '',
        AVG: record["LOT"] ? { formula: Avg } : "",
        IUSER: record["IUSER"],
        DIAMETER: record["DIA"],
        DEP: record["DEP"] ? parseFloat(record["DEP"]) : "",
        RAT: record["RAT"] ? parseFloat(record["RAT"]) : "",
        "R.P(%)": record["R.P(%)"] ? parseFloat(record["R.P(%)"]) : "",
        NEWPLN: record["NEWPLN"],
        COMENT: record["COMENT"],
        MC1: record["MC1"],
        MC2: record["MC2"],
        MC3: record["MC3"],
        DN: record["DN"],
        E1: record["E1"],
        E2: record["E2"],
        E3: record["E3"],
        TENSION: record["TENSION"],
        "DN(-+)": record["DN(-+)"],
        BID: record["LOT"] ? { formula: Bid } : "",
        // LINK: record['LINK'],
        NSRNO: record["NSRNO"],
        IPCS: record["IPCS"],
        R_CARAT: record["R_CARAT"],
        MCOL: record["MCOL"],
        TYP: record["TYP"],
        // IGICol: record["IGIC_NAME"],
        // IGIQua: record["IGIQ_NAME"],
        // IGICut: record["IGICUT"],
        // IGIFlo: record["IGIFL_NAME"],
        // IGIORAP: record["IGIORAP"],
        // IGIDIS: record["IGIDIS"],
        // IGIRATE: { formula: AvgIGI},
        // IGIAMT: { formula: mlIGI},
        // IGISHADE: record["IGISHADE"],
        // IGIMILKY: record["IGIMILKY"],
        // IGIWHINC: record["WHINC"]
        // NRCTS: { formula: nrcts},
      });

      if (record["LINK"]) {
        worksheet2.getCell("AM" + ID2).value = {
          text: "Video",
          hyperlink: record["LINK"] ? record["LINK"] : "",
        };
        worksheet2.getCell("AM" + ID2).font = {
          color: { argb: "1d4ed8" },
          underline: true,
        };
      } else {
        worksheet2.getCell("AM" + ID2).value = "";
      }
      const indexForColor = index;

      colorcheck = TENCOLSVALUE.map((it) => {
        return {
          code: it,
          color: false,
        };
      });
      // console.log("colorche",worksheet2.columns)

      // Color cell as per condition
      if (record["SHAP"] != "ROUND" && record["SHAP"] != "") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        TENCOLSVALUE.map((t) => {
          if ("SHAPE" == t) {
            Temp = true;
          }
        });
        // for(let i = 0; i < TENCOLSVALUE.length; i++){
        //   if(TENCOLSVALUE[i] == "SHAP" ){

        //     Temp = true;
        //   }
        // }
        // })
        if (Temp) {
          worksheet2.getCell("F" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffff00" },
          };
        } else {
          worksheet2.getCell("F" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
      }

      if (parseFloat(record["R.P(%)"]) > 42) {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("R.P(%)" == t) {
        //     Temp = true;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if (TENCOLSVALUE[i] == "R.P(%)") {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("W" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF7D7D" },
          };
        } else {
          worksheet2.getCell("W" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
      }

      if (record["CUT"] == "EX") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("CUT" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if (TENCOLSVALUE[i] == "CUT") {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("G" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF" },
          };
        } else {
          worksheet2.getCell("G" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
      } else if (record["CUT"] == "VG") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("CUT" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // TENCOLSVALUE.map((t) => {
        //   if ("CUT" == t) {
        //   } else {
        //     Temp = false;
        //   }
        // });
        // })
        if (Temp) {
          worksheet2.getCell("G" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "92D050" },
          };
        } else {
          worksheet2.getCell("G" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('G' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: '92D050' },
        // };
      } else if (record["CUT"] == "GD" || record["CUT"] == "FR") {
        let Temp = false;
        // TENCOLSVALUE.map((t) => {
        //   if ("CUT" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("CUT" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        if (Temp) {
          worksheet2.getCell("G" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF0000" },
          };
        } else {
          worksheet2.getCell("G" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('G' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'FF0000' },
        // };
      }

      if (record["FLO"] == "NO") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("FLO" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("FLO" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "dce6f1" },
          };
        } else {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('K' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'dce6f1' },
        // };
      } else if (record["FLO"] == "FA") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("FLO" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("FLO" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        } else {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('K' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'c5d9f1' },
        // };
      } else if (record["FLO"] == "ME") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("FLO" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("FLO" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "95b3d7" },
          };
        } else {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('K' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: '95b3d7' },
        // };
      } else if (record["FLO"] == "ST") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("FLO" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("FLO" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "538dd5" },
          };
        } else {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('K' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: '538dd5' },
        // };
      } else if (record["FLO"] == "VE") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("FLO" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("FLO" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "366092" },
          };
        } else {
          worksheet2.getCell("K" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('K' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: '366092' },
        // };
      }

      if (record["SHADE"] == "VL-BRN") {
        let Temp = false;
        // TENCOLSVALUE.map((t) => {
        //   if ("SHADE" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("SHADE" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        if (Temp) {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "f2dcdb" },
          };
        } else {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('M' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'f2dcdb' },
        // };
      } else if (record["SHADE"] == "L-BRN") {
        let Temp = false;
        // TENCOLSVALUE.map((t) => {
        //   if ("SHADE" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("SHADE" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        if (Temp) {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "da9694" },
          };
        } else {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('M' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'da9694' },
        // };
      } else if (record["SHADE"] == "D-BRN") {
        let Temp = false;
        // worksheet2.columns.map((it) => {
        // TENCOLSVALUE.map((t) => {
        //   if ("SHADE" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("SHADE" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        // })
        if (Temp) {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "963634" },
          };
        } else {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('M' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: '963634' },
        // };
      } else if (record["SHADE"] == "MIX-T") {
        let Temp = false;
        // TENCOLSVALUE.map((t) => {
        //   if ("SHADE" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("SHADE" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        if (Temp) {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "e26b0a" },
          };
        } else {
          worksheet2.getCell("M" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('M' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'e26b0a' },
        // };
      }

      if (record["MILKY"] == "L-MILKY") {
        let Temp = false;
        // TENCOLSVALUE.map((t) => {
        //   if ("MILKY" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("MILKY" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        if (Temp) {
          worksheet2.getCell("N" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c4bd97" },
          };
        } else {
          worksheet2.getCell("N" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('N' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: 'c4bd97' },
        // };
      } else if (record["MILKY"] == "H-MILKY") {
        let Temp = false;
        // TENCOLSVALUE.map((t) => {
        //   if ("MILKY" == t) {
        //     Temp = true;
        //   } else {
        //     Temp = false;
        //   }
        // });
        for (let i = 0; i < TENCOLSVALUE.length; i++) {
          if ("MILKY" == TENCOLSVALUE[i]) {

            Temp = true;
          }
        }
        if (Temp) {
          worksheet2.getCell("N" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "948a54" },
          };
        } else {
          worksheet2.getCell("N" + (indexForColor + 2).toString()).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffff" },
          };
        }
        // worksheet2.getCell('N' + (indexForColor + 2).toString()).fill = {
        //     type: 'pattern',
        //     pattern: 'solid',
        //     fgColor: { argb: '948a54' },
        // };
      }

      // console.log(record['NEWPLN'])
      // if (record['NEWPLN'] == '') {
      //     TempTotCountArr.push(CountMatchTotal)
      //     PreviousTot = ''
      //     CountMatchTotal = 0
      // }
      // if (PreviousTot != record['NEWPLN'] && record['NEWPLN'] != '') {
      //     PreviousTot = record['NEWPLN']
      //     CountMatchTotal++
      // } else {
      //     if (PreviousTot == record['NEWPLN']) {
      //         CountMatchTotal++
      //     }
      // }
    });
    // console.log('CountMatchTotal: ', CountMatchTotal)
    // console.log('PreviousTot: ', PreviousTot)
    // console.log('TempTotCountArr: ', TempTotCountArr)
    // let totalMergeIndex = 2;
    // for (const i in TempTotCountArr) {
    //     if (i == 0) {
    //         let START = totalMergeIndex;
    //         let END = totalMergeIndex + (TempTotCountArr[i] - 1);
    //         totalMergeIndex = totalMergeIndex + TempTotCountArr[i]
    //         worksheet2.mergeCells('S' + START + ':S' + END);
    //         worksheet2.mergeCells('T' + START + ':T' + END);
    //         worksheet2.mergeCells('AC' + START + ':AC' + END);
    //         worksheet2.mergeCells('AD' + START + ':AD' + END);
    //         // console.log(START, END)
    //     } else {
    //         let START = totalMergeIndex + 1;
    //         let END = totalMergeIndex + (TempTotCountArr[i] - 1);
    //         totalMergeIndex = totalMergeIndex + TempTotCountArr[i]
    //         worksheet2.mergeCells('S' + START + ':S' + END);
    //         worksheet2.mergeCells('T' + START + ':T' + END);
    //         worksheet2.mergeCells('AC' + START + ':AC' + END);
    //         worksheet2.mergeCells('AD' + START + ':AD' + END);
    //         // console.log(START, END)
    //     }
    // }

    prevValue = "";
    repeatPlanNo = 0;
    totalArray = [];

    const result = request.recordsets[0].map((item, index) => {
      if (item.NEWPLN == "") item.NEWPLN = null;
      // console.log(prevValue, item.NEWPLN)
      if (index !== 0) {
        if (item.NEWPLN != prevValue) {
          totalArray.push(repeatPlanNo);
          repeatPlanNo = 0;
        } else {
          repeatPlanNo++;
        }
      }
      prevValue = item.NEWPLN;
    });

    let mergeIndex = 2;
    for (const i in totalArray) {
      let start = parseInt(mergeIndex);
      let end = parseInt(mergeIndex + totalArray[i]);
      mergeIndex = mergeIndex + totalArray[i] + 1;
      worksheet2.mergeCells("U" + start + ":U" + end);
      worksheet2.mergeCells("V" + start + ":V" + end);
      worksheet2.mergeCells("AK" + start + ":AK" + end);
      worksheet2.mergeCells("AL" + start + ":AL" + end);
      let totalFormula = "SUM(T" + start + ":T" + end + ")";
      worksheet2.getCell("U" + start).value = { formula: totalFormula };
    }

    // console.log(worksheet2.columns.map((t) => console.log(t)));
    // console.log(worksheet2.getCell('F4'))

    prevValue = "";
    totalArray = [];
    repeatSrNo = 0;
    request.recordsets[0].map((item, index) => {
      if (item.SRNO == "") item.SRNO = null;
      // console.log(prevValue, item.NEWPLN)
      if (index !== 0) {
        if (item.SRNO != prevValue) {
          totalArray.push(repeatSrNo);
          repeatSrNo = 0;
        } else {
          repeatSrNo++;
        }
      }
      prevValue = item.SRNO;
    });

    mergeIndex = 2;
    for (const i in totalArray) {
      let start = parseInt(mergeIndex);
      let end = parseInt(mergeIndex + totalArray[i]);
      mergeIndex = mergeIndex + totalArray[i] + 1;
      worksheet2.mergeCells("E" + start + ":E" + end);
      worksheet2.mergeCells("AN" + start + ":AN" + end);
      worksheet2.mergeCells("AO" + start + ":AO" + end);
      worksheet2.mergeCells("AP" + start + ":AP" + end);
      let ddformula =
        "=IF(AK" + start + '="","",' + "MAX(AK" + start + ":AK" + end + "))";
      let maxvalue =
        "IF(V" + start + '="","",' + "MAX(V" + start + ":V" + end + "))";
      worksheet2.getCell("AO" + start).value = { formula: ddformula };
      worksheet2.getCell("AP" + start).value = { formula: maxvalue };
    }

    worksheet2.getColumn(27).hidden = true;
    worksheet2.getColumn(40).hidden = true;
    worksheet2.getColumn(41).hidden = true;
    worksheet2.getColumn(42).hidden = true;

    // function getColumnIndex(columnName) {
    //     var columnIndex = 0;
    //     var chars = columnName.split('');
    //     for (var i = 0; i < chars.length; i++) {
    //       columnIndex = columnIndex * 26 + (chars[i].charCodeAt(0) - 'A'.charCodeAt(0) + 1);
    //     }
    //     return columnIndex;
    //   }

    worksheet2.columns.map((it, c) => {
      TENDEXSVALUE.map((i) => {
        if (it._key == i) {
          // console.log(it._number)
          worksheet2.getColumn(it._number).hidden = true;
        }
      });
    });

    worksheet2.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      row.eachCell(function (cell, colNumber) {
        cell.alignment = {
          vertical: "middle",
          horizontal: "center",
        };

        for (var i = 1; i < 47; i++) {
          row.getCell(i).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }
      });
    });

    worksheet2.getColumn(7).numFmt = "0.00";

    //total table data
    let totalTableIndex = request.recordsets[0].length + 4;
    // console.log(totalTableIndex)
    // background color and fill border
    for (let i = 0; i < 9; i++) {
      worksheet2.getCell(colorCell[8 + i] + totalTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "cca366" },
      };
      worksheet2.getCell(colorCell[8 + i] + totalTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(
        colorCell[8 + i] + (totalTableIndex + 1).toString()
      ).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }

    worksheet2.addTable({
      name: "tableTotal",
      ref: "I" + totalTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "LOT" },
        { name: "RoughCrt" },
        { name: "RoughPcs" },
        { name: "PolishCrt" },
        { name: "Amtount" },
        { name: "AvgORap" },
        { name: "P.CTS" },
        { name: "AvgDis%" },
        { name: "R.CTS" },
      ],
      rows: request.recordsets[1].map((item) => Object.values(item)),
    });
    let TotalRow = request.recordsets[2].map((item) => Object.values(item));

    //shape table data
    let shapeTableIndex = totalTableIndex + request.recordsets[1].length + 3;
    // console.log('shapeTableIndex', shapeTableIndex);
    let shapeTableMergeIndex = shapeTableIndex - 1;
    let shapeRowData = request.recordsets[3].map((item) => Object.values(item));
    shapeRowData.push(TotalRow[0]);
    worksheet2.mergeCells(
      "B" + shapeTableMergeIndex + ":F" + shapeTableMergeIndex
    );
    worksheet2.getCell("B" + shapeTableMergeIndex).value = "Shape Proposal";
    worksheet2.getCell("B" + shapeTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("B" + shapeTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("B" + (shapeTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("B" + (shapeTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[1 + i] + shapeTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(colorCell[1 + i] + shapeTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[3].length + 1; j++) {
        worksheet2.getCell(
          colorCell[1 + i] + (shapeTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[3].length + 1) {
          worksheet2.getCell(
            colorCell[1 + i] + (shapeTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }

    worksheet2.addTable({
      name: "shapeProposal",
      ref: "B" + shapeTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Shape Detail" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: shapeRowData,
    });

    //size table data
    let sizeTableIndex = shapeTableIndex + request.recordsets[3].length + 5;
    let sizeTableMergeIndex = sizeTableIndex - 1;
    let sizeRowData = request.recordsets[4].map((item) => Object.values(item));
    sizeRowData.push(TotalRow[0]);
    worksheet2.mergeCells(
      "B" + sizeTableMergeIndex + ":F" + sizeTableMergeIndex
    );
    worksheet2.getCell("B" + sizeTableMergeIndex).value = "Size Proposal";
    worksheet2.getCell("B" + sizeTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("B" + sizeTableMergeIndex).font = {
      bold: true,
    };
    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("B" + (sizeTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("B" + (sizeTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[1 + i] + sizeTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(colorCell[1 + i] + sizeTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[4].length + 1; j++) {
        worksheet2.getCell(
          colorCell[1 + i] + (sizeTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[4].length + 1) {
          worksheet2.getCell(
            colorCell[1 + i] + (sizeTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }

    worksheet2.addTable({
      name: "sizeProposal",
      ref: "B" + sizeTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Size" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: sizeRowData,
    });

    //color table data
    let colorTableIndex = totalTableIndex + request.recordsets[1].length + 3;
    let colorTableMergeIndex = colorTableIndex - 1;
    let colorRowData = request.recordsets[5].map((item) => Object.values(item));
    colorRowData.push(TotalRow[0]);
    worksheet2.mergeCells(
      "H" + colorTableMergeIndex + ":L" + colorTableMergeIndex
    );
    worksheet2.getCell("H" + colorTableMergeIndex).value = "Color Proposal";
    worksheet2.getCell("H" + colorTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("H" + colorTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("H" + (colorTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("H" + (colorTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[7 + i] + colorTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(colorCell[7 + i] + colorTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[5].length + 1; j++) {
        worksheet2.getCell(
          colorCell[7 + i] + (colorTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[5].length + 1) {
          worksheet2.getCell(
            colorCell[7 + i] + (colorTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "colorProposal",
      ref: "H" + colorTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Color" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: colorRowData,
    });

    //quality table data
    let qualityTableIndex = colorTableIndex + request.recordsets[5].length + 5;
    let qualityTableMergeIndex = qualityTableIndex - 1;
    let qualityRowData = request.recordsets[6].map((item) =>
      Object.values(item)
    );
    qualityRowData.push(TotalRow[0]);
    worksheet2.mergeCells(
      "H" + qualityTableMergeIndex + ":L" + qualityTableMergeIndex
    );
    worksheet2.getCell("H" + qualityTableMergeIndex).value = "Quality Proposal";
    worksheet2.getCell("H" + qualityTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("H" + qualityTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("H" + (qualityTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("H" + (qualityTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[7 + i] + qualityTableIndex.toString()).fill =
      {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(
        colorCell[7 + i] + qualityTableIndex.toString()
      ).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[6].length + 1; j++) {
        worksheet2.getCell(
          colorCell[7 + i] + (qualityTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[6].length + 1) {
          worksheet2.getCell(
            colorCell[7 + i] + (qualityTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "qualityProposal",
      ref: "H" + qualityTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Quality" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: qualityRowData,
    });

    //cut table data
    let cutTableIndex = totalTableIndex + request.recordsets[1].length + 3;
    let cutTableMergeIndex = cutTableIndex - 1;
    let cutRowData = request.recordsets[8].map((item) => Object.values(item));
    cutRowData.push(TotalRow[0]);
    worksheet2.mergeCells("N" + cutTableMergeIndex + ":R" + cutTableMergeIndex);
    worksheet2.getCell("N" + cutTableMergeIndex).value = "Cut Proposal";
    worksheet2.getCell("N" + cutTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("N" + cutTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("N" + (cutTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("N" + (cutTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[13 + i] + cutTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(colorCell[13 + i] + cutTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[8].length + 1; j++) {
        worksheet2.getCell(
          colorCell[13 + i] + (cutTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[8].length + 1) {
          worksheet2.getCell(
            colorCell[13 + i] + (cutTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "cutProposal",
      ref: "N" + cutTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Cut" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: cutRowData,
    });

    //polish table data
    let polishTableIndex = cutTableIndex + request.recordsets[8].length + 5;
    let polishTableMergeIndex = polishTableIndex - 1;
    let polishRowData = request.recordsets[9].map((item) =>
      Object.values(item)
    );
    polishRowData.push(TotalRow[0]);
    worksheet2.mergeCells(
      "N" + polishTableMergeIndex + ":R" + polishTableMergeIndex
    );
    worksheet2.getCell("N" + polishTableMergeIndex).value = "Polish Proposal";
    worksheet2.getCell("N" + polishTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("N" + polishTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("N" + (polishTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("N" + (polishTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[13 + i] + polishTableIndex.toString()).fill =
      {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(
        colorCell[13 + i] + polishTableIndex.toString()
      ).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[8].length + 1; j++) {
        worksheet2.getCell(
          colorCell[13 + i] + (polishTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[8].length + 1) {
          worksheet2.getCell(
            colorCell[13 + i] + (polishTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "PolishProposal",
      ref: "N" + polishTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Polish" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: polishRowData,
    });

    //sym table data
    let symTableIndex = polishTableIndex + request.recordsets[9].length + 5;
    let symTableMergeIndex = symTableIndex - 1;
    let symRowData = request.recordsets[10].map((item) => Object.values(item));
    symRowData.push(TotalRow[0]);
    worksheet2.mergeCells("N" + symTableMergeIndex + ":R" + symTableMergeIndex);
    worksheet2.getCell("N" + symTableMergeIndex).value = "Symmetry Proposal";
    worksheet2.getCell("N" + symTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("N" + symTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("N" + (symTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("N" + (symTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[13 + i] + symTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(colorCell[13 + i] + symTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[10].length + 1; j++) {
        worksheet2.getCell(
          colorCell[13 + i] + (symTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[10].length + 1) {
          worksheet2.getCell(
            colorCell[13 + i] + (symTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "symProposal",
      ref: "N" + symTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Symmetry" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: symRowData,
    });

    //floro table data
    let floroTableIndex = symTableIndex + request.recordsets[10].length + 5;
    let floroTableMergeIndex = floroTableIndex - 1;
    let floroRowData = request.recordsets[11].map((item) =>
      Object.values(item)
    );
    floroRowData.push(TotalRow[0]);
    worksheet2.mergeCells(
      "N" + floroTableMergeIndex + ":R" + floroTableMergeIndex
    );
    worksheet2.getCell("N" + floroTableMergeIndex).value = "Floro Proposal";
    worksheet2.getCell("N" + floroTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("N" + floroTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("N" + (floroTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("N" + (floroTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[13 + i] + floroTableIndex.toString()).fill =
      {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(
        colorCell[13 + i] + floroTableIndex.toString()
      ).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[11].length + 1; j++) {
        worksheet2.getCell(
          colorCell[13 + i] + (floroTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[11].length + 1) {
          worksheet2.getCell(
            colorCell[13 + i] + (floroTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "floroProposal",
      ref: "N" + floroTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Floro" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: floroRowData,
    });

    //inc table data
    let incTableIndex = floroTableIndex + request.recordsets[11].length + 5;
    let incTableMergeIndex = incTableIndex - 1;
    let incRowData = request.recordsets[12].map((item) => Object.values(item));
    incRowData.push(TotalRow[0]);
    worksheet2.mergeCells("N" + incTableMergeIndex + ":R" + incTableMergeIndex);
    worksheet2.getCell("N" + incTableMergeIndex).value = "Inc Proposal";
    worksheet2.getCell("N" + incTableMergeIndex).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet2.getCell("N" + incTableMergeIndex).font = {
      bold: true,
    };

    // background color and fill border
    for (let i = 0; i < 5; i++) {
      worksheet2.getCell("N" + (incTableIndex - 1).toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell("N" + (incTableIndex - 1).toString()).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet2.getCell(colorCell[13 + i] + incTableIndex.toString()).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f0f0f0" },
      };
      worksheet2.getCell(colorCell[13 + i] + incTableIndex.toString()).border =
      {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      for (let j = 1; j <= request.recordsets[12].length + 1; j++) {
        worksheet2.getCell(
          colorCell[13 + i] + (incTableIndex + j).toString()
        ).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (j == request.recordsets[12].length + 1) {
          worksheet2.getCell(
            colorCell[13 + i] + (incTableIndex + j).toString()
          ).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "c5d9f1" },
          };
        }
      }
    }
    worksheet2.addTable({
      name: "incProposal",
      ref: "N" + incTableIndex,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "",
        showRowStripes: true,
        showFirstColumn: true,
      },
      columns: [
        { name: "Inc" },
        { name: "Pcs" },
        { name: "Weight" },
        { name: "Carat%" },
        { name: "Price%" },
      ],
      rows: incRowData,
    });

    workbook.xlsx.writeBuffer().then(function (buffer) {
      let xlsData = Buffer.concat([buffer]);
      res
        .writeHead(200, {
          "Content-Length": Buffer.byteLength(xlsData),
          "Content-Type": "application/vnd.ms-excel",
          "Content-disposition": "attachment;filename=TENDAR.xlsx",
        })
        .end(xlsData);
    });
  } catch (err) {
    console.log(err);
    res.json({ success: 0, data: err });
  }
};

exports.RapCalcTendarDel = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        let IP = req.headers["x-forwarded-for"] || req.connection.remoteAddress;

        request.input("L_CODE", sql.VarChar(50), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("TAG", sql.VarChar(10), req.body.TAG);
        request.input("PRDTYPE", sql.VarChar(10), req.body.PRDTYPE);
        request.input("PLANNO", sql.Int, parseInt(req.body.PLANNO));
        request.input("EMP_CODE", sql.VarChar(10), req.body.EMP_CODE);
        request.input("PTAG", sql.VarChar(10), req.body.PTAG);
        request.input("IUSER", sql.VarChar(10), req.body.IUSER);
        request.input("IP", sql.VarChar(10), IP);
        request.input("PROC_CODE", sql.Int, parseInt(req.body.PROC_CODE));

        request = await request.execute("USP_RapCalTendarDel");

        res.json({ success: 1, data: "" });
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.RapCalcTendarVidUpload = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        let IP = req.headers["x-forwarded-for"] || req.connection.remoteAddress;

        request.input("L_CODE", sql.VarChar(50), req.body.L_CODE);
        request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
        request.input("SECURE_URL", sql.VarChar(500), req.body.SECURE_URL);
        request.input("URL", sql.VarChar(500), req.body.URL);
        request.input("CLOUDID", sql.VarChar(500), req.body.CLOUDID);
        request.input("PUBLICID", sql.VarChar(500), req.body.PUBLICID);

        request = await request.execute("USP_TendarVidUpload");

        res.json({ success: 1, data: "" });
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.RapCalcTendarVidDisp = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(7991), req.body.L_CODE);
        request.input("FSRNO", sql.Int, parseInt(req.body.FSRNO));
        request.input("TSRNO", sql.Int, parseInt(req.body.TSRNO));

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

exports.RapCalcTendarVidDelete = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(7991), req.body.L_CODE);
        request.input("FSRNO", sql.Int, parseInt(req.body.FSRNO));
        request.input("TSRNO", sql.Int, parseInt(req.body.TSRNO));

        request = await request.execute("USP_TendarVidDelete");

        res.json({ success: 1, data: request.recordset });
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.RapCalExport = async (req, res) => {
  try {
    var request = new sql.Request();

    request.input("L_CODE", sql.VarChar(10), req.body.L_CODE);
    request.input("SRNO", sql.Int, parseInt(req.body.SRNO));
    request.input("TAG", sql.VarChar(3), req.body.TAG);
    request.input("EMP_CODE", sql.VarChar(5), req.body.EMP_CODE);

    request = await request.execute("usp_RapCalExport");

    const Excel = require("exceljs");
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("Plan");

    worksheet.mergeCells("A1:P1");
    worksheet.getCell("A1").value = req.body.HEADER;
    worksheet.getCell("A1").alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet.getCell("A1").font = { bold: true };

    worksheet.mergeCells("S1:AC1");
    worksheet.getCell("S1").value = "IGI";
    worksheet.getCell("S1").alignment = {
      vertical: "middle",
      horizontal: "center",
      bgColor: { argb: "FFFF00" },
    };
    worksheet.getCell("S1").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" },
      font: { bold: true },
    };
    worksheet.getCell("R1").font = { bold: true };
    // worksheet.getCell('O1').fgColor = { background:  'FFFF00' };

    // Main table data
    worksheet.columns = [
      { key: "NO" },
      { key: "SHAPE" },
      { key: "CUT" },
      { key: "POLISH" },
      { key: "Col" },
      { key: "RAPTYPE" },
      { key: "Qua" },
      { key: "INC" },
      { key: "FLO" },
      { key: "SHADE" },
      { key: "MILKY" },
      { key: "WHINC" },
      { key: "ORAP" },
      { key: "DIS" },
      { key: "RATE" },
      { key: "AMT" },
      { key: "COMMENT", width: 10 },
      { key: "DIAMETER / HIGHT / RATIO" },
      { key: "IGICut" },
      { key: "IGICol" },
      { key: "IGIQua" },
      { key: "IGIFlo" },
      { key: "IGISHADE" },
      { key: "IGIMILKY" },
      { key: "IGIWHINC"},
      { key: "IGIORAP" },
      { key: "IGIDIS" },
      { key: "IGIRATE" },
      { key: "IGIAMT" },
    ];

    worksheet.getRow(2).values = [
      "NO",
      "SHAPE",
      "CUT",
      "POLISH",
      "Col",
      "RAPTYPE",
      "Qua",
      "INC",
      "FLO",
      "SHADE",
      "MILKY",
      "WHINC",
      "ORAP",
      "DIS",
      "RATE",
      "AMT",
      "COMMENT",
      "DIAMETER / HIGHT / RATIO",
      "Cut",
      "Col",
      "Qua",
      "Flo",
      "SHADE",
      "MILKY",
      "WHINC",
      "ORAP",
      "DIS",
      "RATE",
      "AMT",
    ];

    // worksheet.addRow({}).commit();
    const Data = request.recordset;
    // console.log(Data)
    const Plans = new Set(Data.map((item) => item.NO));
    const totalRows = Data.length + Plans.size + 3;
    const colID = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W",
      "X",
      "Y",
      "Z",
      "AA",
      "AB",
      "AC",
      "AD"
    ];
    // function to repeat paln no only once
    let NO = "";
    let match = false;
    for (let index = 0; index < Data.length; index++) {
      // if (index === 0) console.log(match, Data[index].NO, NO);
      match = Data[index].NO === NO ? true : false;
      if (match === false) {
        NO = Data[index].NO;
      } else {
        Data[index].NO = "";
      }
      // console.log(matzch, NO, Data[index].NO);
    }
    worksheet.mergeCells("R1:R2");
    worksheet.getCell("R1").value = "DIAMETER / HIGHT / RATIO";
    worksheet.getCell("R1").alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet.getCell("R1").font = { bold: true };
    let oldno = 0
    Data.forEach(function (record, index) {

      if (record["NO"]) {
        worksheet.addRow({}).commit();
      }
      worksheet.addRow({
        NO: record["NO"],
        SHAPE: record["SHAPE"],
        POLISH: record["POLISH"] === 0 ? "" : record["POLISH"],
        Col: record["Col"],
        RAPTYPE: record["RAPTYPE"],
        CUT: record["CUT"],
        Qua: record["Qua"],
        INC: record["INC"],
        FLO: record["FLO"],
        SHADE: record["SHADE"],
        MILKY: record["MILKY"],
        ORAP: record["ORAP"] === 0 ? "" : record["ORAP"],
        DIS: record["DIS"] === 0 ? "" : record['DIS'],
        RATE: "",
        AMT: "",
        COMMENT: record["COMMENT"],
        IGICol: record["IGIC_NAME"],
        IGIQua: record["IGIQ_NAME"],
        IGICut: record["IGICUT"],
        IGIFlo: record["IGIFL_NAME"],
        IGIORAP: record["IGIORAP"],
        IGIDIS: record["IGIDIS"],
        IGIRATE: "",
        IGIAMT: "",
        IGISHADE: record["IGISHADE"],
        IGIMILKY: record["IGIMILKY"],
        WHINC: record["WHINC"],
        IGIWHINC: record["WHINC"]
      });
    });

    // console.log("Datalength",Data.length +4)
    // console.log("totalrows",totalRows)
    let id1 = 4;
    for (let j = 0; j < totalRows; j++) {
      index = j
      ID = index + 4
      // console.log("index", j + 4)
      // console.log("ID", ID)
      let Avg = `=IFERROR(M${ID}-(N${ID}*M${ID}/100), "")`
      let ml = `=IFERROR((O${ID}*D${ID}),"")`
      let AvgIGI = `=Z${ID}-(AA${ID}*Z${ID}/100)`
      let mlIGI = `=(AB${ID}*D${ID})`
      worksheet.getCell('O' + (j + 4)).value = { formula: Avg }
      worksheet.getCell('P' + (j + 4)).value = { formula: ml }
      worksheet.getCell('AC' + (j + 4)).value = { formula: mlIGI }
      worksheet.getCell('AB' + (j + 4)).value = { formula: AvgIGI }

      if (
        !worksheet.getCell("B" + j).value &&
        worksheet.getCell("P" + j).value
      ) {
        worksheet.getCell("P" + j).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "95C1DD" },
        };

        // console.log("ID1",id1)
        // console.log(ID-5)

        let sum = `=SUM(P${id1}:P${ID - 5})`
        // console.log(sum)
        worksheet.getCell("P" + j).value = { formula: sum }

        if (
          !worksheet.getCell("Q" + j).value &&
          worksheet.getCell("AC" + j).value
        ) {
          worksheet.getCell("AC" + j).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "95C1DD" },
          };
          let sumigi = `=SUM(AC${id1}:AC${ID - 5})`

          worksheet.getCell("AC" + j).value = { formula: sumigi }

        }
        id1 = ID - 3

      } else if (!worksheet.getCell("B" + j).value &&
        !worksheet.getCell("P" + j).value) {
        worksheet.getCell("P" + j).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFFF" },
        };
        worksheet.getCell("O" + j).value = ""
      }

    }

    for (let i = 0; i < totalRows; i++) {
      ID1 = 4;
      ID2 = i - 1;

      worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
        row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
          for (var i = 1; i < colID.length + 1; i++) {
            row.getCell(i).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
        });
      });
      // if (
      //   !worksheet.getCell("B" + i.toString()).value &&
      //   worksheet.getCell("M" + i.toString()).value
      // ) {
      //   worksheet.getCell("M" + i.toString()).fill = {
      //     type: "pattern",
      //     pattern: "solid",
      //     fgColor: { argb: "95C1DD" },
      //   };
      //   let sum = `=SUM(M${ID1}:M${ID2})`
      //   console.log(ID1)
      //   console.log(ID2)
      //   worksheet.getCell("M" + i.toString()).value = { formula: sum }


      // }
      // if (
      //   !worksheet.getCell("P" + i.toString()).value &&
      //   worksheet.getCell("W" + i.toString()).value
      // ) {
      //   worksheet.getCell("W" + i.toString()).fill = {
      //     type: "pattern",
      //     pattern: "solid",
      //     fgColor: { argb: "95C1DD" },
      //   };
      //   let sumigi = `=SUM(W${ID1}:W${ID2})`

      //   worksheet.getCell("W" + i.toString()).value = { formula: sumigi }

      // }


    }

    workbook.xlsx.writeBuffer().then(function (buffer) {
      let xlsData = Buffer.concat([buffer]);
      res
        .writeHead(200, {
          "Content-Length": Buffer.byteLength(xlsData),
          "Content-Type": "application/vnd.ms-excel",
          "Content-disposition":
            "attachment;filename=" + req.body.HEADER + ".xls",
        })
        .end(xlsData);
    });


  } catch (err) {
    console.log(err);
    res.json({ success: 0, data: err });
  }
};

exports.RapCalcPrdQuery = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;
      let RapArr = req.body;
      let status = true;
      let ErrorRap = [];
      for (let i = 0; i < RapArr.length; i++) {
        try {
          var request = new sql.Request();
          var date = new Date();
          function createDateAsUTC(date) {
            return new Date(
              Date.UTC(
                date.getFullYear(),
                date.getMonth(),
                date.getDate(),
                date.getHours(),
                date.getMinutes(),
                date.getSeconds()
              )
            );
          }
          let FinalDate = createDateAsUTC(date);

          request.input("L_CODE", sql.VarChar(10), RapArr[i].L_CODE);
          request.input("SRNO", sql.Int, parseInt(RapArr[i].SRNO));
          request.input("TAG", sql.VarChar(3), RapArr[i].TAG);
          if (RapArr[i].I_DATE != "") {
            request.input("I_DATE", sql.DateTime2, FinalDate);
          }
          if (RapArr[i].I_TIME != "") {
            request.input("I_TIME", sql.DateTime2, FinalDate);
          }
          request.input("EMP_CODE", sql.VarChar(10), RapArr[i].EMP_CODE);
          request.input("S_CODE", sql.VarChar(5), RapArr[i].S_CODE);
          request.input("Q_CODE", sql.Int, parseInt(RapArr[i].Q_CODE));
          request.input("C_CODE", sql.Int, parseInt(RapArr[i].C_CODE));
          request.input(
            "CARAT",
            sql.Decimal(10, 3),
            parseFloat(RapArr[i].CARAT)
          );
          request.input("CUT_CODE", sql.Int, parseInt(RapArr[i].CUT_CODE));
          request.input("POL_CODE", sql.Int, parseInt(RapArr[i].POL_CODE));
          request.input("SYM_CODE", sql.Int, parseInt(RapArr[i].SYM_CODE));
          request.input("FL_CODE", sql.Int, parseInt(RapArr[i].FL_CODE));
          request.input("IN_CODE", sql.Int, parseInt(RapArr[i].IN_CODE));
          request.input("SH_CODE", sql.Int, parseInt(RapArr[i].SH_CODE));
          request.input("RAPTYPE", sql.VarChar(5), RapArr[i].RAPTYPE);
          request.input("RATE", sql.Decimal(10, 2), parseFloat(RapArr[i].RATE));
          request.input("RTYPE", sql.VarChar(5), RapArr[i].RTYPE);

          request = await request.execute("usp_RapCalcPrdQuery");
        } catch (err) {
          status = false;
          ErrorRap.push(RapArr[i]);
        }
      }

      res.json({ success: status, data: ErrorRap });
    }
  });
};

exports.TendarLotDelete = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      try {
        const TokenData = await authData;

        var request = new sql.Request();

        request.input("L_CODE", sql.VarChar(50), req.body.L_CODE);

        request = await request.execute("USP_TendarLotDelete");

        res.json({ success: true, data: "" });
      } catch (err) {
        console.log(err);
        res.json({ success: false, data: err });
      }
    }
  });
};
