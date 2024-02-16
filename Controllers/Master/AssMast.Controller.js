const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');


exports.TendarAssMstSave = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();
        let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;

        request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
        request.input('DETID', sql.Int, parseInt(req.body.DETID))
        request.input('I_PCS', sql.Int, parseInt(req.body.I_PCS))
        request.input('I_CARAT', sql.Numeric(10, 3), req.body.I_CARAT)
        request.input('N_PCS', sql.Int, parseInt(req.body.N_PCS))
        request.input('IUSER', sql.VarChar(20), req.body.IUSER)
        request.input('ICOMP', sql.VarChar(30), IP)
        request.input('ROU_FEEL', sql.Int, parseInt(req.body.ROU_FEEL))
        request.input('ST_PRD', sql.Int, parseInt(req.body.ST_PRD))
        request.input('M_QUALITY', sql.VarChar(50), req.body.M_QUALITY)
        request.input('MODEL', sql.VarChar(50), req.body.MODEL)
        request.input('AVG_COLOR', sql.VarChar(50), req.body.AVG_COLOR)
        request.input('AVG_QUALITY', sql.VarChar(50), req.body.AVG_QUALITY)
        request.input('CNT_BLK', sql.Int, parseInt(req.body.CNT_BLK))
        request.input('MIX_TIN', sql.Int, parseInt(req.body.MIX_TIN))
        request.input('AVG_TEN', sql.Int, parseInt(req.body.AVG_TEN))

        request = await request.execute('USP_TendarAssMstSave');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}

exports.TendarAssEntFill = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();
        let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;

        request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
        request.input('DETID', sql.Int, parseInt(req.body.DETID))

        request = await request.execute('USP_TendarAssEntFill');

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

exports.FindRapAss = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();

        request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
        request.input('FQ_CODE', sql.Int, parseInt(req.body.FQ_CODE))
        request.input('TQ_CODE', sql.Int, parseInt(req.body.TQ_CODE))
        request.input('FC_CODE', sql.Int, parseInt(req.body.FC_CODE))
        request.input('TC_CODE', sql.Int, parseInt(req.body.TC_CODE))
        request.input('CARAT', sql.Numeric(10, 3), req.body.CARAT)
        request.input('CUT_CODE', sql.Int, parseInt(req.body.CUT_CODE))
        request.input('FL_CODE', sql.Int, parseInt(req.body.FL_CODE))
        request.input('IN_CODE', sql.Int, parseInt(req.body.IN_CODE))
        request.input('SH_CODE', sql.Int, parseInt(req.body.SH_CODE))
        request.input('ML_CODE', sql.Int, parseInt(req.body.ML_CODE))
        request.input('REF_CODE', sql.Int, parseInt(req.body.REF_CODE))
        request.input('RAPTYPE', sql.VarChar(5), req.body.RAPTYPE)
        request.input('RTYPE', sql.VarChar(5), req.body.RTYPE)

        request = await request.execute('USP_FindRapAss');

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

exports.TendarAssPrdSave = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;
      let PerArr = req.body;
      let status = true;
      let ErrorPer = []
      let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
      for (let i = 0; i < PerArr.length; i++) {
        try {
          var request = new sql.Request();

          request.input('COMP_CODE', sql.VarChar(10), PerArr[i].COMP_CODE)
          request.input('DETID', sql.Int, parseInt(PerArr[i].DETID))
          request.input('DETNO', sql.Int, parseInt(PerArr[i].DETNO))
          request.input('PLANNO', sql.Int, parseInt(PerArr[i].PLANNO))
          request.input('PTAG', sql.VarChar(2), PerArr[i].PTAG)
          request.input('S_CODE', sql.VarChar(5), PerArr[i].S_CODE)
          request.input('FC_CODE', sql.Int, parseInt(PerArr[i].FC_CODE))
          request.input('TC_CODE', sql.Int, parseInt(PerArr[i].TC_CODE))
          request.input('FQ_CODE', sql.Int, parseInt(PerArr[i].FQ_CODE))
          request.input('TQ_CODE', sql.Int, parseInt(PerArr[i].TQ_CODE))
          request.input('CARAT', sql.Numeric(10, 3), PerArr[i].CARAT)
          request.input('CT_CODE', sql.Int, parseInt(PerArr[i].CT_CODE))
          request.input('FL_CODE', sql.Int, parseInt(PerArr[i].FL_CODE))
          request.input('IN_CODE', sql.Int, parseInt(PerArr[i].IN_CODE))
          request.input('SH_CODE', sql.Int, parseInt(PerArr[i].SH_CODE))
          request.input('ML_CODE', sql.Int, parseInt(PerArr[i].ML_CODE))
          request.input('REF_CODE', sql.Int, parseInt(PerArr[i].REF_CODE))
          request.input('RAPTYPE', sql.VarChar(5), PerArr[i].RAPTYPE)
          request.input('LB_CODE', sql.VarChar(2), PerArr[i].LB_CODE)
          request.input('RTYPE', sql.VarChar(5), PerArr[i].RTYPE)
          request.input('ORAP', sql.Numeric(10, 3), PerArr[i].ORAP)
          request.input('RATE', sql.Numeric(10, 3), PerArr[i].RATE)
          request.input('IUSER', sql.VarChar(20), PerArr[i].IUSER)
          request.input('ICOMP', sql.VarChar(30), IP)

          request = await request.execute('USP_TendarAssPrdSave');

          res.json({ success: 1, data: '' })

        } catch (err) {
          res.json({ success: 0, data: err })
        }
      }

      res.json({ success: status, data: ErrorPer })
    }
  });
}

exports.TendarAssTrnSave = async (req, res) => {

  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        var request = new sql.Request();
        let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;

        request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
        request.input('DETID', sql.Int, parseInt(req.body.DETID))
        request.input('DETNO', sql.Int, parseInt(req.body.DETNO))
        request.input('I_PCS', sql.Int, parseInt(req.body.I_PCS))
        request.input('I_CARAT', sql.Numeric(10, 3), parseFloat(req.body.I_CARAT))
        request.input('IUSER', sql.VarChar(20), req.body.IUSER)
        request.input('ICOMP', sql.VarChar(30), IP)
        request.input('ISMC_COL', sql.Bit, req.body.ISMC_COL)

        request = await request.execute('USP_TendarAssTrnSave');

        res.json({ success: 1, data: '' })

      } catch (err) {
        res.json({ success: 0, data: err })
      }
    }
  });
}

