const sql = require("mssql");
const jwt = require('jsonwebtoken');
const Excel = require('exceljs');


var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

var { preServerconn } = require('../../config/db/rapsql');

exports.LabResultGet = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if(err){
            res.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                const pool = await preServerconn();
                let request = await pool.request()

                request.input("LSRNOLIST", sql.VarChar(sql.MAX), req.body.LSRNOLIST)
                request.input("L_CODE", sql.VarChar(sql.MAX), req.body.L_CODE)
                request.input("LSRNO", sql.Int, req.body.LSRNO)
                request.input("LTAG", sql.VarChar(4), req.body.LTAG)

                request = await request.execute('PRED_LabResultGet')

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch(err){
                res.json({success: 0, data: err})
                console.log(err)
            }

        }
    })
}

exports.GetLot = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('USP_GetLot');

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

exports.GetStonId = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                request.input("TYPE", sql.VarChar(10), req.body.TYPE)
                request = await request.execute('USP_GetStonId');

                
                res.json({ success: 1, data: request.recordset })
                
            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.CertiResultSave = async (req, res) => {

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

                    request.input('L_CODE', sql.VarChar(10), PerArr[i].L_CODE)
                    request.input('SRNO', sql.Int, parseInt(PerArr[i].SRNO))
                    request.input('TAG', sql.VarChar(5), PerArr[i].TAG)
                    request.input('STONEID', sql.BigInt, PerArr[i].STONEID)
                    request.input('S_NAME', sql.VarChar(100), PerArr[i].S_NAME)
                    request.input('C_NAME', sql.VarChar(50), PerArr[i].C_NAME)
                    request.input('Q_NAME', sql.VarChar(50), PerArr[i].Q_NAME)
                    request.input('CARAT', sql.Numeric(10,3), PerArr[i].CARAT)
                    request.input('CUT_CODE', sql.VarChar(50), PerArr[i].CUT_CODE)
                    request.input('POL_CODE', sql.VarChar(50), PerArr[i].POL_CODE)
                    request.input('SYM_CODE', sql.VarChar(50), PerArr[i].SYM_CODE)
                    request.input('FL_CODE', sql.VarChar(50), PerArr[i].FL_CODE)
                    request.input('FCOLOR', sql.VarChar(150), PerArr[i].FCOLOR)
                    request.input('FINTENSITY', sql.VarChar(150), PerArr[i].FINTENSITY)
                    request.input('FOVERTONE', sql.VarChar(150), PerArr[i].FOVERTONE)
                    request.input('SDATE', sql.DateTime2, new Date(PerArr[i].SDATE))
                    request.input('CR_NAME', sql.VarChar(10), PerArr[i].CR_NAME)
                    request.input('GI_DATE', sql.DateTime2, new Date(PerArr[i].GI_DATE))
                    request.input('TYPE', sql.VarChar(5), PerArr[i].TYPE)
                   

                    request = await request.execute('USP_CertiResultSave');
                } catch (err) {
                    status = false;
                    ErrorPer.push(PerArr[i])
                }
            }

            res.json({ success: status, data: ErrorPer })
        }
    });
}

exports.LabResultImportUpdate = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if(err){
            res.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                const pool = await preServerconn();
                let request = await pool.request()

                request.input("STONEID", sql.VarChar(sql.MAX), req.body.STONEID)

                request = await request.execute('PRED_LabResultImportUpdate')

                res.json({ success: 1, data: request.recordset })
                

            } catch(err){
                res.json({success: 0, data: err})
                console.log(err)
            }

        }
    })
}