const sql = require("mssql");
const jwt = require('jsonwebtoken');
const Excel = require('exceljs');


var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

var { RapServerConn } = require('../../config/db/rapsql');


exports.ORapParamDisFill = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if(err){
            res.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                const pool = await RapServerConn();
                let request = await pool.request()

                request.input("RAPTYPE", sql.VarChar(5), req.body.RAPTYPE)
                request.input("RTYPE", sql.VarChar(5), req.body.RTYPE)
                request.input("PARAM_NAME", sql.VarChar(20), req.body.PARAM_NAME)

                request = await request.execute('USP_ORapParamDisFill')

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

exports.ORapParamDisSave = async(req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if(err){
            res.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                const pool = await RapServerConn();
                let request = await pool.request()

                request.input("SRNO", sql.Int, req.body.SRNO),
                request.input("S_CODE", sql.VarChar(20),  req.body.S_CODE)
                request.input("F_SIZE", sql.Numeric(10, 3), req.body.F_SIZE,)
                request.input("T_SIZE", sql.Numeric(10, 3), req.body.T_SIZE)
                request.input("PARAM_SIZE", sql.Numeric(10,3),  req.body.PARAM_SIZE,)
                request.input("Q_CODE", sql.VarChar(100),  req.body.Q_CODE)
                request.input("C_CODE", sql.VarChar(100),  req.body.C_CODE)
                request.input("CUT_CODE", sql.VarChar(100),  req.body.CUT_CODE)
                request.input("POL_CODE", sql.VarChar(100),  req.body.POL_CODE)
                request.input("SYM_CODE", sql.VarChar(100),  req.body.SYM_CODE)
                request.input("FL_CODE", sql.VarChar(100),  req.body.FL_CODE)
                request.input("SH_CODE", sql.VarChar(100),  req.body.SH_CODE)
                request.input("PER", sql.Numeric(10, 2), req.body.PER) 
                request.input("RAPTYPE", sql.VarChar(5), req.body.RAPTYPE)
                request.input("RTYPE", sql.VarChar(5), req.body.RTYPE)
                request.input("PARAM_NAME", sql.VarChar(20), req.body.PARAM_NAME)

                request = await request.execute('USP_ORapParamDisSave')

                res.json({ success: 1, data: "" })

            } catch(err){
                res.json({success: 0, data: err})
                console.log(err)
            }

        }
    })   
}

exports.ORapParamDisDelete = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if(err){
            res.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                const pool = await RapServerConn();
                let request = await pool.request()

                request.input("RAPTYPE", sql.VarChar(5), req.body.RAPTYPE)
                request.input("RTYPE", sql.VarChar(5), req.body.RTYPE)
                request.input("PARAM_NAME", sql.VarChar(20), req.body.PARAM_NAME)

                request = await request.execute('USP_ORapParamDisDelete')

                    res.json({ success: 1, data: request.recordset })

            } catch(err){
                console.log(err)
                res.json({success: 0, data: err})
            }

        }
    })
}