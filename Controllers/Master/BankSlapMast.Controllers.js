const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');


// [dbo].[USP_BankSlabSave]
// @F_SAL INT=0,
// @T_SAL INT=0,
// @PER NUMERIC(10,3)=0.00


exports.BankSlabSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('F_SAL', sql.Int, parseInt(req.body.F_SAL))
                request.input('T_SAL', sql.Int, parseInt(req.body.T_SAL))
                request.input('PER', sql.Numeric(10,3), req.body.PER)
        
                request = await request.execute('USP_BankSlabSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

// [dbo].[USP_BankSlabDelete]
// @F_SAL INT=0,
// @T_SAL INT=0

exports.BankSlabDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('F_SAL', sql.Int, parseInt(req.body.F_SAL))
                request.input('T_SAL', sql.Int, parseInt(req.body.T_SAL))
                
                request = await request.execute('USP_BankSlabDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

// [dbo].[USP_BankSlabFill]

exports.BankSlabFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('USP_BankSlabFill');

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