const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.AvgSizeMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SZTYPE_CODE', sql.Int ,req.body.SZTYPE_CODE)
                request.input('SZ_CODE', sql.Int ,req.body.SZ_CODE)
                request.input('SZ_NAME', sql.VarChar(30),req.body.SZ_NAME)
                request.input('F_SIZE', sql.Numeric(10,3), req.body.F_SIZE)
                request.input('T_SIZE', sql.Numeric(10,3), req.body.T_SIZE)
                request.input('OSZ_CODE', sql.Int ,req.body.OSZ_CODE)
                request.input('ORD', sql.Int ,req.body.ORD)
                
                request = await request.execute('USP_AvgSizeMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}


// [dbo].[USP_AvgSizeMastDelete]
// @SZTYPE_CODE Int=0

exports.AvgSizeMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SZTYPE_CODE', sql.Int ,req.body.SZTYPE_CODE)

                request = await request.execute('USP_AvgSizeMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}


// [dbo].[USP_AvgSizeMastFill]
// @SZTYPE_CODE Int=0

exports.AvgSizeMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SZTYPE_CODE', sql.Int ,req.body.SZTYPE_CODE)
        
                request = await request.execute('USP_AvgSizeMastFill');

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