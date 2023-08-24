const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');


// [dbo].[USP_AvgVerMastDelete]
// @V_CODE Int=0

exports.AvgVerMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('V_CODE', sql.Int, parseInt(req.body.V_CODE))

                request = await request.execute('USP_AvgVerMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

// [dbo].[USP_AvgVerMastFill]


exports.AvgVerMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('USP_AvgVerMastFill');

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

exports.AvgVerMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                
                request.input('V_CODE', sql.Int, parseInt(req.body.V_CODE))
                if (req.body.V_DATE) { request.input('V_DATE', sql.DateTime2, new Date(req.body.V_DATE)) }
                request.input('VFLG', sql.Bit, req.body.VFLG)
                request.input('ORD', sql.Int, parseInt(req.body.ORD))

                request = await request.execute('USP_AvgVerMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}