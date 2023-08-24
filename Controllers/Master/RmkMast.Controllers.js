const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');



exports.RmkMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('USP_RemarkMastFill');

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

exports.RmkMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('R_CODE', sql.Int, parseInt(req.body.R_CODE))
                request.input('R_NAME', sql.VarChar(50), req.body.R_NAME)
                request.input('PROC_CODE', sql.VarChar(100), req.body.PROC_CODE)

                request = await request.execute('USP_RemarkMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.RmkMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if (req.body.R_CODE) { request.input('R_CODE', sql.Int, parseInt(req.body.R_CODE)) }

                request = await request.execute('USP_RemarkMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}