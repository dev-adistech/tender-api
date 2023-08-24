const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.IncMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('USP_IncMastFill');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })

                } else {
                    res.json({ success: 1, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.IncMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('IN_CODE', sql.Int, parseInt(req.body.I_CODE))
                request.input('IN_NAME', sql.VarChar(64), req.body.I_NAME)
                request.input('ORD', sql.Int, parseInt(req.body.I_ORD))

                request = await request.execute('USP_IncMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.IncMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('IN_CODE', sql.Int, parseInt(req.body.I_CODE))

                request = await request.execute('USP_IncMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.IncMastGetMaxId = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('I_TYPE', sql.VarChar(32), req.body.I_TYPE)
                request.input('I_CODE', sql.Int, parseInt(req.body.I_CODE))

                request = await request.execute('USP_IncMastGetMaxId');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "Not Found" })
                }
            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}