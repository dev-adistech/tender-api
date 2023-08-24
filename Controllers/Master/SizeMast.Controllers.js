const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.SizeMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SZ_TYPE', sql.VarChar(10), req.body.SZ_TYPE)

                request = await request.execute('USP_SizeMastFill');

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

exports.SizeMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SZ_TYPE', sql.VarChar(10), req.body.SZ_TYPE)
                request.input('SZ_CODE', sql.Int, parseInt(req.body.SZ_CODE))
                request.input('F_SIZE', sql.Decimal(8, 3), parseFloat(req.body.F_SIZE))
                request.input('T_SIZE', sql.Decimal(8, 3), parseFloat(req.body.T_SIZE))
                request.input('SZ_GROUP', sql.VarChar(16), req.body.SZ_GROUP)
                request.input('OSZ_CODE', sql.Int, parseInt(req.body.OSZ_CODE))

                request = await request.execute('USP_SizeMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.SizeMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SZ_TYPE', sql.VarChar(10), req.body.SZ_TYPE)
                if (req.body.SZ_CODE) { request.input('SZ_CODE', sql.Int, parseInt(req.body.SZ_CODE)) }

                request = await request.execute('USP_SizeMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}