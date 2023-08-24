const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.ParamMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('PARAM_NAME', sql.VarChar(50), req.body.PARAM_NAME)
                request.input('S_CODE', sql.VarChar(50), req.body.S_CODE)
                request.input('SZ_CODE', sql.Int, parseInt(req.body.SZ_CODE))

                request = await request.execute('USP_ParamMastFill');

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

exports.ParamMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('PARAM_NAME', sql.VarChar(50), req.body.PARAM_NAME)
                request.input('PARAM_CODE', sql.Int, parseInt(req.body.PARAM_CODE))
                request.input('S_CODE', sql.VarChar(50), req.body.S_CODE)
                request.input('SZ_CODE', sql.Int, parseInt(req.body.SZ_CODE))
                request.input('FMETER', sql.Decimal(10, 2), parseFloat(req.body.FMETER));
                request.input('TMETER', sql.Decimal(10, 2), parseFloat(req.body.TMETER));

                request = await request.execute('USP_ColMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.ParamMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('PARAM_NAME', sql.VarChar(50), req.body.PARAM_NAME)
                request.input('PARAM_CODE', sql.Int, parseInt(req.body.PARAM_CODE))
                request.input('S_CODE', sql.VarChar(50), req.body.S_CODE)
                request.input('SZ_CODE', sql.Int, parseInt(req.body.SZ_CODE))

                request = await request.execute('USP_ParamMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}