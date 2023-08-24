const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.BonusSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)
                request.input('F_SAL', sql.Int, parseInt(req.body.F_SAL))
                request.input('T_SAL', sql.Int, parseInt(req.body.T_SAL))
                request.input('POL_CODE', sql.Int, parseInt(req.body.POL_CODE))
                request.input('SYM_CODE', sql.Int, parseInt(req.body.SYM_CODE))
                request.input('FLAT', sql.Int, parseInt(req.body.FLAT))
                request.input('PER', sql.Numeric(10,2), req.body.PER)
        
                request = await request.execute('USP_BonusSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.BonusDelte = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)
                request.input('F_SAL', sql.Int, parseInt(req.body.F_SAL))
                request.input('T_SAL', sql.Int, parseInt(req.body.T_SAL))
                request.input('POL_CODE', sql.Int, parseInt(req.body.POL_CODE))
                request.input('SYM_CODE', sql.Int, parseInt(req.body.SYM_CODE))

                request = await request.execute('USP_BonusDelte');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.BonusFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)

                request = await request.execute('USP_BonusFill');

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