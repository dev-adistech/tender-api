const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.RapLabMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('LAB_TYPE', sql.VarChar(10), req.body.LAB_TYPE)
                request.input('COL_TYPE', sql.VarChar(10), req.body.COL_TYPE)
                request.input('SHP_TYPE', sql.VarChar(2), req.body.SHP_TYPE)

                request = await request.execute('usp_RapLabMastFill');

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

exports.RapLabMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('LAB_TYPE', sql.VarChar(10), req.body.LAB_TYPE)
                request.input('COL_TYPE', sql.VarChar(10), req.body.COL_TYPE)
                request.input('SHP_TYPE', sql.VarChar(2), req.body.SHP_TYPE)
                request.input('CODE', sql.Int, parseInt(req.body.CODE))
                request.input('FSIZE', sql.Decimal(10, 3), parseFloat(req.body.FSIZE));
                request.input('TSIZE', sql.Decimal(10, 3), parseFloat(req.body.TSIZE));
                request.input('RATE', sql.Decimal(10, 3), parseFloat(req.body.RATE));

                request = await request.execute('usp_RapLabMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.RapLabMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('LAB_TYPE', sql.VarChar(10), req.body.LAB_TYPE)
                request.input('COL_TYPE', sql.VarChar(10), req.body.COL_TYPE)
                request.input('SHP_TYPE', sql.VarChar(2), req.body.SHP_TYPE)
                request.input('CODE', sql.VarChar(2), req.body.CODE)

                request = await request.execute('usp_RapLabMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}