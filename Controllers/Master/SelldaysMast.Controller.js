const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.SellDaysMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('USP_SellDaysMastFill');

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

exports.SellDaysMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('S_CODE', sql.VarChar(100), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(100), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(100), req.body.Q_CODE)
                request.input('F_CARAT', sql.Numeric(10,3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10,3), req.body.T_CARAT)
                request.input('LAB', sql.VarChar(100), req.body.LAB)
                request.input('DD', sql.Int, parseInt(req.body.DD))

                request = await request.execute('USP_SellDaysMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.SellDaysMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))

                request = await request.execute('USP_SellDaysMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}