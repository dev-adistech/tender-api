const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.AvgQuaMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SZTYPE_CODE', sql.Int, req.body.SZTYPE_CODE)
                request.input('Q_CODE', sql.Int, req.body.Q_CODE)
                
                request = await request.execute('USP_AvgQuaMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.AvgQuaMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SZTYPE_CODE', sql.Int, req.body.SZTYPE_CODE)

                request = await request.execute('USP_AvgQuaMastFill');

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

exports.AvgQuaMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SZTYPE_CODE', sql.Int, req.body.SZTYPE_CODE)
                request.input('Q_CODE', sql.Int, req.body.Q_CODE)
                request.input('Q_NAME', sql.VarChar(35),req.body.Q_NAME)
                request.input('ORD', sql.Int, req.body.ORD)
        
                request = await request.execute('USP_AvgQuaMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}