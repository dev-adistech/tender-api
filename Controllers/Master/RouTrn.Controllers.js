const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.RouTrnFillTrnNo = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('USP_RouTrnFillTrnNo');

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

exports.RouTrnFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('USP_RouTrnFill');

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

exports.RouTrnSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('TRNNO', sql.Int, parseInt(req.body.TRNNO))
                request.input('R_CODE', sql.VarChar(5), req.body.R_CODE)
                request.input('R_NAME', sql.VarChar(10), req.body.R_NAME)
                request.input('P_CODE', sql.VarChar(10), req.body.P_CODE)
                request.input('B_CODE', sql.VarChar(10), req.body.B_CODE)
                request.input('SITE', sql.VarChar(20), req.body.SITE)
                request.input('SIZE', sql.Decimal(10, 2), parseFloat(req.body.SIZE))
                if (req.body.I_DATE) { request.input('I_DATE', sql.DateTime2, new Date(req.body.I_DATE)) }
                request.input('I_CARAT', sql.Decimal(10, 3), parseFloat(req.body.I_CARAT))
                request.input('I_PCS', sql.Int, parseInt(req.body.I_PCS))
                request.input('RATE', sql.Int, parseInt(req.body.RATE))

                request = await request.execute('USP_RouTrnSave');


                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.RouTrnDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if (req.body.TRNNO) { request.input('TRNNO', sql.Int, parseInt(req.body.TRNNO)) }
                if (req.body.R_CODE) { request.input('R_CODE', sql.VarChar(5), req.body.R_CODE) }


                request = await request.execute('USP_RouTrnDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}