const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.RptMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            try {
                var request = new sql.Request();

                request = await request.execute('USP_FrmRepMastFill');

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

exports.RptMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_NAME', sql.VarChar(50), req.body.DEPT_NAME)
                request.input('REP_ID', sql.Int, parseInt(req.body.REP_ID))
                request.input('PARENT_ID', sql.Int, parseInt(req.body.PARENT_ID))
                request.input('CAT_CODE', sql.VarChar(50), req.body.CAT_CODE)
                request.input('REP_NAME', sql.VarChar(50), req.body.REP_NAME)
                request.input('RPT_NAME', sql.VarChar(1000), req.body.RPT_NAME)
                request.input('SCHEMA_NAME', sql.VarChar(50), req.body.SCHEMA_NAME)
                request.input('SP_NAME', sql.VarChar(50), req.body.SP_NAME)
                request.input('DESCR', sql.VarChar(50), req.body.DESCR)
                request.input('ORD', sql.Int, parseInt(req.body.ORD))
                request.input('RPT_TYPE', sql.VarChar(50), req.body.RPT_TYPE)

                request = await request.execute('USP_FrmRepMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.RptMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_NAME', sql.VarChar(50), req.body.DEPT_NAME)
                request.input('REP_ID', sql.Int, parseInt(req.body.REP_ID))

                request = await request.execute('USP_FrmRepMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}