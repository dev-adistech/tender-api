const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');


exports.RptParaMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            try {
                var request = new sql.Request();

                request = await request.execute('USP_FrmRepParaFill');

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

exports.RptParaMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('NAME', sql.VarChar(50), req.body.NAME)
                request.input('DISPLAYNAME', sql.VarChar(50), req.body.DISPLAYNAME)
                request.input('CATE_NAME', sql.VarChar(50), req.body.CATE_NAME)
                request.input('DATATYPE', sql.VarChar(50), req.body.DATATYPE)
                request.input('LOOKUP_STRING', sql.VarChar(1000), req.body.LOOKUP_STRING)
                request.input('DATATABLE', sql.VarChar(50), req.body.DATATABLE)
                request.input('DISPLAYMEMBER', sql.VarChar(50), req.body.DISPLAYMEMBER)
                request.input('VALUEMEMBER', sql.VarChar(50), req.body.VALUEMEMBER)
                request.input('COLUMNS', sql.VarChar(1000), req.body.COLUMNS)
                request.input('CAPTIONS', sql.VarChar(1000), req.body.CAPTIONS)

                request = await request.execute('USP_FrmRepParaSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.RptParaMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('NAME', sql.VarChar(50), req.body.NAME)

                request = await request.execute('USP_FrmRepParaDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}