const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.ViewParaFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if (req.body.FORMNAME) { request.input('FORMNAME', sql.VarChar(64), req.body.FORMNAME) }

                request = await request.execute('USP_ViewParaFill');
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

exports.ViewParaComboFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('USP_ViewParaComboFill');

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

exports.ViewParaSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('FORMNAME', sql.VarChar(64), req.body.FORMNAME)
                request.input('FIELDNAME', sql.VarChar(32), req.body.FIELDNAME)
                request.input('DISPNAME', sql.VarChar(32), req.body.DISPNAME)
                request.input('COLWIDTH', sql.Int, parseInt(req.body.COLWIDTH))
                request.input('HEADERALIGN', sql.VarChar(32), req.body.HEADERALIGN)
                request.input('HEADERCOLOR', sql.VarChar(32), req.body.HEADERCOLOR)
                request.input('HEADERFONTSIZE', sql.Int, parseInt(req.body.HEADERFONTSIZE))
                request.input('CELLALIGN', sql.VarChar(16), req.body.CELLALIGN)
                request.input('FONTCOLOR', sql.VarChar(32), req.body.FONTCOLOR)
                request.input('CELLFONTSIZE', sql.Int, parseInt(req.body.CELLFONTSIZE))
                request.input('BACKCOLOR', sql.VarChar(32), req.body.BACKCOLOR)
                request.input('FONTNAME', sql.VarChar(64), req.body.FONTNAME)
                request.input('ISBOLD', sql.Bit, req.body.ISBOLD)
                request.input('ISRESIZE', sql.Bit, req.body.ISRESIZE)
                request.input('ORD', sql.Int, parseInt(req.body.ORD))
                request.input('FORMAT', sql.VarChar(32), req.body.FORMAT)
                request.input('COLUMNSTYLE', sql.VarChar(50), req.body.COLUMNSTYLE)
                request.input('GROUPKEY', sql.VarChar(50), req.body.GROUPKEY)
                request.input('CUSTSUMMARY', sql.VarChar(50), req.body.CUSTSUMMARY)
                request.input('MAXLEN', sql.Int, parseInt(req.body.MAXLEN))
                request.input('DISP', sql.Bit, req.body.DISP)
                request.input('LOCK', sql.Bit, req.body.LOCK)
                request.input('ISMERGE', sql.Bit, req.body.ISMERGE)

                request = await request.execute('USP_ViewParaSave');

                res.json({ success: 1, data: '' })


            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.ViewParaDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('FORMNAME', sql.VarChar(64), req.body.FORMNAME)
                request.input('FIELDNAME', sql.VarChar(32), req.body.FIELDNAME)

                request = await request.execute('USP_ViewParaDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}
