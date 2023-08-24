const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.RptPerMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(3), req.body.DEPT_CODE)
                request.input('USER_NAME', sql.VarChar(20), req.body.USER_NAME)

                request = await request.execute('usp_FrmRepPerMastFill');

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

exports.RptPerMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            let PerArr = req.body;
            let status = true;
            let ErrorPer = []

            for (let i = 0; i < PerArr.length; i++) {
                try {
                    var request = new sql.Request();

                    request.input('USERNAME', sql.VarChar(10), PerArr[i].USERNAME)
                    request.input('REP_ID', sql.Int, PerArr[i].REP_ID)
                    request.input('VIEWFLG', sql.Bit, PerArr[i].VIEWFLG)

                    request = await request.execute('usp_FrmRepPerMastSave');

                } catch (err) {
                    status = false;
                    ErrorPer.push(PerArr[i])
                }
            }

            res.json({ success: status, data: ErrorPer })
        }
    });

    // jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    //     if (err) {
    //         res.sendStatus(401);
    //     } else {
    //         const TokenData = await authData;

    //         try {
    //             var request = new sql.Request();

    //             request.input('USERNAME', sql.VarChar(10), req.body.USERNAME)
    //             request.input('REP_ID', sql.Int, parseInt(req.body.REP_ID))
    //             request.input('VIEWFLG', sql.Bit, req.body.VIEWFLG)

    //             request = await request.execute('usp_FrmRepPerMastSave');

    //             res.json({ success: 1, data: '' })

    //         } catch (err) {
    //             res.json({ success: 0, data: err })
    //         }
    //     }
    // });
}

exports.RptPerMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('USER_NAME', sql.VarChar(10), req.body.USER_NAME)
                request.input('REP_ID', sql.Int, parseInt(req.body.REP_ID))

                request = await request.execute('usp_FrmRepPerMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.RptPerMastCopy = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('USERNAME', sql.VarChar(10), req.body.USERNAME)
                request.input('USERS', sql.VarChar, req.body.USERS)

                request = await request.execute('usp_FrmRepPerMastCopy');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

