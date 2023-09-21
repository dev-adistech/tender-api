const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');


exports.PerMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            try {
                var request = new sql.Request();

                request.input('UserId', sql.VarChar(20), req.body.UserId)

                request = await request.execute('USP_PerMastFill');

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

exports.PerMastSave = async (req, res) => {

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

                    request.input('USER_NAME', sql.VarChar(10), PerArr[i].USER_NAME)
                    request.input('FORM_NAME', sql.VarChar(32), PerArr[i].FORM_NAME)
                    request.input('INS', sql.Bit, PerArr[i].INS)
                    request.input('DEL', sql.Bit, PerArr[i].DEL)
                    request.input('UPD', sql.Bit, PerArr[i].UPD)
                    request.input('VIW', sql.Bit, PerArr[i].VIW)
                    request.input('PASS', sql.VarChar(10), PerArr[i].PASS)

                    request = await request.execute('USP_PerMastSave');
                } catch (err) {
                    status = false;
                    ErrorPer.push(PerArr[i])
                }
            }

            res.json({ success: status, data: ErrorPer })
        }
    });
}

exports.PerMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('USER_NAME', sql.VarChar(10), req.body.USER_NAME)
                request.input('FORM_NAME', sql.VarChar(32), req.body.FORM_NAME)

                request = await request.execute('USP_PerMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.PerMastCopy = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('USERNAME', sql.VarChar(10), req.body.USERNAME)
                request.input('USERS', sql.VarChar, req.body.USERS)

                request = await request.execute('usp_FrmPerMastCopy');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.RapTrf = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                
                request = await request.execute('USP_RapTrf');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}