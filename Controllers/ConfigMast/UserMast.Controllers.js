const sql = require("mssql");
const jwt = require('jsonwebtoken');
var CryptoJS = require("crypto-js");

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.UserMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('USERID', sql.VarChar(20), req.body.USERID)
                request.input('USER_FULLNAME', sql.VarChar(64), req.body.USER_FULLNAME)
                request.input('U_PASS', sql.VarChar(256), CryptoJS.AES.encrypt(req.body.U_PASS, process.env.PASS_KEY).toString())
                request.input('U_CAT', sql.VarChar(8), req.body.U_CAT)
                request.input('IP', sql.VarChar(25), req.body.IP)
                request.input('ISACCESS', sql.Bit, req.body.ISACCESS)
                request.input('EMAIL', sql.VarChar(100), req.body.EMAIL)

                request = await request.execute('USP_UserMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.UserCreate = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('USERID', sql.VarChar(10), req.body.USERID)
                request.input('USER_FULLNAME', sql.VarChar(64), req.body.USER_FULLNAME)
                request.input('U_PASS', sql.VarChar(256), CryptoJS.AES.encrypt(req.body.U_PASS, process.env.PASS_KEY).toString())
                request.input('U_CAT', sql.VarChar(8), req.body.U_CAT)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('PROC_CODE', sql.Int, parseInt(req.body.PROC_CODE))

                request = await request.execute('USP_UserCreate');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.UserMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('USERID', sql.VarChar(20), req.body.USERID)

                request = await request.execute('USP_UserMastFill');

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

exports.UserMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if (req.body.USERID) { request.input('USERID', sql.VarChar(20), req.body.USERID) }

                request = await request.execute('USP_UserMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

