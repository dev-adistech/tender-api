const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.LoginDetailFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('USERID', sql.VarChar(10), req.body.USERID)
                request.input('RUN', sql.VarChar(5), req.body.RUN)

                request = await request.execute('usp_UserLogDetFill');

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

exports.LoginDetailUpdate = async (req, res) => {

    try {
        var request = new sql.Request();
        let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;

        // console.log(req.body)
        request.input('USERID', sql.VarChar(10), req.body.USERID)
        if (req.body.TYPE === "GRID") {
            request.input('IP', sql.VarChar(50), req.body.IP)
        } else if (req.body.TYPE === 'Logout') {
            request.input('IP', sql.VarChar(50), IP)
        }
        request = await request.execute('usp_UserLogDetUpd');

        res.json({ success: 1, data: '' })

    } catch (err) {
        res.json({ success: 0, data: err })
    }
}

exports.LoginDetailSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {

                var request = new sql.Request();

                let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;

                request.input('USERID', sql.VarChar(10), req.body.USERID)
                request.input('IP', sql.VarChar(50), IP)
                request.input('RUN', sql.Bit, req.body.RUN)
                request.input('URL', sql.VarChar(200), req.body.URL)

                request = await request.execute('usp_UserLogDetSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}
