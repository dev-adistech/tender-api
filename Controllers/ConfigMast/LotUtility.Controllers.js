const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.LotChange = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('NL_CODE', sql.VarChar(10), req.body.NL_CODE)

                request = await request.execute('USP_LotChange');

                res.json({ success: 1, data: request.recordset })

            } catch (err) {

                res.json({ success: 0, data: err })
            }
        }
    });
}


exports.RateUpdateUtility = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('PRDTYPE', sql.VarChar(5), req.body.PRDTYPE)

                request = await request.execute('usp_RateUpdateUtility');

                res.json({ success: 1, data: request.recordset })

            } catch (err) {

                res.json({ success: 0, data: err })
            }
        }
    });
}