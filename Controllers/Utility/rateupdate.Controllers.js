const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');


exports.TendarRateUpd = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, req.body.DETID)
                request.input('TYPE', sql.VarChar(10), req.body.TYPE)
                
                request = await request.execute('USP_TendarRateUpd');

                res.json({ success: 1, data: request.recordset })
                
            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}