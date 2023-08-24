const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.FrmMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();


                request = await request.execute('USP_FrmMastFill');

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

exports.FrmMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            let FrmArr = req.body;
            let status = true;
            let ErrorFrm = []
            for (let i = 0; i < FrmArr.length; i++) {
                try {
                    var request = new sql.Request();

                    request.input('FORM_GROUP', sql.VarChar(64), FrmArr[i].FORM_GROUP)
                    request.input('FORM_NAME', sql.VarChar(50), FrmArr[i].FORM_NAME)
                    request.input('DESCR', sql.VarChar(64), FrmArr[i].DESCR)

                    request = await request.execute('USP_FrmMastSave');
                } catch (err) {
                    status = false;
                    ErrorFrm.push(FrmArr[i])
                }
            }

            res.json({ success: status, data: ErrorFrm })
        }
    });
}





exports.FrmMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('FORM_ID', sql.Int, parseInt(req.body.FORM_ID))

                request = await request.execute('USP_FrmMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}