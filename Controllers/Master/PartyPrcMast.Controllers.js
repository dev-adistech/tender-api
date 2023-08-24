const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.PartyPrcMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('Usp_PartyPrcFill');

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

exports.PartyPrcMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('P_CODE', sql.VarChar(10), req.body.P_CODE)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('PRC_CODE', sql.VarChar(200), req.body.PRC_CODE)
                request.input('S_CODE', sql.VarChar(200), req.body.S_CODE)
                request.input('P_TYPE', sql.VarChar(5), req.body.P_TYPE)
                request.input('FSIZE', sql.Numeric(10,3), req.body.F_SIZE)
                request.input('TSIZE', sql.Numeric(10,3), req.body.T_SIZE)

                request = await request.execute('Usp_PartyPrcSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.PartyPrcMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('P_CODE', sql.VarChar(10), req.body.P_CODE)
                request.input('DEPT_CODE', sql.VarChar(3), req.body.DEPT_CODE)

                request = await request.execute('Usp_PartyPrcDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}
