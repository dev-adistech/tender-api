const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');


// [dbo].[USP_AvgSzTypeMastDelete]
// @SZTYPE_CODE Int=0

exports.AvgSzTypeMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SZTYPE_CODE', sql.Int, req.body.SZTYPE_CODE)
                
                request = await request.execute('USP_AvgSzTypeMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}



// [dbo].[USP_AvgSzTypeMastFill]


exports.AvgSzTypeMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('USP_AvgSzTypeMastFill');

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

exports.AvgShpMastSave = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SZTYPE_CODE', sql.Int, req.body.SZTYPE_CODE)
                request.input('SZTYPE_NAME', sql.VarChar(20),req.body.SZTYPE_NAME)
                request.input('ORD', sql.Int, req.body.ORD)
        
                request = await request.execute('USP_AvgSzTypeMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}