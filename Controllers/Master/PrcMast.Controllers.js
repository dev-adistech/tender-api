const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.PrcMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)

                request = await request.execute('USP_PrcMastFill');

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

exports.PrcMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('PRC_CODE', sql.VarChar(3), req.body.PRC_CODE)
                request.input('PRC_NAME', sql.VarChar(15), req.body.PRC_NAME)
                request.input('ORD', sql.Int, parseInt(req.body.ORD))
                request.input('ISACTIVE', sql.Bit, req.body.ISACTIVE)
                request.input('ACFM', sql.Bit, req.body.ACFM)
                request.input('SPRC_CODE', sql.VarChar(3), req.body.SPRC_CODE)
                request.input('PRVPRC', sql.VarChar(100), req.body.PRVPRC)
                request.input('INWD', sql.Bit, req.body.INWD)
                request.input('CFM', sql.Bit, req.body.CFM)
                request.input('RCV', sql.Bit, req.body.RCV)
                request.input('RCVINWD', sql.Bit, req.body.RCVINWD)
                request.input('RTN', sql.Bit, req.body.RTN)
                request.input('PRTN', sql.Bit, req.body.PRTN)
                request.input('PRTNINWD', sql.Bit, req.body.PRTNINWD)
                request.input('BRK', sql.Bit, req.body.BRK)
                request.input('INNER_PRC', sql.Bit, req.body.INNER_PRC)
                request.input('EPRC', sql.Bit, req.body.EPRC)
                request.input('LPER', sql.Int, parseInt(req.body.LPER))
                request.input('LHOUR', sql.Int, parseInt(req.body.LHOUR))
                request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE)
                request.input('SDEPT_CODE', sql.VarChar(50), req.body.SDEPT_CODE)
                request.input('ISCLV', sql.Bit, req.body.ISCLV)
                request.input('RPRC', sql.VarChar(100), req.body.RPRC)
                request.input('TRFPRC', sql.VarChar(50), req.body.TRFPRC)
                request.input('ISMRK', sql.Bit, req.body.ISMRK)

                request = await request.execute('USP_PrcMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.PrcMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('PRC_CODE', sql.VarChar(3), req.body.PRC_CODE)

                request = await request.execute('USP_PrcMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}
