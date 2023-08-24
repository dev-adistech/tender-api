var { _tokenSecret } = require('../../Config/token/TokenConfig.json');
const sql = require("mssql");
const jwt = require('jsonwebtoken');

exports.ChrMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)
                request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('CT_CODE', sql.Int, parseInt(req.body.CT_CODE))
                request.input('PO_CODE', sql.Int, parseInt(req.body.PO_CODE))
                request.input('SY_CODE', sql.Int, parseInt(req.body.SY_CODE))
                request.input('CH_CODE', sql.Int, parseInt(req.body.CH_CODE))
                request.input('FQ_CODE', sql.Int, parseInt(req.body.FQ_CODE))
                request.input('TQ_CODE', sql.Int, parseInt(req.body.TQ_CODE))
                request.input('F_CARAT', sql.Numeric(10,3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10,3), req.body.T_CARAT)
                request.input('CH_NAME', sql.VarChar(20), req.body.CH_NAME)
                request.input('RATE', sql.Int, parseInt(req.body.RATE))
                request.input('RTYPE', sql.VarChar(5), req.body.RTYPE)
                request.input('FFP_CARAT', sql.Numeric(10,3), req.body.FFP_CARAT)
                request.input('TFP_CARAT', sql.Numeric(10,3), req.body.TFP_CARAT)
                request.input('FPRATE', sql.Int, parseInt(req.body.FPRATE))
        
                request = await request.execute('USP_ChrMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.CHRMASTDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)
                request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('CT_CODE', sql.Int, parseInt(req.body.CT_CODE))
                request.input('PO_CODE', sql.Int, parseInt(req.body.PO_CODE))
                request.input('SY_CODE', sql.Int, parseInt(req.body.SY_CODE))
                request.input('CH_CODE', sql.Int, parseInt(req.body.CH_CODE))
                request.input('FQ_CODE', sql.Int, parseInt(req.body.FQ_CODE))
                request.input('TQ_CODE', sql.Int, parseInt(req.body.TQ_CODE))
                request.input('F_CARAT', sql.Numeric(10,3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10,3), req.body.T_CARAT)
                request.input('CH_NAME', sql.VarChar(20), req.body.CH_NAME)
                request.input('RATE', sql.Int, parseInt(req.body.RATE))
                request.input('RTYPE', sql.VarChar(5), req.body.RTYPE)
                request.input('FFP_CARAT', sql.Numeric(10,3), req.body.FFP_CARAT)
                request.input('TFP_CARAT', sql.Numeric(10,3), req.body.TFP_CARAT)
                request.input('FPRATE', sql.Int, parseInt(req.body.FPRATE))
        
                request = await request.execute('USP_CHRMASTDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.ChrMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE)
                request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)

                request = await request.execute('USP_ChrMastFill');

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