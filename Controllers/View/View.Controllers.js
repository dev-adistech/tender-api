const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.PricingWrk = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if(req.body.F_DATE){request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE))}
                if(req.body.T_DATE){request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE))}
                request.input('S_CODE', sql.VarChar(7991), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(7991), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(7991), req.body.Q_CODE)
                request.input('F_CARAT', sql.Numeric(10,3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10,3), req.body.T_CARAT)
                
                request = await request.execute('VW_PricingWrk');

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

exports.ParcelgWrk = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if(req.body.F_DATE){request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE))}
                if(req.body.T_DATE){request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE))}
                request.input('S_CODE', sql.VarChar(7991), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(7991), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(7991), req.body.Q_CODE)
                request.input('F_CARAT', sql.Numeric(10,3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10,3), req.body.T_CARAT)
                
                request = await request.execute('VW_ParcelgWrk');

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

exports.PricingWrkDisp = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if(req.body.F_DATE){request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE))}
                if(req.body.T_DATE){request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE))}
                request.input('S_CODE', sql.VarChar(7991), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(7991), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(7991), req.body.Q_CODE)
                request.input('F_CARAT', sql.Numeric(10,3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10,3), req.body.T_CARAT)
                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)

                
                request = await request.execute('VW_PricingWrkDisp');

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

exports.ParcelWrkDisp = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if(req.body.F_DATE){request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE))}
                if(req.body.T_DATE){request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE))}
                request.input('S_CODE', sql.VarChar(7991), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(7991), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(7991), req.body.Q_CODE)
                request.input('F_CARAT', sql.Numeric(10,3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10,3), req.body.T_CARAT)
                if(req.body.COMP_CODE){request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)}

                
                request = await request.execute('VW_ParcelWrkDisp');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordsets })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.PricingWrkMperSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('MPER', sql.Numeric(10,3), req.body.MPER)
                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('PLANNO', sql.Int, parseInt(req.body.PLANNO))
                request.input('PTAG', sql.VarChar(2), req.body.PTAG)

                
                request = await request.execute('VW_PricingWrkMperSave');

        
                    res.json({ success: 1, data: '' })
            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.BVView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                
                request = await request.execute('VW_BVView');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordsets })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.BidDataView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                
                request = await request.execute('VW_BidDataView');

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

exports.ColAnalysis = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                request.input('S_CODE', sql.VarChar(7991), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(7991), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(7991), req.body.Q_CODE)
                request.input('CT_CODE', sql.VarChar(7991), req.body.CT_CODE)
                request.input('F_CARAT', sql.Numeric(10,3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10,3), req.body.T_CARAT)
                request.input('FINAL', sql.VarChar(7991), req.body.FINAL)
                request.input('RESULT', sql.VarChar(7991), req.body.RESULT)
                request.input('FLNO', sql.VarChar(7991), req.body.FLNO)
                request.input('MACHFL', sql.VarChar(7991), req.body.MACHFL)
                request.input('COMENT', sql.VarChar(7991), req.body.COMENT)
                request.input('LAB', sql.VarChar(10), req.body.LAB)
                
                request = await request.execute('VW_ColAnalysis');

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