const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');


exports.TendarMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)

                request = await request.execute('USP_TendarMastFill');

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

exports.TendarMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))

                request = await request.execute('USP_TendarMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.TendarMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('IUSER', sql.VarChar(20), req.body.IUSER)
                request.input('ICOMP', sql.VarChar(30), IP)
                request.input('T_NAME', sql.VarChar(50), req.body.T_NAME)
                request.input('ISACTIVE',sql.Bit,req.body.ISACTIVE)
                request.input('ISMIX',sql.Bit,req.body.ISMIX)
                request.input('ISASSORT',sql.Bit,req.body.ISASSORT)
                request.input('T_CARAT', sql.Numeric(10,3), req.body.T_CARAT)
                request.input('T_PCS', sql.Int, parseInt(req.body.T_PCS))
                request.input('D_DIS', sql.Numeric(10,2), req.body.D_DIS)
                

                request = await request.execute('USP_TendarMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.TendarPktEntFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }

                request = await request.execute('USP_TendarPktEntFill');

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

exports.TendarPktEntSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('I_CARAT', sql.Numeric(10,3), req.body.I_CARAT)
                request.input('COMMENT', sql.VarChar(100), req.body.COMMENT)
                request.input('IUSER', sql.VarChar(20), req.body.IUSER)
                request.input('ICOMP', sql.VarChar(30), IP)
                

                request = await request.execute('USP_TendarPktEntSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.TendarPktEntDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))

                request = await request.execute('USP_TendarPktEntDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.GetTendarNumber = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)

                request = await request.execute('USP_GetTendarNumber');

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

exports.ChkTendarNumber = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))

                request = await request.execute('ufn_ChkTendarNumber');

                if (request.output) {
                    res.json({ success: 1, data: request.output })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.GetTendarPktNumber = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))

                request = await request.execute('USP_GetTendarPktNumber');

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

exports.ChkTendarPktEnt = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))

                request = await request.execute('ufn_ChkTendarPktEnt');

                if (request.output) {
                    res.json({ success: 1, data: request.output })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}