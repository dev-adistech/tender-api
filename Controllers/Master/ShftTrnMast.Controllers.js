const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.ShiftTrnInsert = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                request.input('S_DATE', sql.DateTime2, new Date(req.body.S_DATE))
                request = await request.execute('usp_ShiftTrnInsert');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.ShiftTrnEmpFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('PRC_CODE', sql.VarChar(5), req.body.PRC_CODE)
                request = await request.execute('usp_ShiftTrnEmpFill');

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

exports.ShiftTrnSaveChk = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('S_DATE', sql.DateTime2, new Date(req.body.S_DATE))
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('MC_CODE', sql.VarChar(5), req.body.MC_CODE)

                request = await request.execute('ufn_FrmShiftTrnSaveChk');

                res.json({ success: 1, data: request.output })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.ShiftTrnSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('S_DATE', sql.DateTime2, new Date(req.body.S_DATE))
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('MC_CODE', sql.VarChar(5), req.body.MC_CODE)
                request.input('PRC_CODE', sql.VarChar(5), req.body.PRC_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)

                request = await request.execute('usp_ShiftTrnSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.ShiftTrnDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('S_DATE', sql.DateTime2, new Date(req.body.S_DATE))
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('MC_CODE', sql.VarChar(5), req.body.MC_CODE)
                request.input('PRC_CODE', sql.VarChar(5), req.body.PRC_CODE)

                request = await request.execute('usp_ShiftTrnDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.ShiftTrnFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                request.input('DEPT_CODE', sql.VarChar(3), req.body.DEPT_CODE)
                request.input('S_DATE', sql.DateTime2, new Date(req.body.S_DATE))
                request.input('S_CODE', sql.VarChar(2), req.body.S_CODE)
                request.input('PRC_CODE', sql.VarChar(5), req.body.PRC_CODE)
                request = await request.execute('usp_ShiftTrnFill');

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

exports.ShiftTrnCmbFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('TYPE', sql.VarChar(10), req.body.TYPE)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('PRC_CODE', sql.VarChar(5), req.body.PRC_CODE)

                request = await request.execute('usp_ShifttrnCmbFill');
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