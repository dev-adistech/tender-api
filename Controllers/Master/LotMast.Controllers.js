const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');


exports.LotMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('IUSER', sql.VarChar(10), req.body.IUSER);
                request.input("PNT", sql.Int, parseInt(req.body.PNT))

                request = await request.execute('USP_LotMastFill');

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

exports.LotMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(5), req.body.L_CODE);
                request.input('TRNNO', sql.Int, parseInt(req.body.TRNNO));
                request.input('R_CODE', sql.VarChar(5), req.body.R_CODE);
                request.input('L_NAME', sql.VarChar(10), req.body.L_NAME);
                request.input('L_DATE', sql.DateTime2, new Date(req.body.L_DATE));
                request.input('L_PCS', sql.Int, parseInt(req.body.L_PCS));
                request.input('L_CARAT', sql.Decimal(10, 3), parseFloat(req.body.L_CARAT));
                request.input('SIZE', sql.Decimal(10, 3), parseFloat(req.body.SIZE));
                request.input('C_DATE', sql.DateTime2, new Date(req.body.C_DATE));
                request.input('PNT', sql.Int, parseInt(req.body.PNT));
                request.input('GRP', sql.VarChar(5), req.body.GRP);
                request.input('LOVER', sql.Bit, req.body.LOVER);
                request.input('RAP_TYPE', sql.VarChar(5), req.body.RAP_TYPE);
                request.input('IS_URGENT', sql.Bit, req.body.IS_URGENT);
                request.input('LOCK_TYPE', sql.VarChar(1), req.body.LOCK_TYPE);
                if (req.body.M_DATE) { request.input('M_DATE', sql.DateTime2, new Date(req.body.M_DATE)) };
                request.input('MDAY', sql.Int, parseInt(req.body.MDAY));
                if (req.body.P_DATE) { request.input('P_DATE', sql.DateTime2, new Date(req.body.P_DATE)) };
                request.input('PDAY', sql.Int, parseInt(req.body.PDAY));
                request.input('EPER', sql.Decimal(10, 2), parseFloat(req.body.EPER));
                request.input('VWORD', sql.Int, parseInt(req.body.VWORD));
                request.input('PRCORD', sql.Int, parseInt(req.body.PRCORD));
                request.input('IPRCORD', sql.Int, parseInt(req.body.IPRCORD));
                request.input('ISPRD', sql.Int, parseInt(req.body.ISPRD));
                request.input('ISDIR', sql.Bit, req.body.ISDIR);
                request.input('COVER', sql.Bit, req.body.COVER);


                request = await request.execute('USP_LotMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }

    });
}

exports.LotMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(5), req.body.L_CODE)

                request = await request.execute('USP_LotMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.LotTRf = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input("L_CODE", sql.VarChar(7991), req.body.L_CODE)
                request.input('TYP', sql.VarChar(50), req.body.TYP)

                request = await request.execute('Usp_LotTRf')

                // console.log(req.body)
                // console.log(request)
                res.json({ success: true, data: '' })

            } catch (err) {
                console.log(err)
                res.json({ success: false, data: err })
            }
        }
    })
}

exports.LotTRfLotFill = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('TYP', sql.VarChar(50), req.body.TYP)

                request = await request.execute('Usp_LotTRfLotFill')

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 0, data: 'Not Found' })
                }

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    })
}

exports.LotMastChart = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('SRNO', sql.Int, req.body.SRNO)
                request.input('TAG', sql.VarChar(3), req.body.TAG)
                request.input('Type', sql.VarChar(10), req.body.Type)

                request = await request.execute('Usp_LotProcessChart')

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 0, data: 'Not Found' })
                }

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    })
}
