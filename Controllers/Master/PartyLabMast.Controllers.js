const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.PartyLabMastFil = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)

                request = await request.execute('USP_PartyLabMastFill');

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

exports.PartyLabOrdCheck = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('P_CODE', sql.VarChar(10), req.body.P_CODE)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('R_ORD', sql.Int, parseInt(req.body.R_ORD))
                request.input('F_SIZE', sql.Decimal(10, 3), parseFloat(req.body.F_SIZE));
                request.input('T_SIZE', sql.Decimal(10, 3), parseFloat(req.body.T_SIZE));
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))

                request = await request.execute('uFn_PartyLabOrdCheck');

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

exports.PartyLabMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('P_CODE', sql.VarChar(50), req.body.P_CODE)
                request.input('DEPT_CODE', sql.VarChar(50), req.body.DEPT_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('S_CODE', sql.VarChar(100), req.body.S_CODE)
                request.input('FC_CODE', sql.Int, parseInt(req.body.FC_CODE))
                request.input('TC_CODE', sql.Int, parseInt(req.body.TC_CODE))
                request.input('F_SIZE', sql.Decimal(10, 3), parseFloat(req.body.F_SIZE));
                request.input('T_SIZE', sql.Decimal(10, 3), parseFloat(req.body.T_SIZE));
                request.input('RATEON', sql.VarChar(2), req.body.RATEON)
                request.input('RATEONC', sql.VarChar(2), req.body.RATEONC)
                request.input('RATE', sql.Decimal(10, 2), parseFloat(req.body.RATE));
                request.input('R_ORD', sql.Int, parseInt(req.body.R_ORD))

                request = await request.execute('USP_PartyLabMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.PartyLabMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))

                request = await request.execute('USP_PartyLabMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}




// exports.PartyLabOrdCheck = async (req, res) => {

//     jwt.verify(req.token, _tokenSecret, async (err, authData) => {
//         if (err) {
//             res.sendStatus(401);
//         } else {
//             const TokenData = await authData;

//             try {

//                 var request = new sql.Request();

//                 request.input('P_CODE', sql.VarChar(10), req.body.P_CODE)
//                 request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
//                 request.input('R_ORD', sql.Int, parseInt(req.body.R_ORD))
//                 request.input('F_SIZE', sql.Decimal(10, 3), parseFloat(req.body.F_SIZE));
//                 request.input('T_SIZE', sql.Decimal(10, 3), parseFloat(req.body.T_SIZE));
//                 request.input('SRNO', sql.Int, parseInt(req.body.SRNO))

//                 request = await request.execute('uFn_PartyLabOrdCheck');
//                 console.log(request);

//                 if (request.recordset) {
//                     res.json({ success: 1, data: request.recordsets })
//                 } else {
//                     res.json({ success: 0, data: "Not Found" })
//                 }

//             } catch (err) {
//                 res.json({ success: 0, data: err })
//             }
//         }
//     });
// }
