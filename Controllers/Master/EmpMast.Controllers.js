const sql = require("mssql");
const jwt = require('jsonwebtoken');
var CryptoJS = require("crypto-js");

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.EmpMastFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('PROC_CODE', sql.Int, parseInt(req.body.PROC_CODE))
                request.input("PNT", sql.Int, parseInt(req.body.PNT))

                request = await request.execute('USP_EmpMastFill');

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

exports.EmpMastSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('PROC_CODE', sql.Int, parseInt(req.body.PROC_CODE))
                request.input('OLDPROC_CODE', sql.Int, parseInt(req.body.OLDPROC_CODE))
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('EMP_NAME', sql.VarChar(50), req.body.EMP_NAME)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('USERID', sql.VarChar(10), req.body.USERID)
                request.input('GRD', sql.VarChar(5), req.body.GRD)
                request.input('CEMP_CODE', sql.VarChar(5), req.body.CEMP_CODE)
                request.input('OEMP_CODE', sql.VarChar(5), req.body.OEMP_CODE)
                request.input('M_CODE', sql.VarChar(50), req.body.M_CODE)
                request.input('SALARY', sql.Int, parseInt(req.body.SALARY))
                if (req.body.J_DATE) { request.input('J_DATE', sql.DateTime2, new Date(req.body.J_DATE)) }
                if (req.body.L_DATE) { request.input('L_DATE', sql.DateTime2, new Date(req.body.L_DATE)) }
                request.input('INNER_PRC', sql.VarChar(200), req.body.INNER_PRC)
                request.input('SFLG', sql.VarChar(20), req.body.SFLG)
                request.input('ADD1', sql.VarChar(50), req.body.ADD1)
                request.input('ADD2', sql.VarChar(50), req.body.ADD2)
                request.input('ADD3', sql.VarChar(50), req.body.ADD3)
                request.input('ADD4', sql.VarChar(50), req.body.ADD4)
                request.input('ADD5', sql.VarChar(50), req.body.ADD5)
                request.input('LABID', sql.VarChar(10), req.body.LABID)
                request.input('ISDASH', sql.Bit, req.body.ISDASH)
                request.input('ISUPD', sql.Bit, req.body.ISUPD)
                request.input('OPKT', sql.Int, parseInt(req.body.OPKT))
                request.input('LPKT', sql.Int, parseInt(req.body.LPKT))
                request.input('TYPE', sql.VarChar(5), req.body.TYPE)
                request.input('MEMP_CODE', sql.VarChar(100), req.body.MEMP_CODE)
                request.input('PDAY', sql.Int, parseInt(req.body.PDAY))
                request.input('TRPID', sql.Int, parseInt(req.body.TRPID))
                request.input('TAB', sql.VarChar(10), req.body.TAB)

                request = await request.execute('USP_EmpMastSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.EmpMastDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if (req.body.DEPT_CODE) { request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE) }
                if (req.body.PROC_CODE) { request.input('PROC_CODE', sql.Int, parseInt(req.body.PROC_CODE)) }
                if (req.body.EMP_CODE) { request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE) }

                request = await request.execute('USP_EmpMastDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.USERMASTSAVETEMP = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            let RapArr = req.body;
            // console.log(RapArr)
            let status = true;
            let ErrorRap = []
            for (let i = 0; i < RapArr.length; i++) {
                try {
                    var request = new sql.Request();

                    request.input('USERID', sql.VarChar(10), RapArr[i].USERID)
                    request.input('USER_FULLNAME', sql.VarChar(64), RapArr[i].USER_FULLNAME)
                    request.input('U_PASS', sql.VarChar(256), CryptoJS.AES.encrypt(RapArr[i].U_PASS, process.env.PASS_KEY).toString())
                    request.input('U_CAT', sql.VarChar(8), 'U')
                    request.input('PNT', sql.Int, parseInt(RapArr[i].PNT))
                    request.input('DEPT_CODE', sql.VarChar(5), RapArr[i].DEPT_CODE)
                    request.input('PROC_CODE', sql.Int, parseInt(RapArr[i].PROC_CODE))

                    request = await request.execute('USP_UserMastSave');
                } catch (err) {
                    console.log(err)
                    status = false;
                    ErrorRap.push(RapArr[i])
                }
            }
            res.json({ success: status, data: ErrorRap })

        }
    });
}

exports.GrpUpdate = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('usp_GrpUpdate');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.EmpMastChart = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)

                request = await request.execute('USP_EmpMastChart');

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
