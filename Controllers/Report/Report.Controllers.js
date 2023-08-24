const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.ReportPrint = async (req, res) => {

    try {
        res.render('index', { 'data': JSON.parse(req.body.Data), 'mrtname': req.body.mrtname });
    } catch (err) {

        res.json({ success: 0, data: err })
    }

}

exports.ReportViewer = async (req, res) => {
    // console.log(req.body.mrtname)
    try {
        res.render('report', { 'data': JSON.parse(req.body.Data), 'mrtname': req.body.mrtname });
    } catch (err) {

        res.json({ success: 0, data: err })
    }

}

// function groupByArray(xs, PRC_NAME, L_CODE, PEMP_CODE) {
//     return xs.reduce(function (rv, x) {
//         let _PRC_NAME = PRC_NAME instanceof Function ? PRC_NAME(x) : x[PRC_NAME];
//         let _L_CODE = L_CODE instanceof Function ? L_CODE(x) : x[L_CODE];

//         let el = rv.find(r => r && r.PRC_NAME === _PRC_NAME);

//         if (el) {
//             el.Data.push(x);
//         } else {
//             rv.push({
//                 PRC_NAME: _PRC_NAME,
//                 Data: [x]
//             });
//         }

//         return rv;
//     }, []);
// }

exports.FetchReportData = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('TABLENAME', sql.VarChar(100), req.body.TABLENAME)
                request.input('FIELDS', sql.VarChar(8000), req.body.FIELDS)
                request.input('CRITERIA', sql.VarChar(8000), req.body.CRITERIA)

                request = await request.execute('USP_SELECT');

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

exports.FillReportList = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('USER_NAME', sql.VarChar(8000), req.body.USER_NAME)
                request.input('PARENT_ID', sql.Int, parseInt(req.body.PARENT_ID))

                request = await request.execute('usp_FrmReportFill');

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

exports.GetReportParams = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('SP_NAME', sql.VarChar(50), req.body.SP_NAME)
                request.input('REP_ID', sql.Int, parseInt(req.body.REP_ID))

                request = await request.execute('USP_GetRepSPParam');

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

exports.RPTDATA = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {

            const TokenData = await authData;

            let RAPDATAArr = req.body;

            let status = true;
            let ErrorArr = []
            let TempArr = []
            var request = new sql.Request();
            for (let i = 0; i < RAPDATAArr.SP_PARAMS.length; i++) {
                try {

                    if (RAPDATAArr.SP_PARAMS[i].DATATYPE == 'varchar') {
                        if (RAPDATAArr.SP_PARAMS[i].VALUE) {
                            request.input(RAPDATAArr.SP_PARAMS[i].FIELD, sql.VarChar, RAPDATAArr.SP_PARAMS[i].VALUE)
                            // TempArr.push(RAPDATAArr.SP_PARAMS[i])
                        }
                    } else if (RAPDATAArr.SP_PARAMS[i].DATATYPE == 'int') {
                        if (RAPDATAArr.SP_PARAMS[i].VALUE) {
                            request.input(RAPDATAArr.SP_PARAMS[i].FIELD, sql.Int, parseInt(RAPDATAArr.SP_PARAMS[i].VALUE))
                            // TempArr.push(RAPDATAArr.SP_PARAMS[i])
                        }
                    } else if (RAPDATAArr.SP_PARAMS[i].DATATYPE == 'numeric') {
                        if (RAPDATAArr.SP_PARAMS[i].VALUE) {
                            request.input(RAPDATAArr.SP_PARAMS[i].FIELD, sql.Decimal(10, 3), parseFloat(RAPDATAArr.SP_PARAMS[i].VALUE))
                            // TempArr.push(RAPDATAArr.SP_PARAMS[i])
                        }
                    } else if (RAPDATAArr.SP_PARAMS[i].DATATYPE == 'datetime') {
                        if (RAPDATAArr.SP_PARAMS[i].VALUE) {
                            request.input(RAPDATAArr.SP_PARAMS[i].FIELD, sql.DateTime2, new Date(RAPDATAArr.SP_PARAMS[i].VALUE))
                            // TempArr.push(RAPDATAArr.SP_PARAMS[i])
                        }
                    } else if (RAPDATAArr.SP_PARAMS[i].DATATYPE == 'date') {
                        if (RAPDATAArr.SP_PARAMS[i].VALUE) {
                            request.input(RAPDATAArr.SP_PARAMS[i].FIELD, sql.DateTime2, new Date(RAPDATAArr.SP_PARAMS[i].VALUE))
                            // TempArr.push(RAPDATAArr.SP_PARAMS[i])
                        }
                    } else if (RAPDATAArr.SP_PARAMS[i].DATATYPE == 'time') {
                        if (RAPDATAArr.SP_PARAMS[i].VALUE) {
                            request.input(RAPDATAArr.SP_PARAMS[i].FIELD, sql.Time, new Date(RAPDATAArr.SP_PARAMS[i].VALUE))
                            // TempArr.push(RAPDATAArr.SP_PARAMS[i])
                        }
                    }
                    // console.log(TempArr);
                } catch (err) {
                    status = false;
                    ErrorArr.push(RAPDATAArr[i])
                    s.log('ErrorArr', ErrorArr);
                }
            }

            request = await request.execute(RAPDATAArr.SP_DETAILS[0].SP_NAME);

            if (request.recordset) {
                res.json({ success: 1, data: request.recordsets })
            } else {
                res.json({ success: 0, data: "Not Found" })
            }

        }
    });
}

exports.FrmRptDataFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('USP_FrmRptDataFill');

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

