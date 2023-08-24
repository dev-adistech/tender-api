const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

var { RapServerConn } = require('../../config/db/rapsql');

exports.RapOrgFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                let request = new sql.Request()
                request.input('SZ_CODE', sql.Int, parseInt(req.body.SZ_Code))
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)

                request = await request.execute('USP_FrmRapOrgFill');

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

exports.RapOrgSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                let request = new sql.Request()

                request.input('C_CODE', sql.Int, parseInt(req.body.C_CODE))
                request.input('SZ_CODE', sql.Int, parseInt(req.body.SZ_CODE))
                request.input('S_CODE', sql.VarChar(3), req.body.S_CODE)
                request.input('Q1', sql.Decimal(10, 2), parseFloat(req.body.Q1))
                request.input('Q2', sql.Decimal(10, 2), parseFloat(req.body.Q2))
                request.input('Q3', sql.Decimal(10, 2), parseFloat(req.body.Q3))
                request.input('Q4', sql.Decimal(10, 2), parseFloat(req.body.Q4))
                request.input('Q5', sql.Decimal(10, 2), parseFloat(req.body.Q5))
                request.input('Q6', sql.Decimal(10, 2), parseFloat(req.body.Q6))
                request.input('Q7', sql.Decimal(10, 2), parseFloat(req.body.Q7))
                request.input('Q8', sql.Decimal(10, 2), parseFloat(req.body.Q8))
                request.input('Q9', sql.Decimal(10, 2), parseFloat(req.body.Q9))
                request.input('Q10', sql.Decimal(10, 2), parseFloat(req.body.Q10))
                request.input('Q11', sql.Decimal(10, 2), parseFloat(req.body.Q11))
                request.input('Q12', sql.Decimal(10, 2), parseFloat(req.body.Q12))
                request.input('Q13', sql.Decimal(10, 2), parseFloat(req.body.Q13))
                request.input('Q14', sql.Decimal(10, 2), parseFloat(req.body.Q14))
                request.input('Q15', sql.Decimal(10, 2), parseFloat(req.body.Q15))
                request.input('Q16', sql.Decimal(10, 2), parseFloat(req.body.Q16))
                request.input('Q17', sql.Decimal(10, 2), parseFloat(req.body.Q17))
                request.input('F_CARAT', sql.Decimal(10, 2), parseFloat(req.body.F_CARAT))
                request.input('T_CARAT', sql.Decimal(10, 2), parseFloat(req.body.T_CARAT))

                request = await request.execute('USP_FrmRapOrgSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.RapDownload = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {

        if (err) {
            res.sendStatus(401);
        } else {
            try {
                const TokenData = await authData;
                let AuthOptions = {
                    method: 'POST',
                    headers: {
                        "Content-Type": "application/json"
                    },
                    formData: {
                        "Username": req.body.Username,
                        "Password": req.body.Password
                    }
                };
                let AuthRes = await CallExtApi("https://technet.rapaport.com/HTTP/Authenticate.aspx", AuthOptions)

                let DownloadOptions = {
                    method: 'POST',
                    headers: {
                        "Content-Type": "Application/X-Www-Form-Urlencoded",
                        "Accept-Encoding": "gzip"
                    }
                }

                let DownloadCsvRound = await CallExtApi("https://technet.rapaport.com/HTTP/Prices/CSV2_Round.aspx?ticket=" + AuthRes, DownloadOptions)
                let DownloadCsvPear = await CallExtApi("https://technet.rapaport.com/HTTP/Prices/CSV2_Pear.aspx?ticket=" + AuthRes, DownloadOptions)
                res.json({ success: 1, DataRound: DownloadCsvRound, DataPear: DownloadCsvPear })
            } catch (err) {
                res.json({ success: 0, data: "" })
            }
        }
    })
}

function CallExtApi(url, options) {
    return new Promise(function (resolve, reject) {
        _request(url, options, (error, res, body) => {
            if (!error && res.statusCode == 200) {
                resolve(body);
            } else {
                reject(error);
            }
        })
    });
}

exports.RapTrf = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {

                let request = new sql.Request()

                request = await request.execute('USP_RapTrf');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

