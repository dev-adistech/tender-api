const sql = require("mssql");
const jwt = require('jsonwebtoken');
const http = require('http');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');
const { captureRejectionSymbol } = require("events");

exports.FillAllMaster = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request = await request.execute('USP_FillAllMaster');

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

exports.DashStockFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(10), req.body.GRP)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)

                request = await request.execute('usp_DashStock');

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

exports.DashReceiveFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(10), req.body.GRP)

                request = await request.execute('usp_DashReceive');

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

exports.DashReceiveStockConfirm = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            let PerArr = req.body;
            let status = true;
            let ErrorPer = []
            for (let i = 0; i < PerArr.length; i++) {
                try {
                    var request = new sql.Request();

                    request.input('L_CODE', sql.VarChar(10), PerArr[i].L_CODE)
                    request.input('SRNO', sql.Int, parseInt(PerArr[i].SRNO))
                    request.input('TAG', sql.VarChar(5), PerArr[i].TAG)
                    request.input('DETID', sql.Int, parseInt(PerArr[i].DETID))
                    request.input('R_DATE', sql.DateTime2, new Date(PerArr[i].R_DATE))
                    request.input('TPROC_CODE', sql.VarChar(10), PerArr[i].TPROC_CODE)
                    request.input('RNO', sql.Int, parseInt(PerArr[i].RNO))
                    request.input('EMP_CODE', sql.VarChar(5), PerArr[i].EMP_CODE)
                    request.input('STCOMP', sql.VarChar(20), PerArr[i].STCOMP)
                    request.input('PNT', sql.Int, parseInt(PerArr[i].PNT))

                    request = await request.execute('usp_DashRecStockConf');
                } catch (err) {
                    status = false;
                    ErrorPer.push(PerArr[i])
                }
            }

            res.json({ success: status, data: ErrorPer })
        }
    });
}

exports.DashCurrentRollingFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(10), req.body.GRP)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)

                request = await request.execute('usp_DashCurRolling');

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

exports.DashTodayMakeableFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(10), req.body.GRP)

                request = await request.execute('usp_DashTodayMak');

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

exports.DashMakeableViewFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(10), req.body.GRP)

                request = await request.execute('USP_DashMakableView');

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

exports.DashPredictionPendingFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(10), req.body.GRP)

                request = await request.execute('usp_DashPrdPen');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DashPredictionPendingFileDownload = async (req, res) => {
    const ftp = require("basic-ftp")
    const client = new ftp.Client()

    var FTPrequest = new sql.Request();

    FTPrequest.input('L_CODE', sql.VarChar(10), '')
    FTPrequest.input('SRNO', sql.Int, 0)
    FTPrequest.input('TAG', sql.VarChar(10), '')
    FTPrequest.input('I_DATE', sql.DateTime2, new Date())
    FTPrequest.input('TYPE', sql.VarChar(10), 'QCOK')

    FTPrequest = await FTPrequest.execute('usp_StnFilePath');

    const PORT = process.env.MAIN_PORT
    const DOMAIN_IP = process.env.DOMAIN_IP

    const FILENAME = `${req.body.FILENAME}.cap`
    const URL = `http://${DOMAIN_IP}:${PORT}/file/QCFILE/${FILENAME}`

    // client.ftp.verbose = true
    // console.log(`FTP HOST: ${FTPrequest.recordset[0].HOST}`)
    // console.log(`FTP USER: ${FTPrequest.recordset[0].USER}`)
    // console.log(`FTP PASS: ${FTPrequest.recordset[0].PASS}`)
    // console.log(`FTP PORT: ${FTPrequest.recordset[0].PORT}`)

    await client.access({
        host: FTPrequest.recordset[0].HOST,
        user: FTPrequest.recordset[0].USER,
        password: FTPrequest.recordset[0].PASS,
        port: FTPrequest.recordset[0].PORT,
        secure: false
    })

    await client.downloadTo(`./Public/file/QCFILE/${FILENAME}`, `GALAXY/${FILENAME}`)
    await console.log(`${FILENAME} downloaded successfully.`)

    http.get(URL, function (response) {
        var data = []
        response.on('data', function (chunk) {
            data.push(chunk);
        });
        response.on('end', function () {
            data = Buffer.concat(data);

            res.writeHead(200, {
                'Content-Length': Buffer.byteLength(data),
                'Content-Type': 'application/octet-stream',
                'Content-disposition': 'attachment;filename=' + FILENAME
            }).end(data);
        });
        data = []
    });

    const fs = require('fs');
    try {
        fs.unlinkSync(`./Public/file/QCFILE/${FILENAME}`);
        // console.log(`${FILENAME} is deleted.`);
    } catch (error) {
        console.log(error)
    }
}

exports.DashTop20MarkerFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(10), req.body.GRP)
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)
                if (req.body.SDATE) { request.input('SDATE', sql.VarChar(10), req.body.SDATE) }
                if (req.body.TDATE) request.input('TDATE', sql.VarChar(10), req.body.TDATE)
                // console.log(req.body)
                request = await request.execute('usp_DashTop20');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                console.log('----------------------------', err)
                res.json({ success: 0, data: err })
            }
        }
    });
}


exports.DashGrdView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(5), req.body.GRP)

                request = await request.execute('usp_DashGrdView');

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

exports.DashMrkAccView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                // request.input('SDATE', sql.DateTime2, req.body.SDATE)
                // request.input('TDATE', sql.DateTime2, req.body.TDATE)
                // request.input('SDATE', sql.DateTime2, new Date(req.body.SDATE))
                // request.input('TDATE', sql.DateTime2, new Date(req.body.TDATE))
                if (req.body.SDATE) { request.input('SDATE', sql.DateTime2, new Date(req.body.SDATE)) }
                if (req.body.TDATE) { request.input('TDATE', sql.DateTime2, new Date(req.body.TDATE)) }
                request.input('PNT', sql.Int, parseInt(req.body.PNT))

                request = await request.execute('usp_DashMrkAccView');

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

exports.DashMrkAccViewDet = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7999), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('PRDTYPE', sql.VarChar(5), req.body.PRDTYPE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(5), req.body.GRP)

                request = await request.execute('usp_DashMrkAccViewDet');

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

exports.DashTop20Det = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(7991), req.body.EMP_CODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(10), req.body.GRP)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                if (req.body.SDATE !== "") { request.input('SDATE', sql.DateTime, new Date(req.body.SDATE)) }
                if (req.body.TDATE !== "") { request.input('TDATE', sql.DateTime, new Date(req.body.TDATE)) }
                request.input('COLNAME', sql.VarChar(10), req.body.COLNAME)
                request.input('TYP', sql.VarChar(10), req.body.TYP)
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)

                request = await request.execute('usp_DashTop20Det');
                // console.log(request)
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

exports.DashMrkPrdView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('MONTH', sql.VarChar(2), req.body.MONTH)
                request.input('YEAR', sql.VarChar(4), req.body.YEAR)
                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('SDATE', sql.VarChar(10), req.body.SDATE)
                request.input('TDATE', sql.VarChar(10), req.body.TDATE)
                request.input('GRP', sql.VarChar(10), req.body.GRP)

                request = await request.execute('Usp_DashMrkPrdView');
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    });

}

exports.DashMrkPrdViewDet = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('PROC_CODE', sql.VarChar(10), req.body.PROC_CODE)
                request.input('SDATE', sql.VarChar(10), req.body.SDATE)
                request.input('TDATE', sql.VarChar(10), req.body.TDATE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('DATE', sql.VarChar(5), req.body.DATE)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))

                request = await request.execute('Usp_DashMrkPrdViewDet');
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

exports.DashAttence = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(7991), req.body.EMP_CODE)
                if (req.body.SDATE) { request.input('SDATE', sql.VarChar(10), req.body.SDATE) }
                if (req.body.TDATE) { request.input('TDATE', sql.VarChar(10), req.body.TDATE) }
                request.input('GRP', sql.VarChar(7991), req.body.GRP)
                request.input('PNT', sql.Int, req.body.PNT)

                request = await request.execute('Usp_DashAttence');
                if (request.recordsets) {
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

exports.DashAttenceDet = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('TRPID', sql.VarChar(10), req.body.TRPID)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('DATE', sql.VarChar(5), req.body.DATE)
                request.input('PNT', sql.Int, req.body.PNT)


                request = await request.execute('Usp_DashAttencDet');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DashMrkAccDoubleClick = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7999), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('PRDTYPE', sql.VarChar(5), req.body.PRDTYPE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)

                request = await request.execute('usp_DashMrkAccDoubleClick');
                if (request.recordsets) {
                    res.json({ success: 1, data: request.recordsets })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    })
}

exports.DashMrkAccPrdSave = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(3), req.body.TAG)
                request.input('PRDTYPE', sql.VarChar(5), req.body.PRDTYPE)
                request.input('PLANNO', sql.Int, parseInt(req.body.PLANNO))
                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('PTAG', sql.VarChar(5), req.body.PTAG)
                request.input('I_CARAT', sql.Decimal(10, 3), req.body.I_CARAT)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('CARAT', sql.Decimal(10, 3), req.body.CARAT)
                request.input('C_CODE', sql.Int, parseInt(req.body.C_CODE))
                request.input('IN_CODE', sql.Int, parseInt(req.body.IN_CODE))
                request.input('Q_CODE', sql.Int, parseInt(req.body.Q_CODE))
                request.input('CUT_CODE', sql.Int, parseInt(req.body.CUT_CODE))
                request.input('POL_CODE', sql.Int, parseInt(req.body.POL_CODE))
                request.input('SYM_CODE', sql.Int, parseInt(req.body.SYM_CODE))
                request.input('FL_CODE', sql.Int, parseInt(req.body.FL_CODE))
                request.input('SH_CODE', sql.Int, parseInt(req.body.SH_CODE))
                request.input('MILKY_CODE', sql.Int, parseInt(req.body.MILKY_CODE))
                request.input('REDSPOT', sql.Int, parseInt(req.body.REDSPORT))
                request.input('RAPTYPE', sql.VarChar(5), req.body.RAPTYPE)
                request.input('RATE', sql.Int, parseInt(req.body.RATE))
                request.input('IUSER', sql.VarChar(10), req.body.IUSRE)
                request.input('ICOMP', sql.VarChar(25), req.body.ICOMP)

                console.log(req.body)

                request = await request.execute('usp_DashMrkAccPrdSave');

                res.json({ success: 1, data: '' })
            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }

        }
    })
}

exports.DashMrkAccDobDel = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(5), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(2), req.body.TAG)
                request.input('PRDTYPE', sql.VarChar(5), req.body.PRDTYPE)
                request.input('PLANNO', sql.Int, parseInt(req.body.PLANNO))
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('PTAG', sql.VarChar(2), req.body.PTAG)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('ICOMP', sql.VarChar(30), req.body.ICOMP)

                request = await request.execute('usp_MrkAccPlanDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.MrkAccDel = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7999), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)

                request = await request.execute('usp_MrkAccDel')

                res.json({ success: 1, data: '' })
            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    })
}

exports.DashPrdPenUpd = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData;
            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(5), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('ISPRD', sql.Bit, req.body.ISPRD)

                // console.log(req.body)
                request = await request.execute('usp_DashPrdPenUpd')
                // console.log(request)
                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.DashOTPrdPen = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input("EMP_CODE", sql.VarChar(7991), req.body.EMP_CODE)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))
                request.input("GRP", sql.VarChar(10), req.body.GRP)

                request = await request.execute('usp_DashOTPrdPen')

                res.json({ success: 1, data: request.recordset })
            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.DashDailyPrc = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input("PRC_CODE", sql.VarChar(5), req.body.PRC_CODE)
                request.input("F_DATE", sql.VarChar(20), req.body.F_DATE)
                request.input("T_DATE", sql.VarChar(20), req.body.T_DATE)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))

                request = await request.execute('usp_DashDailyProcess')

                res.json({ success: 1, data: request.recordset })
            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.DashSarin = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input("EMP_CODE", sql.VarChar(7991), req.body.EMP_CODE)
                if (req.body.SDATE) { request.input('SDATE', sql.DateTime2, new Date(req.body.SDATE)) }
                if (req.body.TDATE) { request.input('TDATE', sql.DateTime2, new Date(req.body.TDATE)) }
                request.input("GRP", sql.VarChar(10), req.body.GRP)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))

                request = await request.execute('Usp_dashSarin')

                res.json({ success: 1, data: request.recordset })
            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.DashIssStock = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input("EMP_CODE", sql.VarChar(10), req.body.EMP_CODE)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))
                request.input("GRP", sql.VarChar(10), req.body.GRP)
                request.input("ORDTYP", sql.VarChar(5), req.body.ORDTYP)

                request = await request.execute('usp_DashIssStock')

                res.json({ success: 1, data: request.recordset })
            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.DashStockConfSave = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            sendStatus(401)
        } else {

            const TokenData = await authData;
            let PerArr = req.body;
            let status = true;
            let ErrorPer = []
            let NEWERROR = []
            for (let i = 0; i < PerArr.length; i++) {
                try {
                    var request = new sql.Request();

                    request.input('L_CODE', sql.VarChar(5), PerArr[i].L_CODE)
                    request.input('SRNO', sql.Int, parseInt(PerArr[i].SRNO))
                    request.input('TAG', sql.VarChar(3), PerArr[i].TAG)
                    request.input("PRC_TYP", sql.VarChar(5), PerArr[i].PRC_TYP)
                    request.input('DETID', sql.Int, parseInt(PerArr[i].DETID))
                    request.input('TPROC_CODE', sql.VarChar(10), PerArr[i].TPROC_CODE)
                    request.input('EMP_CODE', sql.VarChar(5), PerArr[i].EMP_CODE)
                    request.input('IUSER', sql.VarChar(20), PerArr[i].IUSER)
                    request.input('ICOMP', sql.Int, parseInt(PerArr[i].ICOMP))

                    request = await request.execute('USP_DashStkConfSave');
                    ErrorPer.push(request.recordset)
                } catch (err) {
                    NEWERROR.push(err)
                }
            }

            res.json({ success: 1, Err: NEWERROR, data: ErrorPer })



            // try {
            //     var request = new sql.Request();

            //     request.input("L_CODE", sql.VarChar(5), req.body.L_CODE)
            //     request.input("SRNO", sql.Int, parseInt(req.body.SRNO))
            //     request.input("TAG", sql.VarChar(3), req.body.TAG)
            //     request.input("PRC_TYP", sql.VarChar(5), req.body.PRC_TYP)
            //     request.input("DETID", sql.Int, parseInt(req.body.DETID))
            //     request.input("TPROC_CODE", sql.VarChar(5), req.body.TPROC_CODE)
            //     request.input("EMP_CODE", sql.VarChar(5), req.body.EMP_CODE)
            //     request.input("IUSER", sql.VarChar(10), req.body.IUSER)
            //     request.input("ICOMP", sql.VarChar(30), req.body.ICOMP)

            //     request = await request.execute('USP_DashStkConfSave')

            //     res.json({ success: 1, data: request.recordset })
            // } catch (err) {
            //     res.json({ success: 0, data: err })
            // }
        }
    })
}


exports.DashStkConfSave = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input("LOT", sql.VarChar(10), req.body.LOT)
                request.input("SRNO", sql.Int, parseInt(req.body.SRNO))
                request.input("TAG", sql.VarChar(5), req.body.TAG)
                request.input("TPROC_CODE", sql.VarChar(10), req.body.TPROC_CODE)
                request.input("PRC_TYP", sql.VarChar(5), req.body.PRC_TYP)

                request = await request.execute('uFn_DashStkConfSave')

                res.json({ success: 1, data: request.recordset })
            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}
