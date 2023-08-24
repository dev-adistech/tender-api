const sql = require("mssql");
const jwt = require('jsonwebtoken');
const Excel = require('exceljs');
const fs = require('fs');
var path = require('path');
const { exec } = require("child_process");

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');
const { markAsUntransferable } = require("worker_threads");
const { send } = require("process");

var { RapServerConn } = require('../../config/db/rapsql');

exports.PacketView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('GRD', sql.Int, parseInt(req.body.GRD))
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('USP_VWFrmPktView');

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

exports.SawMistView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7999), req.body.L_CODE)
                request.input('EMP_CODE', sql.VarChar(7), req.body.EMP_CODE)
                request.input('OEMP_CODE', sql.VarChar(7), req.body.OEMP_CODE)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('F_SRNO', sql.Int, parseInt(req.body.F_SRNO))
                request.input('T_SRNO', sql.Int, parseInt(req.body.T_SRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('F_INO', sql.Int, parseInt(req.body.F_INO))
                request.input('T_INO', sql.Int, parseInt(req.body.T_INO))

                request = await request.execute('usp_VWFrmSawMistakeView');

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

exports.SawMacPro = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(7), req.body.EMP_CODE)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }

                request = await request.execute('USP_VWFrmSawMcWiseProdView');

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

exports.SawMacProSave = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {

            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_NAME', sql.VarChar(100), req.body.EMP_NAME)
                request.input('AVG_PCS', sql.Int, parseInt(req.body.AVG_PCS))

                request = request.execute('usp_VWSawMcAvgSave')

                res.json({ success: 1, data: "" })


            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.SawMacProDet = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(100), req.body.EMP_CODE)
                request.input('DAY', sql.Int, req.body.DAY)
                request.input('SHIFT', sql.VarChar(10), req.body.SHIFT)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }

                request = await request.execute('usp_VWSawMcWiseProdDet');

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

exports.BioDataView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('TYPE', sql.VarChar(10), req.body.TYPE)
                request.input('DEPT_CODE', sql.NVarChar(5), req.body.DEPT_CODE)

                request = await request.execute('USP_VWFrmBioDataView');

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

exports.PacketViewTotal = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('F_CARAT', sql.Decimal(10, 3), parseFloat(req.body.F_CARAT));
                request.input('T_CARAT', sql.Decimal(10, 3), parseFloat(req.body.T_CARAT));
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('USP_VWFrmPktViewTotal');

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

exports.ProcessView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('F_INO', sql.Int, parseInt(req.body.F_INO))
                request.input('T_INO', sql.Int, parseInt(req.body.T_INO))
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('TPROC_CODE', sql.VarChar(5), req.body.TPROC_CODE)
                request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)
                request.input('GRP', sql.VarChar(3), req.body.GRP)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('USP_VWFrmPrcView');
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

exports.GradingAvgEntryView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(2), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)
                request.input('F_CARAT', sql.Numeric(10, 3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10, 3), req.body.T_CARAT)
                request.input('TYPE', sql.VarChar(5), req.body.TYPE)

                request = await request.execute('usp_VWFrmGrdAvgEntView');
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

exports.ProcessViewTotal = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('F_INO', sql.Int, parseInt(req.body.F_INO))
                request.input('T_INO', sql.Int, parseInt(req.body.T_INO))
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('TPROC_CODE', sql.VarChar(5), req.body.TPROC_CODE)
                request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('USP_VWFrmPrcViewTotal');

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

exports.InnerPacketView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('INO', sql.Int, parseInt(req.body.INO))
                request.input('INOTO', sql.Int, parseInt(req.body.INOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('M_CODE', sql.VarChar(5), req.body.M_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('ISFASTTAG', sql.Bit, req.body.ISFASTTAG)
                request.input('ISORG', sql.Bit, req.body.ISORG)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('usp_VWInnerPrcPktViewDisp');

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

exports.InnerPacketViewTotal = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('INO', sql.Int, parseInt(req.body.INO))
                request.input('INOTO', sql.Int, parseInt(req.body.INOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('M_CODE', sql.VarChar(5), req.body.M_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('ISFASTTAG', sql.Bit, req.body.ISFASTTAG)
                request.input('ISORG', sql.Bit, req.body.ISORG)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('usp_VWInnerPrcPktViewTot');

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

exports.InnerProcessView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('INO', sql.Int, parseInt(req.body.INO))
                request.input('INOTO', sql.Int, parseInt(req.body.INOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('M_CODE', sql.VarChar(5), req.body.M_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('OEMP_CODE', sql.VarChar(5), req.body.OEMP_CODE)
                request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)
                request.input('ISFASTTAG', sql.Bit, req.body.ISFASTTAG)
                request.input('ISORG', sql.Bit, req.body.ISORG)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                // console.log(req.body)
                request = await request.execute('usp_VWInnerPrcPrcViewDisp');
                // console.log(request)
                // if (request.recordset) {
                res.json({ success: 1, data: request.recordset })
                // } else {
                // res.json({ success: 0, data: "Not Found" })
                // }

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.InnerProcessViewTotal = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('INO', sql.Int, parseInt(req.body.INO))
                request.input('INOTO', sql.Int, parseInt(req.body.INOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('M_CODE', sql.VarChar(5), req.body.M_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('OEMP_CODE', sql.VarChar(5), req.body.OEMP_CODE)
                request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)
                request.input('ISFASTTAG', sql.Bit, req.body.ISFASTTAG)
                request.input('ISORG', sql.Bit, req.body.ISORG)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)
                // request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                // request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('usp_VWInnerPrcPrcViewTot');

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

exports.PrcViewPrint = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('F_INO', sql.Int, parseInt(req.body.INO))
                request.input('T_INO', sql.Int, parseInt(req.body.INOTO))
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('TPROC_CODE', sql.VarChar(5), req.body.TPROC_CODE)
                request.input('GRP', sql.VarChar(3), req.body.GRP)
                request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('USP_VWFrmPrcViewPrint');

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

exports.PktViewPrint = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('GRP', sql.VarChar(3), req.body.GRP)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('F_CARAT', sql.Decimal(10, 3), parseFloat(req.body.F_CARAT));
                request.input('T_CARAT', sql.Decimal(10, 3), parseFloat(req.body.T_CARAT));
                request.input('GRD', sql.Int, parseInt(req.body.GRD))
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('USP_VWFrmPktViewPrint');

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

exports.PktViewSummPrint = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('GRP', sql.VarChar(3), req.body.GRP)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('F_CARAT', sql.Decimal(10, 3), parseFloat(req.body.F_CARAT));
                request.input('T_CARAT', sql.Decimal(10, 3), parseFloat(req.body.T_CARAT));
                // request.input('F_TIME', sql.Time, req.body.F_TIME)
                // request.input('T_TIME', sql.Time, req.body.T_TIME)
                request.input('GRD', sql.Int, parseInt(req.body.GRD))
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('USP_VWFrmPktViewSummPrint');

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

exports.PktViewDetPrint = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('GRP', sql.VarChar(3), req.body.GRP)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('F_CARAT', sql.Decimal(10, 3), parseFloat(req.body.F_CARAT));
                request.input('T_CARAT', sql.Decimal(10, 3), parseFloat(req.body.T_CARAT));
                // request.input('F_TIME', sql.Time, req.body.F_TIME)
                // request.input('T_TIME', sql.Time, req.body.T_TIME)
                request.input('GRD', sql.Int, parseInt(req.body.GRD))
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('USP_VWFrmPktViewDetPrint');

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

exports.PrcViewDetPrint = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('F_INO', sql.Int, parseInt(req.body.INO))
                request.input('T_INO', sql.Int, parseInt(req.body.INOTO))
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('TPROC_CODE', sql.VarChar(5), req.body.TPROC_CODE)
                request.input('GRP', sql.VarChar(3), req.body.GRP)
                request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('USP_VWFrmPrcViewDetPrint');

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

exports.InnPktViewPrint = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('INO', sql.Int, parseInt(req.body.INO))
                request.input('INOTO', sql.Int, parseInt(req.body.INOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('M_CODE', sql.VarChar(5), req.body.M_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('usp_VWInnerPrcPktViewPrintSum');

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

exports.InnPrcViewPrint = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('INO', sql.Int, parseInt(req.body.INO))
                request.input('INOTO', sql.Int, parseInt(req.body.INOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('M_CODE', sql.VarChar(5), req.body.M_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('OEMP_CODE', sql.VarChar(5), req.body.OEMP_CODE)
                request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('usp_VWInnerPrcPrcViewPrintSum');

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

exports.EmpIdView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('PEMP_CODE', sql.VarChar(10), req.body.PEMP_CODE)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('M_CODE', sql.VarChar(5), req.body.M_CODE)
                request.input('PRDTYPE', sql.VarChar(5), req.body.PRDTYPE)
                request.input('TYP', sql.VarChar(10), req.body.TYP)
                request.input('TRF', sql.VarChar(5), req.body.TRF)
                request.input('COLUMN', sql.VarChar(20), req.body.COLUMN)

                request = await request.execute('usp_EmpIdView');

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

exports.EmpPrdView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('PEMP_CODE', sql.VarChar(10), req.body.PEMP_CODE)
                request.input('TYP', sql.VarChar(10), req.body.TYP)
                request.input('COLUMN', sql.VarChar(20), req.body.COLUMN)

                request = await request.execute('usp_EmpPrdView');

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

exports.PolishStkView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('P_CODE', sql.VarChar(10), req.body.P_CODE)
                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('TYPE', sql.VarChar(5), req.body.TYPE)

                request = await request.execute('USP_VWMPolishStkView');

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

exports.PolishStkViewTot = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('P_CODE', sql.VarChar(10), req.body.P_CODE)
                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('TYPE', sql.VarChar(5), req.body.TYPE)

                request = await request.execute('USP_VWMPolishStkViewTot');

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

exports.LivePktView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(5), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(5), req.body.Q_CODE)
                request.input('F_CARAT', sql.Decimal(10, 3), parseFloat(req.body.F_CARAT));
                request.input('T_CARAT', sql.Decimal(10, 3), parseFloat(req.body.T_CARAT));
                // request.input('F_DATE', sql.DateTime2, null)
                // request.input('T_DATE', sql.DateTime2, null)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('CUT_CODE', sql.Int, parseInt(req.body.CUT_CODE))
                request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('TYP', sql.VarChar(5), req.body.TYP)
                request.input('PRDTYPE', sql.VarChar(5), req.body.PRDTYPE)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)

                request = await request.execute('usp_VWFrmLivePktView');

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

exports.LivePktViewExcel = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7999), req.body.L_CODE)
                request.input('PRDTYPE', sql.VarChar(10), req.body.PRDTYPE)
                request.input('TYP', sql.VarChar(5), req.body.TYP)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)

                request = await request.execute('USP_LivePktViewExcel');

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

exports.LivePktViewExcelExport = async (req, res) => {
    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet('Proposal');
    const DataToExport = JSON.parse(req.body.DataRow);
    const colorCell = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']

    let shapeRowData = DataToExport[1].map(item => Object.values(item));
    let shapeIndex = 3
    let shapeMergeCellIndex = shapeIndex - 1
    worksheet.mergeCells('B' + shapeMergeCellIndex + ':F' + shapeMergeCellIndex);
    worksheet.getCell('B' + shapeMergeCellIndex).value = 'Shape Proposal';
    worksheet.getCell('B' + shapeMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('B' + shapeMergeCellIndex).font = {
        bold: true,
    };
    worksheet.addTable({
        name: 'shapeProposal',
        ref: 'B' + shapeIndex,
        headerRow: true,
        totalsRow: false,
        style: {
            theme: '',
            showRowStripes: true,
            showFirstColumn: true,
        },
        columns: [
            { name: 'Shape Detail' },
            { name: 'Pcs' },
            { name: 'Weight' },
            { name: 'Carat%' },
            { name: 'Price%' },
        ],
        rows: shapeRowData,
    });

    // background color and fill border
    for (let i = 0; i < 5; i++) {
        worksheet.getCell('B' + (shapeIndex - 1).toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell('B' + (shapeIndex - 1).toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell(colorCell[1 + i] + shapeIndex.toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell(colorCell[1 + i] + shapeIndex.toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        for (let j = 1; j <= DataToExport[1].length; j++) {
            worksheet.getCell(colorCell[1 + i] + (shapeIndex + j).toString()).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            if (j == DataToExport[1].length) {
                worksheet.getCell(colorCell[1 + i] + (shapeIndex + j).toString()).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'c5d9f1' },
                };
            }
        }
    }

    let SizeRowData = DataToExport[2].map(item => Object.values(item));
    let SizeIndex = shapeIndex + DataToExport[1].length + 4
    let SizeMergeCellIndex = SizeIndex - 1
    worksheet.mergeCells('B' + SizeMergeCellIndex + ':F' + SizeMergeCellIndex);
    worksheet.getCell('B' + SizeMergeCellIndex).value = 'Size Proposal';
    worksheet.getCell('B' + SizeMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('B' + SizeMergeCellIndex).font = {
        bold: true,
    };
    worksheet.addTable({
        name: 'SizeProposal',
        ref: 'B' + SizeIndex,
        headerRow: true,
        totalsRow: false,
        style: {
            theme: '',
            showRowStripes: true,
            showFirstColumn: true,
        },
        columns: [
            { name: 'Size Detail' },
            { name: 'Pcs' },
            { name: 'Weight' },
            { name: 'Carat%' },
            { name: 'Price%' },
        ],
        rows: SizeRowData,
    });

    // background color and fill border
    for (let i = 0; i < 5; i++) {
        worksheet.getCell('B' + (SizeIndex - 1).toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell('B' + (SizeIndex - 1).toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell(colorCell[1 + i] + SizeIndex.toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell(colorCell[1 + i] + SizeIndex.toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        for (let j = 1; j <= DataToExport[2].length; j++) {
            worksheet.getCell(colorCell[1 + i] + (SizeIndex + j).toString()).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            if (j == DataToExport[2].length) {
                worksheet.getCell(colorCell[1 + i] + (SizeIndex + j).toString()).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'c5d9f1' },
                };
            }
        }
    }

    let ColorRowData = DataToExport[3].map(item => Object.values(item));
    let ColorIndex = 3
    let ColorMergeCellIndex = ColorIndex - 1
    worksheet.mergeCells('H' + ColorMergeCellIndex + ':L' + ColorMergeCellIndex);
    worksheet.getCell('H' + ColorMergeCellIndex).value = 'Color Proposal';
    worksheet.getCell('H' + ColorMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('H' + ColorMergeCellIndex).font = {
        bold: true,
    };
    worksheet.addTable({
        name: 'ColorProposal',
        ref: 'H' + ColorIndex,
        headerRow: true,
        totalsRow: false,
        style: {
            theme: '',
            showRowStripes: true,
            showFirstColumn: true,
        },
        columns: [
            { name: 'Color Detail' },
            { name: 'Pcs' },
            { name: 'Weight' },
            { name: 'Carat%' },
            { name: 'Price%' },
        ],
        rows: ColorRowData,
    });

    // background color and fill border
    for (let i = 0; i < 5; i++) {
        worksheet.getCell('H' + (ColorIndex - 1).toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell('H' + (ColorIndex - 1).toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell(colorCell[7 + i] + ColorIndex.toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell(colorCell[7 + i] + ColorIndex.toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        for (let j = 1; j <= DataToExport[3].length; j++) {
            worksheet.getCell(colorCell[7 + i] + (ColorIndex + j).toString()).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            if (j == DataToExport[3].length) {
                worksheet.getCell(colorCell[7 + i] + (ColorIndex + j).toString()).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'c5d9f1' },
                };
            }
        }
    }

    let QualityRowData = DataToExport[4].map(item => Object.values(item));
    let QualityIndex = shapeIndex + DataToExport[3].length + 4
    let QualityMergeCellIndex = QualityIndex - 1
    worksheet.mergeCells('H' + QualityMergeCellIndex + ':L' + QualityMergeCellIndex);
    worksheet.getCell('H' + QualityMergeCellIndex).value = 'Quality Proposal';
    worksheet.getCell('H' + QualityMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('H' + QualityMergeCellIndex).font = {
        bold: true,
    };
    worksheet.addTable({
        name: 'ColorProposal',
        ref: 'H' + QualityIndex,
        headerRow: true,
        totalsRow: false,
        style: {
            theme: '',
            showRowStripes: true,
            showFirstColumn: true,
        },
        columns: [
            { name: 'Color Detail' },
            { name: 'Pcs' },
            { name: 'Weight' },
            { name: 'Carat%' },
            { name: 'Price%' },
        ],
        rows: QualityRowData,
    });

    // background color and fill border
    for (let i = 0; i < 5; i++) {
        worksheet.getCell('H' + (QualityIndex - 1).toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell('H' + (QualityIndex - 1).toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell(colorCell[7 + i] + QualityIndex.toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell(colorCell[7 + i] + QualityIndex.toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        for (let j = 1; j <= DataToExport[4].length; j++) {
            worksheet.getCell(colorCell[7 + i] + (QualityIndex + j).toString()).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            if (j == DataToExport[4].length) {
                worksheet.getCell(colorCell[7 + i] + (QualityIndex + j).toString()).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'c5d9f1' },
                };
            }
        }
    }

    let ShadeRowData = DataToExport[5].map(item => Object.values(item));
    let ShadeIndex = QualityIndex + DataToExport[4].length + 4
    let ShadeMergeCellIndex = ShadeIndex - 1
    worksheet.mergeCells('H' + ShadeMergeCellIndex + ':L' + ShadeMergeCellIndex);
    worksheet.getCell('H' + ShadeMergeCellIndex).value = 'Shade Proposal';
    worksheet.getCell('H' + ShadeMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('H' + ShadeMergeCellIndex).font = {
        bold: true,
    };
    worksheet.addTable({
        name: 'ColorProposal',
        ref: 'H' + ShadeIndex,
        headerRow: true,
        totalsRow: false,
        style: {
            theme: '',
            showRowStripes: true,
            showFirstColumn: true,
        },
        columns: [
            { name: 'Color Detail' },
            { name: 'Pcs' },
            { name: 'Weight' },
            { name: 'Carat%' },
            { name: 'Price%' },
        ],
        rows: ShadeRowData,
    });

    // background color and fill border
    for (let i = 0; i < 5; i++) {
        worksheet.getCell('H' + (ShadeIndex - 1).toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell('H' + (ShadeIndex - 1).toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell(colorCell[7 + i] + ShadeIndex.toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell(colorCell[7 + i] + ShadeIndex.toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        for (let j = 1; j <= DataToExport[5].length; j++) {
            worksheet.getCell(colorCell[7 + i] + (ShadeIndex + j).toString()).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            if (j == DataToExport[5].length) {
                worksheet.getCell(colorCell[7 + i] + (ShadeIndex + j).toString()).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'c5d9f1' },
                };
            }
        }
    }

    let CutRowData = DataToExport[6].map(item => Object.values(item));
    let CutIndex = 3
    let CutMergeCellIndex = CutIndex - 1
    worksheet.mergeCells('N' + CutMergeCellIndex + ':R' + CutMergeCellIndex);
    worksheet.getCell('N' + CutMergeCellIndex).value = 'Cut Proposal';
    worksheet.getCell('N' + CutMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('N' + CutMergeCellIndex).font = {
        bold: true,
    };
    worksheet.addTable({
        name: 'CutProposal',
        ref: 'N' + CutIndex,
        headerRow: true,
        totalsRow: false,
        style: {
            theme: '',
            showRowStripes: true,
            showFirstColumn: true,
        },
        columns: [
            { name: 'Cut Detail' },
            { name: 'Pcs' },
            { name: 'Weight' },
            { name: 'Carat%' },
            { name: 'Price%' },
        ],
        rows: CutRowData,
    });

    // background color and fill border
    for (let i = 0; i < 5; i++) {
        worksheet.getCell('N' + (CutIndex - 1).toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell('N' + (CutIndex - 1).toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell(colorCell[13 + i] + CutIndex.toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell(colorCell[13 + i] + CutIndex.toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        for (let j = 1; j <= DataToExport[6].length; j++) {
            worksheet.getCell(colorCell[13 + i] + (CutIndex + j).toString()).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            if (j == DataToExport[6].length) {
                worksheet.getCell(colorCell[13 + i] + (CutIndex + j).toString()).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'c5d9f1' },
                };
            }
        }
    }

    let PolishRowData = DataToExport[7].map(item => Object.values(item));
    let PolishIndex = CutIndex + DataToExport[6].length + 4
    let PolishMergeCellIndex = PolishIndex - 1
    worksheet.mergeCells('N' + PolishMergeCellIndex + ':R' + PolishMergeCellIndex);
    worksheet.getCell('N' + PolishMergeCellIndex).value = 'Polish Proposal';
    worksheet.getCell('N' + PolishMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('N' + PolishMergeCellIndex).font = {
        bold: true,
    };
    worksheet.addTable({
        name: 'PolishProposal',
        ref: 'N' + PolishIndex,
        headerRow: true,
        totalsRow: false,
        style: {
            theme: '',
            showRowStripes: true,
            showFirstColumn: true,
        },
        columns: [
            { name: 'Polish Detail' },
            { name: 'Pcs' },
            { name: 'Weight' },
            { name: 'Carat%' },
            { name: 'Price%' },
        ],
        rows: PolishRowData,
    });

    // background color and fill border
    for (let i = 0; i < 5; i++) {
        worksheet.getCell('N' + (PolishIndex - 1).toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell('N' + (PolishIndex - 1).toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell(colorCell[13 + i] + PolishIndex.toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell(colorCell[13 + i] + PolishIndex.toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        for (let j = 1; j <= DataToExport[7].length; j++) {
            worksheet.getCell(colorCell[13 + i] + (PolishIndex + j).toString()).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            if (j == DataToExport[7].length) {
                worksheet.getCell(colorCell[13 + i] + (PolishIndex + j).toString()).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'c5d9f1' },
                };
            }
        }
    }

    let SymRowData = DataToExport[8].map(item => Object.values(item));
    let SymIndex = PolishIndex + DataToExport[7].length + 4
    let SymMergeCellIndex = SymIndex - 1
    worksheet.mergeCells('N' + SymMergeCellIndex + ':R' + SymMergeCellIndex);
    worksheet.getCell('N' + SymMergeCellIndex).value = 'Symmentry Proposal';
    worksheet.getCell('N' + SymMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('N' + SymMergeCellIndex).font = {
        bold: true,
    };
    worksheet.addTable({
        name: 'PolishProposal',
        ref: 'N' + SymIndex,
        headerRow: true,
        totalsRow: false,
        style: {
            theme: '',
            showRowStripes: true,
            showFirstColumn: true,
        },
        columns: [
            { name: 'Polish Detail' },
            { name: 'Pcs' },
            { name: 'Weight' },
            { name: 'Carat%' },
            { name: 'Price%' },
        ],
        rows: SymRowData,
    });

    // background color and fill border
    for (let i = 0; i < 5; i++) {
        worksheet.getCell('N' + (SymIndex - 1).toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell('N' + (SymIndex - 1).toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell(colorCell[13 + i] + SymIndex.toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell(colorCell[13 + i] + SymIndex.toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        for (let j = 1; j <= DataToExport[8].length; j++) {
            worksheet.getCell(colorCell[13 + i] + (SymIndex + j).toString()).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            if (j == DataToExport[8].length) {
                worksheet.getCell(colorCell[13 + i] + (SymIndex + j).toString()).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'c5d9f1' },
                };
            }
        }
    }

    let FloRowData = DataToExport[9].map(item => Object.values(item));
    let FloIndex = SymIndex + DataToExport[8].length + 4
    let FloMergeCellIndex = FloIndex - 1
    worksheet.mergeCells('N' + FloMergeCellIndex + ':R' + FloMergeCellIndex);
    worksheet.getCell('N' + FloMergeCellIndex).value = 'Floro Proposal';
    worksheet.getCell('N' + FloMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('N' + FloMergeCellIndex).font = {
        bold: true,
    };
    worksheet.addTable({
        name: 'PolishProposal',
        ref: 'N' + FloIndex,
        headerRow: true,
        totalsRow: false,
        style: {
            theme: '',
            showRowStripes: true,
            showFirstColumn: true,
        },
        columns: [
            { name: 'Polish Detail' },
            { name: 'Pcs' },
            { name: 'Weight' },
            { name: 'Carat%' },
            { name: 'Price%' },
        ],
        rows: FloRowData,
    });

    // background color and fill border
    for (let i = 0; i < 5; i++) {
        worksheet.getCell('N' + (FloIndex - 1).toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell('N' + (FloIndex - 1).toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell(colorCell[13 + i] + FloIndex.toString()).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f0f0f0' },
        };
        worksheet.getCell(colorCell[13 + i] + FloIndex.toString()).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        for (let j = 1; j <= DataToExport[9].length; j++) {
            worksheet.getCell(colorCell[13 + i] + (FloIndex + j).toString()).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            if (j == DataToExport[9].length) {
                worksheet.getCell(colorCell[13 + i] + (FloIndex + j).toString()).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'c5d9f1' },
                };
            }
        }
    }

    workbook.xlsx.writeBuffer().then(function (buffer) {
        let xlsData = Buffer.concat([buffer]);
        res.writeHead(200, {
            'Content-Length': Buffer.byteLength(xlsData),
            'Content-Type': 'application/vnd.ms-excel',
            'Content-disposition': 'attachment;filename=data.xlsx',
        }).end(xlsData);
    });
}

exports.MakableView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7999), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(3), req.body.TAG)
                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('Q_CODE', sql.Int, parseInt(req.body.Q_CODE))
                request.input('C_CODE', sql.Int, parseInt(req.body.C_CODE))
                request.input('CT_CODE', sql.Int, parseInt(req.body.CT_CODE))
                request.input('PO_CODE', sql.Int, parseInt(req.body.PO_CODE))
                request.input('SY_CODE', sql.Int, parseInt(req.body.SY_CODE))
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('F_SIZE', sql.Decimal(10, 3), parseFloat(req.body.F_SIZE));
                request.input('T_SIZE', sql.Decimal(10, 3), parseFloat(req.body.T_SIZE));
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('PROC_CODE', sql.VarChar(10), req.body.PROC_CODE)
                request.input('TYPE', sql.VarChar(5), req.body.TYPE)
                request.input('TYP', sql.VarChar(5), req.body.TYP)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))

                request = await request.execute('usp_VWFrmMakableView');

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

exports.RouEntViewDisp = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('F_SRNO', sql.Int, parseInt(req.body.F_SRNO))
                request.input('T_SRNO', sql.Int, parseInt(req.body.T_SRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)

                request = await request.execute('usp_VWFrmRouEntViewDisp');

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

exports.RouEntViewTotDisp = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7999), req.body.L_CODE)
                request.input('F_SRNO', sql.Int, parseInt(req.body.F_SRNO))
                request.input('T_SRNO', sql.Int, parseInt(req.body.T_SRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMode', sql.VarChar(10), req.body.SELECTMode)

                request = await request.execute('usp_VEFrmRouEntViewTotDisp');

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

exports.PrdQueryDisp = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('Q_CODE', sql.Int, parseInt(req.body.Q_CODE))
                request.input('C_CODE', sql.Int, parseInt(req.body.C_CODE))
                request.input('CUT_CODE', sql.Int, parseInt(req.body.CUT_CODE))
                request.input('SYM_CODE', sql.Int, parseInt(req.body.SYM_CODE))
                request.input('POL_CODE', sql.Int, parseInt(req.body.POL_CODE))
                request.input('F_SIZE', sql.Decimal(10, 3), parseFloat(req.body.F_SIZE));
                request.input('T_SIZE', sql.Decimal(10, 3), parseFloat(req.body.T_SIZE));
                request.input('CARAT', sql.Decimal(10, 3), parseFloat(req.body.CARAT));
                request.input('LDIS', sql.Decimal(10, 2), parseFloat(req.body.LDIS));
                request.input('GDIS', sql.Decimal(10, 2), parseFloat(req.body.GDIS));
                request.input('RTYPE', sql.VarChar(5), req.body.RTYPE)
                request.input('DETNO', sql.Int, parseInt(req.body.DETNO))
                request.input('SOLVE', sql.Int, parseInt(req.body.SOLVE))
                request.input('TYP', sql.VarChar(10), req.body.TYP)
                request.input('RAPTYPE', sql.VarChar(10), req.body.RAPTYPE)

                request = await request.execute('usp_VWFrmPrdQueryDisp');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "" })
                }

            } catch (err) {
                console.log("---------------", err)
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.EstToGiaView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('C_CODE', sql.Int, parseInt(req.body.C_CODE))
                request.input('Q_CODE', sql.Int, parseInt(req.body.Q_CODE))
                request.input('CUT_CODE', sql.Int, parseInt(req.body.CUT_CODE))
                request.input('POL_CODE', sql.Int, parseInt(req.body.POL_CODE))
                request.input('SYM_CODE', sql.Int, parseInt(req.body.SYM_CODE))
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)
                request.input('DIFF', sql.VarChar(5), req.body.DIFF)
                request.input('PRDTYP', sql.VarChar(5), req.body.PRDTYP)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))

                request = await request.execute('usp_VWFrmEstToGiaView');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.EstToGiaDiffView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('C_CODE', sql.Int, parseInt(req.body.C_CODE))
                request.input('Q_CODE', sql.Int, parseInt(req.body.Q_CODE))
                request.input('CUT_CODE', sql.Int, parseInt(req.body.CUT_CODE))
                request.input('POL_CODE', sql.Int, parseInt(req.body.POL_CODE))
                request.input('SYM_CODE', sql.Int, parseInt(req.body.SYM_CODE))
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)
                request.input('DIFF', sql.VarChar(5), req.body.DIFF)
                request.input('PRDTYPE', sql.VarChar(5), req.body.PRDTYPE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('DISTYP', sql.VarChar(5), req.body.DISTYP)

                request = await request.execute('usp_VWFrmEstToGiaDiffView');

                if (request.recordsets) {
                    res.json({ success: 1, data: request.recordsets })
                } else {
                    res.json({ success: 1, data: "" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.VWCompare = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                let L_CODE = ''
                if (req.body.L_CODE != '') {
                    let tempTxt = (req.body.L_CODE).split(",")
                    tempTxt.map((x, i) => {
                        if (i == tempTxt.length - 1) {
                            L_CODE += "'" + x + "'"
                        } else {
                            L_CODE += "'" + x + "',"
                        }
                    })
                }

                request.input('L_CODE', sql.VarChar(7991), L_CODE)
                request.input('TYP', sql.VarChar(10), req.body.TYP)
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('FComp', sql.VarChar(10), req.body.FComp)
                request.input('TComp', sql.VarChar(10), req.body.TComp)
                request.input('F_DATE', sql.VarChar(20), req.body.F_DATE)
                request.input('T_DATE', sql.VarChar(20), req.body.T_DATE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('SELECT', sql.VarChar(10), req.body.SELECT)
                request.input('EMPTYP', sql.VarChar(10), req.body.EMPTYP)

                request = await request.execute('usp_VWCompare');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.VWCompareDet = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                let L_CODE = ''
                if (req.body.L_CODE != '') {
                    let tempTxt = (req.body.L_CODE).split(",")
                    tempTxt.map((x, i) => {
                        if (i == tempTxt.length - 1) {
                            L_CODE += "'" + x + "'"
                        } else {
                            L_CODE += "'" + x + "',"
                        }
                    })
                }

                request.input('L_CODE', sql.VarChar(7991), L_CODE)
                request.input('TAG', sql.VarChar(7991), req.body.TAG)
                request.input('TYP', sql.VarChar(10), req.body.TYP)
                request.input('FComp', sql.VarChar(10), req.body.FComp)
                request.input('TComp', sql.VarChar(10), req.body.TComp)
                request.input('FCode', sql.VarChar(10), req.body.FCode)
                request.input('TCode', sql.VarChar(10), req.body.TCode)
                request.input('F_DATE', sql.VarChar(20), req.body.F_DATE)
                request.input('T_DATE', sql.VarChar(20), req.body.T_DATE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('SELECT', sql.VarChar(10), req.body.SELECT)
                request.input('EMPTYP', sql.VarChar(10), req.body.EMPTYP)

                request = await request.execute('usp_VWCompareDet');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.CompareViewPrint = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                let L_CODE = ''
                if (req.body.L_CODE != '') {
                    let tempTxt = (req.body.L_CODE).split(",")
                    tempTxt.map((x, i) => {
                        if (i == tempTxt.length - 1) {
                            L_CODE += "'" + x + "'"
                        } else {
                            L_CODE += "'" + x + "',"
                        }
                    })
                }
                request.input('L_CODE', sql.VarChar(7991), L_CODE)
                request.input('TYP', sql.VarChar(10), req.body.TYP)
                request.input('FComp', sql.VarChar(10), req.body.FComp)
                request.input('TComp', sql.VarChar(10), req.body.TComp)
                request.input('F_DATE', sql.VarChar(20), req.body.F_DATE)
                request.input('T_DATE', sql.VarChar(20), req.body.T_DATE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('SELECT', sql.VarChar(10), req.body.SELECT)
                request.input('EMPTYP', sql.VarChar(10), req.body.EMPTYP)

                request = await request.execute('RP_CompareViewPrint');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.CompareViewExport = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                let L_CODE = ''
                if (req.body.L_CODE != '') {
                    let tempTxt = (req.body.L_CODE).split(",")
                    tempTxt.map((x, i) => {
                        if (i == tempTxt.length - 1) {
                            L_CODE += "'" + x + "'"
                        } else {
                            L_CODE += "'" + x + "',"
                        }
                    })
                }

                request.input('L_CODE', sql.VarChar(7991), L_CODE)
                request.input('TYP', sql.VarChar(10), req.body.TYP)
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('FComp', sql.VarChar(10), req.body.FComp)
                request.input('TComp', sql.VarChar(10), req.body.TComp)
                request.input('F_DATE', sql.VarChar(20), req.body.F_DATE)
                request.input('T_DATE', sql.VarChar(20), req.body.T_DATE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('SELECT', sql.VarChar(10), req.body.SELECT)
                request.input('EMPTYP', sql.VarChar(10), req.body.EMPTYP)

                request = await request.execute('RP_CompareViewExcel');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });

}

// exports.CompareViewExcelDownload = async (req, res) => {
//     let workbook = new Excel.Workbook();
//     let worksheet = workbook.addWorksheet('Proposal');
//     const DataToExport = JSON.parse(req.body.DataRow);
//     const colorCell = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']

//     let shapeRowData = DataToExport.map(item => Object.values(item));
//     let shapeIndex = 3
//     let shapeMergeCellIndex = shapeIndex - 1
//     worksheet.mergeCells('B' + shapeMergeCellIndex + ':F' + shapeMergeCellIndex);
//     worksheet.getCell('B' + shapeMergeCellIndex).value = 'Shape Proposal';
//     worksheet.getCell('B' + shapeMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
//     worksheet.getCell('B' + shapeMergeCellIndex).font = {
//         bold: true,
//     };
//     worksheet.addTable({
//         name: 'shapeProposal',
//         ref: 'B' + shapeIndex,
//         headerRow: true,
//         totalsRow: false,
//         style: {
//             theme: '',
//             showRowStripes: true,
//             showFirstColumn: true,
//         },
//         columns: [
//             { name: 'Shape Detail' },
//             { name: 'Pcs' },
//             { name: 'Weight' },
//             { name: 'Carat%' },
//             { name: 'Price%' },
//         ],
//         rows: shapeRowData,
//     });

//     // background color and fill border
//     for (let i = 0; i < 5; i++) {
//         worksheet.getCell('B' + (shapeIndex - 1).toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell('B' + (shapeIndex - 1).toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         worksheet.getCell(colorCell[1 + i] + shapeIndex.toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell(colorCell[1 + i] + shapeIndex.toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         for (let j = 1; j <= DataToExport[1].length; j++) {
//             worksheet.getCell(colorCell[1 + i] + (shapeIndex + j).toString()).border = {
//                 top: { style: 'thin' },
//                 left: { style: 'thin' },
//                 bottom: { style: 'thin' },
//                 right: { style: 'thin' }
//             };
//             if (j == DataToExport[1].length) {
//                 worksheet.getCell(colorCell[1 + i] + (shapeIndex + j).toString()).fill = {
//                     type: 'pattern',
//                     pattern: 'solid',
//                     fgColor: { argb: 'c5d9f1' },
//                 };
//             }
//         }
//     }

//     let SizeRowData = DataToExport[2].map(item => Object.values(item));
//     let SizeIndex = shapeIndex + DataToExport[1].length + 4
//     let SizeMergeCellIndex = SizeIndex - 1
//     worksheet.mergeCells('B' + SizeMergeCellIndex + ':F' + SizeMergeCellIndex);
//     worksheet.getCell('B' + SizeMergeCellIndex).value = 'Size Proposal';
//     worksheet.getCell('B' + SizeMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
//     worksheet.getCell('B' + SizeMergeCellIndex).font = {
//         bold: true,
//     };
//     worksheet.addTable({
//         name: 'SizeProposal',
//         ref: 'B' + SizeIndex,
//         headerRow: true,
//         totalsRow: false,
//         style: {
//             theme: '',
//             showRowStripes: true,
//             showFirstColumn: true,
//         },
//         columns: [
//             { name: 'Size Detail' },
//             { name: 'Pcs' },
//             { name: 'Weight' },
//             { name: 'Carat%' },
//             { name: 'Price%' },
//         ],
//         rows: SizeRowData,
//     });

//     // background color and fill border
//     for (let i = 0; i < 5; i++) {
//         worksheet.getCell('B' + (SizeIndex - 1).toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell('B' + (SizeIndex - 1).toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         worksheet.getCell(colorCell[1 + i] + SizeIndex.toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell(colorCell[1 + i] + SizeIndex.toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         for (let j = 1; j <= DataToExport[2].length; j++) {
//             worksheet.getCell(colorCell[1 + i] + (SizeIndex + j).toString()).border = {
//                 top: { style: 'thin' },
//                 left: { style: 'thin' },
//                 bottom: { style: 'thin' },
//                 right: { style: 'thin' }
//             };
//             if (j == DataToExport[2].length) {
//                 worksheet.getCell(colorCell[1 + i] + (SizeIndex + j).toString()).fill = {
//                     type: 'pattern',
//                     pattern: 'solid',
//                     fgColor: { argb: 'c5d9f1' },
//                 };
//             }
//         }
//     }

//     let ColorRowData = DataToExport[3].map(item => Object.values(item));
//     let ColorIndex = 3
//     let ColorMergeCellIndex = ColorIndex - 1
//     worksheet.mergeCells('H' + ColorMergeCellIndex + ':L' + ColorMergeCellIndex);
//     worksheet.getCell('H' + ColorMergeCellIndex).value = 'Color Proposal';
//     worksheet.getCell('H' + ColorMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
//     worksheet.getCell('H' + ColorMergeCellIndex).font = {
//         bold: true,
//     };
//     worksheet.addTable({
//         name: 'ColorProposal',
//         ref: 'H' + ColorIndex,
//         headerRow: true,
//         totalsRow: false,
//         style: {
//             theme: '',
//             showRowStripes: true,
//             showFirstColumn: true,
//         },
//         columns: [
//             { name: 'Color Detail' },
//             { name: 'Pcs' },
//             { name: 'Weight' },
//             { name: 'Carat%' },
//             { name: 'Price%' },
//         ],
//         rows: ColorRowData,
//     });

//     // background color and fill border
//     for (let i = 0; i < 5; i++) {
//         worksheet.getCell('H' + (ColorIndex - 1).toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell('H' + (ColorIndex - 1).toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         worksheet.getCell(colorCell[7 + i] + ColorIndex.toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell(colorCell[7 + i] + ColorIndex.toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         for (let j = 1; j <= DataToExport[3].length; j++) {
//             worksheet.getCell(colorCell[7 + i] + (ColorIndex + j).toString()).border = {
//                 top: { style: 'thin' },
//                 left: { style: 'thin' },
//                 bottom: { style: 'thin' },
//                 right: { style: 'thin' }
//             };
//             if (j == DataToExport[3].length) {
//                 worksheet.getCell(colorCell[7 + i] + (ColorIndex + j).toString()).fill = {
//                     type: 'pattern',
//                     pattern: 'solid',
//                     fgColor: { argb: 'c5d9f1' },
//                 };
//             }
//         }
//     }

//     let QualityRowData = DataToExport[4].map(item => Object.values(item));
//     let QualityIndex = shapeIndex + DataToExport[3].length + 4
//     let QualityMergeCellIndex = QualityIndex - 1
//     worksheet.mergeCells('H' + QualityMergeCellIndex + ':L' + QualityMergeCellIndex);
//     worksheet.getCell('H' + QualityMergeCellIndex).value = 'Quality Proposal';
//     worksheet.getCell('H' + QualityMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
//     worksheet.getCell('H' + QualityMergeCellIndex).font = {
//         bold: true,
//     };
//     worksheet.addTable({
//         name: 'ColorProposal',
//         ref: 'H' + QualityIndex,
//         headerRow: true,
//         totalsRow: false,
//         style: {
//             theme: '',
//             showRowStripes: true,
//             showFirstColumn: true,
//         },
//         columns: [
//             { name: 'Color Detail' },
//             { name: 'Pcs' },
//             { name: 'Weight' },
//             { name: 'Carat%' },
//             { name: 'Price%' },
//         ],
//         rows: QualityRowData,
//     });

//     // background color and fill border
//     for (let i = 0; i < 5; i++) {
//         worksheet.getCell('H' + (QualityIndex - 1).toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell('H' + (QualityIndex - 1).toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         worksheet.getCell(colorCell[7 + i] + QualityIndex.toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell(colorCell[7 + i] + QualityIndex.toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         for (let j = 1; j <= DataToExport[4].length; j++) {
//             worksheet.getCell(colorCell[7 + i] + (QualityIndex + j).toString()).border = {
//                 top: { style: 'thin' },
//                 left: { style: 'thin' },
//                 bottom: { style: 'thin' },
//                 right: { style: 'thin' }
//             };
//             if (j == DataToExport[4].length) {
//                 worksheet.getCell(colorCell[7 + i] + (QualityIndex + j).toString()).fill = {
//                     type: 'pattern',
//                     pattern: 'solid',
//                     fgColor: { argb: 'c5d9f1' },
//                 };
//             }
//         }
//     }

//     let ShadeRowData = DataToExport[5].map(item => Object.values(item));
//     let ShadeIndex = QualityIndex + DataToExport[4].length + 4
//     let ShadeMergeCellIndex = ShadeIndex - 1
//     worksheet.mergeCells('H' + ShadeMergeCellIndex + ':L' + ShadeMergeCellIndex);
//     worksheet.getCell('H' + ShadeMergeCellIndex).value = 'Shade Proposal';
//     worksheet.getCell('H' + ShadeMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
//     worksheet.getCell('H' + ShadeMergeCellIndex).font = {
//         bold: true,
//     };
//     worksheet.addTable({
//         name: 'ColorProposal',
//         ref: 'H' + ShadeIndex,
//         headerRow: true,
//         totalsRow: false,
//         style: {
//             theme: '',
//             showRowStripes: true,
//             showFirstColumn: true,
//         },
//         columns: [
//             { name: 'Color Detail' },
//             { name: 'Pcs' },
//             { name: 'Weight' },
//             { name: 'Carat%' },
//             { name: 'Price%' },
//         ],
//         rows: ShadeRowData,
//     });

//     // background color and fill border
//     for (let i = 0; i < 5; i++) {
//         worksheet.getCell('H' + (ShadeIndex - 1).toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell('H' + (ShadeIndex - 1).toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         worksheet.getCell(colorCell[7 + i] + ShadeIndex.toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell(colorCell[7 + i] + ShadeIndex.toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         for (let j = 1; j <= DataToExport[5].length; j++) {
//             worksheet.getCell(colorCell[7 + i] + (ShadeIndex + j).toString()).border = {
//                 top: { style: 'thin' },
//                 left: { style: 'thin' },
//                 bottom: { style: 'thin' },
//                 right: { style: 'thin' }
//             };
//             if (j == DataToExport[5].length) {
//                 worksheet.getCell(colorCell[7 + i] + (ShadeIndex + j).toString()).fill = {
//                     type: 'pattern',
//                     pattern: 'solid',
//                     fgColor: { argb: 'c5d9f1' },
//                 };
//             }
//         }
//     }

//     let CutRowData = DataToExport[6].map(item => Object.values(item));
//     let CutIndex = 3
//     let CutMergeCellIndex = CutIndex - 1
//     worksheet.mergeCells('N' + CutMergeCellIndex + ':R' + CutMergeCellIndex);
//     worksheet.getCell('N' + CutMergeCellIndex).value = 'Cut Proposal';
//     worksheet.getCell('N' + CutMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
//     worksheet.getCell('N' + CutMergeCellIndex).font = {
//         bold: true,
//     };
//     worksheet.addTable({
//         name: 'CutProposal',
//         ref: 'N' + CutIndex,
//         headerRow: true,
//         totalsRow: false,
//         style: {
//             theme: '',
//             showRowStripes: true,
//             showFirstColumn: true,
//         },
//         columns: [
//             { name: 'Cut Detail' },
//             { name: 'Pcs' },
//             { name: 'Weight' },
//             { name: 'Carat%' },
//             { name: 'Price%' },
//         ],
//         rows: CutRowData,
//     });

//     // background color and fill border
//     for (let i = 0; i < 5; i++) {
//         worksheet.getCell('N' + (CutIndex - 1).toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell('N' + (CutIndex - 1).toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         worksheet.getCell(colorCell[13 + i] + CutIndex.toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell(colorCell[13 + i] + CutIndex.toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         for (let j = 1; j <= DataToExport[6].length; j++) {
//             worksheet.getCell(colorCell[13 + i] + (CutIndex + j).toString()).border = {
//                 top: { style: 'thin' },
//                 left: { style: 'thin' },
//                 bottom: { style: 'thin' },
//                 right: { style: 'thin' }
//             };
//             if (j == DataToExport[6].length) {
//                 worksheet.getCell(colorCell[13 + i] + (CutIndex + j).toString()).fill = {
//                     type: 'pattern',
//                     pattern: 'solid',
//                     fgColor: { argb: 'c5d9f1' },
//                 };
//             }
//         }
//     }

//     let PolishRowData = DataToExport[7].map(item => Object.values(item));
//     let PolishIndex = CutIndex + DataToExport[6].length + 4
//     let PolishMergeCellIndex = PolishIndex - 1
//     worksheet.mergeCells('N' + PolishMergeCellIndex + ':R' + PolishMergeCellIndex);
//     worksheet.getCell('N' + PolishMergeCellIndex).value = 'Polish Proposal';
//     worksheet.getCell('N' + PolishMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
//     worksheet.getCell('N' + PolishMergeCellIndex).font = {
//         bold: true,
//     };
//     worksheet.addTable({
//         name: 'PolishProposal',
//         ref: 'N' + PolishIndex,
//         headerRow: true,
//         totalsRow: false,
//         style: {
//             theme: '',
//             showRowStripes: true,
//             showFirstColumn: true,
//         },
//         columns: [
//             { name: 'Polish Detail' },
//             { name: 'Pcs' },
//             { name: 'Weight' },
//             { name: 'Carat%' },
//             { name: 'Price%' },
//         ],
//         rows: PolishRowData,
//     });

//     // background color and fill border
//     for (let i = 0; i < 5; i++) {
//         worksheet.getCell('N' + (PolishIndex - 1).toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell('N' + (PolishIndex - 1).toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         worksheet.getCell(colorCell[13 + i] + PolishIndex.toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell(colorCell[13 + i] + PolishIndex.toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         for (let j = 1; j <= DataToExport[7].length; j++) {
//             worksheet.getCell(colorCell[13 + i] + (PolishIndex + j).toString()).border = {
//                 top: { style: 'thin' },
//                 left: { style: 'thin' },
//                 bottom: { style: 'thin' },
//                 right: { style: 'thin' }
//             };
//             if (j == DataToExport[7].length) {
//                 worksheet.getCell(colorCell[13 + i] + (PolishIndex + j).toString()).fill = {
//                     type: 'pattern',
//                     pattern: 'solid',
//                     fgColor: { argb: 'c5d9f1' },
//                 };
//             }
//         }
//     }

//     let SymRowData = DataToExport[8].map(item => Object.values(item));
//     let SymIndex = PolishIndex + DataToExport[7].length + 4
//     let SymMergeCellIndex = SymIndex - 1
//     worksheet.mergeCells('N' + SymMergeCellIndex + ':R' + SymMergeCellIndex);
//     worksheet.getCell('N' + SymMergeCellIndex).value = 'Symmentry Proposal';
//     worksheet.getCell('N' + SymMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
//     worksheet.getCell('N' + SymMergeCellIndex).font = {
//         bold: true,
//     };
//     worksheet.addTable({
//         name: 'PolishProposal',
//         ref: 'N' + SymIndex,
//         headerRow: true,
//         totalsRow: false,
//         style: {
//             theme: '',
//             showRowStripes: true,
//             showFirstColumn: true,
//         },
//         columns: [
//             { name: 'Polish Detail' },
//             { name: 'Pcs' },
//             { name: 'Weight' },
//             { name: 'Carat%' },
//             { name: 'Price%' },
//         ],
//         rows: SymRowData,
//     });

//     // background color and fill border
//     for (let i = 0; i < 5; i++) {
//         worksheet.getCell('N' + (SymIndex - 1).toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell('N' + (SymIndex - 1).toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         worksheet.getCell(colorCell[13 + i] + SymIndex.toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell(colorCell[13 + i] + SymIndex.toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         for (let j = 1; j <= DataToExport[8].length; j++) {
//             worksheet.getCell(colorCell[13 + i] + (SymIndex + j).toString()).border = {
//                 top: { style: 'thin' },
//                 left: { style: 'thin' },
//                 bottom: { style: 'thin' },
//                 right: { style: 'thin' }
//             };
//             if (j == DataToExport[8].length) {
//                 worksheet.getCell(colorCell[13 + i] + (SymIndex + j).toString()).fill = {
//                     type: 'pattern',
//                     pattern: 'solid',
//                     fgColor: { argb: 'c5d9f1' },
//                 };
//             }
//         }
//     }

//     let FloRowData = DataToExport[9].map(item => Object.values(item));
//     let FloIndex = SymIndex + DataToExport[8].length + 4
//     let FloMergeCellIndex = FloIndex - 1
//     worksheet.mergeCells('N' + FloMergeCellIndex + ':R' + FloMergeCellIndex);
//     worksheet.getCell('N' + FloMergeCellIndex).value = 'Floro Proposal';
//     worksheet.getCell('N' + FloMergeCellIndex).alignment = { vertical: 'middle', horizontal: 'center' };
//     worksheet.getCell('N' + FloMergeCellIndex).font = {
//         bold: true,
//     };
//     worksheet.addTable({
//         name: 'PolishProposal',
//         ref: 'N' + FloIndex,
//         headerRow: true,
//         totalsRow: false,
//         style: {
//             theme: '',
//             showRowStripes: true,
//             showFirstColumn: true,
//         },
//         columns: [
//             { name: 'Polish Detail' },
//             { name: 'Pcs' },
//             { name: 'Weight' },
//             { name: 'Carat%' },
//             { name: 'Price%' },
//         ],
//         rows: FloRowData,
//     });

//     // background color and fill border
//     for (let i = 0; i < 5; i++) {
//         worksheet.getCell('N' + (FloIndex - 1).toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell('N' + (FloIndex - 1).toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         worksheet.getCell(colorCell[13 + i] + FloIndex.toString()).fill = {
//             type: 'pattern',
//             pattern: 'solid',
//             fgColor: { argb: 'f0f0f0' },
//         };
//         worksheet.getCell(colorCell[13 + i] + FloIndex.toString()).border = {
//             top: { style: 'thin' },
//             left: { style: 'thin' },
//             bottom: { style: 'thin' },
//             right: { style: 'thin' }
//         };
//         for (let j = 1; j <= DataToExport[9].length; j++) {
//             worksheet.getCell(colorCell[13 + i] + (FloIndex + j).toString()).border = {
//                 top: { style: 'thin' },
//                 left: { style: 'thin' },
//                 bottom: { style: 'thin' },
//                 right: { style: 'thin' }
//             };
//             if (j == DataToExport[9].length) {
//                 worksheet.getCell(colorCell[13 + i] + (FloIndex + j).toString()).fill = {
//                     type: 'pattern',
//                     pattern: 'solid',
//                     fgColor: { argb: 'c5d9f1' },
//                 };
//             }
//         }
//     }

//     workbook.xlsx.writeBuffer().then(function (buffer) {
//         let xlsData = Buffer.concat([buffer]);
//         res.writeHead(200, {
//             'Content-Length': Buffer.byteLength(xlsData),
//             'Content-Type': 'application/vnd.ms-excel',
//             'Content-disposition': 'attachment;filename=data.xlsx',
//         }).end(xlsData);
//     });

// }

exports.CompareViewExcelDownload = async (req, res) => {

    let ColData = [];
    let Data = JSON.parse(req.body.DataRow);
    // console.log(Data)
    ColData = JSON.parse(req.body.ColData);
    const Excel = require('exceljs');

    // const mergeRange = { s: { c: 2, r: 0 }, e: { c: 7, r: 0 } }; // C1:H1

    // Merge the cells
    // worksheet1['!merges'] = worksheet['!merges'] || [];
    // worksheet1['!merges'].push(mergeRange);

    const workbook = new Excel.Workbook();
    const worksheet1 = workbook.addWorksheet("Summary");
    const colorCell1 = [
        "A",
        "B",
        "C",
        "D",
        "E",
        "F",
        "G",
        "H",
        "I",
        "J",
        "K",
        "L",
        "M",
        "N",
        "O",
        "P",
        "Q",
        "R",
        "S",
        "T",
        "U",
        "V",
        "W",
        "X",
        "Y",
        "Z",
        "AA",
        "AB",
        "AC",
        "AD",
        "AE",
        "AF",
        "AG",
        "AG",
        "AH",
        "AI",
        "AJ",
        "AK",
        "AL",
    ];
    worksheet1.columns = [
        { key: "Emp" },
        { key: "Psc" },
        { key: "Q.Same" },
        { key: "Q.Same%" },
        { key: "Q.Plus" },
        { key: "Q.Plus%" },
        { key: "Q.Down" },
        { key: "Q.Down%" },
        { key: "C.Same" },
        { key: "C.Same%" },
        { key: "C.Plus" },
        { key: "C.Plus%" },
        { key: "C.Down" },
        { key: "C.Down%" },
        { key: "CU.Same" },
        { key: "CU.Same%" },
        { key: "CU.Plus" },
        { key: "CU.Plus%" },
        { key: "CU.Down" },
        { key: "CU.Down%" },
        { key: "P.Same" },
        { key: "P.Same%" },
        { key: "P.Plus" },
        { key: "P.Plus%" },
        { key: "P.Down" },
        { key: "P.Down%" },
        { key: "S.Same" },
        { key: "S.Same%" },
        { key: "S.Plus" },
        { key: "S.Plus%" },
        { key: "S.Down" },
        { key: "S.Down%" },
        { key: "F.Same" },
        { key: "F.Same%" },
        { key: "F.Plus" },
        { key: "F.Plus%" },
        { key: "F.Down" },
        { key: "F.Down%" },
    ];


    // let shapeTableIndex = totalTableIndex + request.recordsets[1].length + 3;
    // let shapeTableMergeIndex = shapeTableIndex - 1;

    worksheet1.mergeCells(
        "AG" + 1 + ":AL" + 1
    );

    worksheet1.getCell("AG" + 1).value = "FLO";
    worksheet1.getCell("AG" + 1).alignment = {
        vertical: "middle",
        horizontal: "center",
    };
    worksheet1.getCell("AG" + 1).font = {
        bold: true,
    };


    worksheet1.mergeCells(
        "C" + 1 + ":H" + 1
    );
    worksheet1.getCell("C" + 1).value = "QUALITY";
    worksheet1.getCell("C" + 1).alignment = {
        vertical: "middle",
        horizontal: "center",
    };
    worksheet1.getCell("C" + 1).font = {
        bold: true,
    };

    worksheet1.mergeCells(
        "I" + 1 + ":N" + 1
    );
    worksheet1.getCell("I" + 1).value = "COLOR";
    worksheet1.getCell("I" + 1).alignment = {
        vertical: "middle",
        horizontal: "center",
    };
    worksheet1.getCell("I" + 1).font = {
        bold: true,
    };

    worksheet1.mergeCells(
        "O" + 1 + ":T" + 1
    );
    worksheet1.getCell("O" + 1).value = "CUT";
    worksheet1.getCell("O" + 1).alignment = {
        vertical: "middle",
        horizontal: "center",
    };
    worksheet1.getCell("O" + 1).font = {
        bold: true,
    };

    worksheet1.mergeCells(
        "U" + 1 + ":Z" + 1
    );
    worksheet1.getCell("U" + 1).value = "POLISH";
    worksheet1.getCell("U" + 1).alignment = {
        vertical: "middle",
        horizontal: "center",
    };
    worksheet1.getCell("U" + 1).font = {
        bold: true,
    };

    worksheet1.mergeCells(
        "AA" + 1 + ":AF" + 1
    );
    worksheet1.getCell("AA" + 1).value = "SYMMETRY";
    worksheet1.getCell("AA" + 1).alignment = {
        vertical: "middle",
        horizontal: "center",
    };
    worksheet1.getCell("AA" + 1).font = {
        bold: true,
    };

    worksheet1.getRow(2).values = [
        "Emp",
        "Psc",
        "Q.Same",
        "Q.Same%",
        "Q.Plus",
        "Q.Plus%",
        "Q.Down",
        "Q.Down%",
        "C.Same",
        "C.Same%",
        "C.Plus",
        "C.Plus%",
        "C.Down",
        "C.Down%",
        "CU.Same",
        "CU.Same%",
        "CU.Plus",
        "CU.Plus%",
        "CU.Down",
        "CU.Down%",
        "P.Same",
        "P.Same%",
        "P.Plus",
        "P.Plus%",
        "P.Down",
        "P.Down%",
        "S.Same",
        "S.Same%",
        "S.Plus",
        "S.Plus%",
        "S.Down",
        "S.Down%",
        "F.Same",
        "F.Same%",
        "F.Plus",
        "F.Plus%",
        "F.Down",
        "F.Down%",

    ];

    Data.forEach(function (record, index) {
        // console.log(Data.length)
        ID1 = index + 1;
        ID2 = index + 2;

        // console.log(ID1)
        // console.log(ID2)

        let sum =
            "IF(P" + ID2 + '=0,"",O' + ID2 + "-(O" + ID2 + "*(P" + ID2 + "/100)))";
        // console.log(sum)
        // console.log(record)
        worksheet1.addRow({
            // NO: index + 1,
            Emp: record["EMP_CODE"],
            Psc: record["PKT"],
            // QUALITY

            'Q.Same': record["QUASAME"],
            'Q.Same%': record["QUASAMEPER"].toFixed(2),
            'Q.Plus': record["QUAPLUS"],
            'Q.Plus%': record["QUAPLUSPER"].toFixed(2),
            'Q.Down': record["QUADOWN"],
            'Q.Down%': record["QUADOWNPER"].toFixed(2),
            // COL
            'C.Same': record["COLSAME"],
            'C.Same%': record["COLSAMEPER"].toFixed(2),
            'C.Plus': record["COLPLUS"],
            'C.Plus%': record["COLPLUSPER"].toFixed(2),
            'C.Down': record["COLDOWN"],
            'C.Down%': record["COLDOWNPER"].toFixed(2),
            // CUT
            'CU.Same': record["CUTSAME"],
            'CU.Same%': record["CUTSAMEPER"].toFixed(2),
            'CU.Plus': record["CUTPLUS"],
            'CU.Plus%': record["CUTPLUSPER"].toFixed(2),
            'CU.Down': record["CUTDOWN"],
            'CU.Down%': record["CUTDOWNPER"].toFixed(2),
            // POLISH
            'P.Same': record["POLSAME"],
            'P.Same%': record["POLSAMEPER"].toFixed(2),
            'P.Plus': record["POLPLUS"],
            'P.Plus%': record["POLPLUSPER"].toFixed(2),
            'P.Down': record["POLDOWN"],
            'P.Down%': record["POLDOWNPER"].toFixed(2),
            // SYM
            'S.Same': record["SYMSAME"],
            'S.Same%': record["SYMSAMEPER"].toFixed(2),
            'S.Plus': record["SYMPLUS"],
            'S.Plus%': record["SYMPLUSPER"].toFixed(2),
            'S.Down': record["SYMDOWN"],
            'S.Down%': record["SYMDOWNPER"].toFixed(2),
            // Flo
            'F.Same': record["FLSAME"],
            'F.Same%': record["FLSAMEPER"].toFixed(2),
            'F.Plus': record["FLPLUS"],
            'F.Plus%': record["FLPLUSPER"].toFixed(2),
            'F.Down': record["FLDOWN"],
            'F.Down%': record["FLDOWNPER"].toFixed(2),

        });
    })

    // FOR A TOTAL
    const cellToUpdate = worksheet1.getCell('A' + (Data.length + 3));
    cellToUpdate.value = 'Total';
    cellToUpdate.font = {
        bold: true
    }

    const cellToUpdat = worksheet1.getCell('B' + (Data.length + 3));
    cellToUpdat.value = { formula: `SUM(B3:B${Data.length + 2})`, result: undefined };

    const cellToUpda = worksheet1.getCell('C' + (Data.length + 3));
    cellToUpda.value = { formula: `SUM(C3:C${Data.length + 2})`, result: undefined };

    const cellToUpd = worksheet1.getCell('D' + (Data.length + 3));
    cellToUpd.value = { formula: `SUM(C${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellToUp = worksheet1.getCell('E' + (Data.length + 3));
    cellToUp.value = { formula: `SUM(E3:E${Data.length + 2})`, result: undefined };

    const cellToU = worksheet1.getCell('F' + (Data.length + 3));
    cellToU.value = { formula: `SUM(E${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellTo = worksheet1.getCell('G' + (Data.length + 3));
    cellTo.value = { formula: `SUM(G3:G${Data.length + 2})`, result: undefined };

    const cellT = worksheet1.getCell('H' + (Data.length + 3));
    cellT.value = { formula: `SUM(G${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellUPDATE = worksheet1.getCell('I' + (Data.length + 3));
    cellUPDATE.value = { formula: `SUM(I3:I${Data.length + 2})`, result: undefined };

    const cellUPDAT = worksheet1.getCell('J' + (Data.length + 3));
    cellUPDAT.value = { formula: `SUM(I${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellUPDA = worksheet1.getCell('K' + (Data.length + 3));
    cellUPDA.value = { formula: `SUM(K3:K${Data.length + 2})`, result: undefined };

    const cellUPD = worksheet1.getCell('L' + (Data.length + 3));
    cellUPD.value = { formula: `SUM(K${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellUP = worksheet1.getCell('M' + (Data.length + 3));
    cellUP.value = { formula: `SUM(M3:M${Data.length + 2})`, result: undefined };

    const cellU = worksheet1.getCell('N' + (Data.length + 3));
    cellU.value = { formula: `SUM(M${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellRENDERCHANGE = worksheet1.getCell('O' + (Data.length + 3));
    cellRENDERCHANGE.value = { formula: `SUM(O3:O${Data.length + 2})`, result: undefined };

    const cellRENDERCHANG = worksheet1.getCell('P' + (Data.length + 3));
    cellRENDERCHANG.value = { formula: `SUM(O${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellRENDERCHAN = worksheet1.getCell('Q' + (Data.length + 3));
    cellRENDERCHAN.value = { formula: `SUM(Q3:Q${Data.length + 2})`, result: undefined };

    const cellRENDERCHA = worksheet1.getCell('R' + (Data.length + 3));
    cellRENDERCHA.value = { formula: `SUM(Q${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellRENDERCH = worksheet1.getCell('S' + (Data.length + 3));
    cellRENDERCH.value = { formula: `SUM(S3:S${Data.length + 2})`, result: undefined };

    const cellRENDERC = worksheet1.getCell('T' + (Data.length + 3));
    cellRENDERC.value = { formula: `SUM(S${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellRENDER = worksheet1.getCell('U' + (Data.length + 3));
    cellRENDER.value = { formula: `SUM(U3:U${Data.length + 2})`, result: undefined };

    const cellRENDE = worksheet1.getCell('V' + (Data.length + 3));
    cellRENDE.value = { formula: `SUM(U${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellREND = worksheet1.getCell('W' + (Data.length + 3));
    cellREND.value = { formula: `SUM(W3:W${Data.length + 2})`, result: undefined };

    const cellREN = worksheet1.getCell('X' + (Data.length + 3));
    cellREN.value = { formula: `SUM(W${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellRE = worksheet1.getCell('Y' + (Data.length + 3));
    cellRE.value = { formula: `SUM(Y3:Y${Data.length + 2})`, result: undefined };

    const cellR = worksheet1.getCell('Z' + (Data.length + 3));
    cellR.value = { formula: `SUM(Y${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellRECOVER = worksheet1.getCell('AA' + (Data.length + 3));
    cellRECOVER.value = { formula: `SUM(AA3:AA${Data.length + 2})`, result: undefined };

    const cellRECOVE = worksheet1.getCell('AB' + (Data.length + 3));
    cellRECOVE.value = { formula: `SUM(AA${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellRECOV = worksheet1.getCell('AC' + (Data.length + 3));
    cellRECOV.value = { formula: `SUM(AC3:AC${Data.length + 2})`, result: undefined };

    const cellRECO = worksheet1.getCell('AD' + (Data.length + 3));
    cellRECO.value = { formula: `SUM(AC${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellREC = worksheet1.getCell('AE' + (Data.length + 3));
    cellREC.value = { formula: `SUM(AE3:AE${Data.length + 2})`, result: undefined };

    const cellFORMULA = worksheet1.getCell('AF' + (Data.length + 3));
    cellFORMULA.value = { formula: `SUM(AE${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellFORMUL = worksheet1.getCell('AG' + (Data.length + 3));
    cellFORMUL.value = { formula: `SUM(AG3:AG${Data.length + 2})`, result: undefined };

    const cellFORMU = worksheet1.getCell('AH' + (Data.length + 3));
    cellFORMU.value = { formula: `SUM(AG${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellFORM = worksheet1.getCell('AI' + (Data.length + 3));
    cellFORM.value = { formula: `SUM(AI3:AI${Data.length + 2})`, result: undefined };

    const cellFOR = worksheet1.getCell('AJ' + (Data.length + 3));
    cellFOR.value = { formula: `SUM(AI${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    const cellFO = worksheet1.getCell('AK' + (Data.length + 3));
    cellFO.value = { formula: `SUM(AK3:AK${Data.length + 2})`, result: undefined };

    const cellF = worksheet1.getCell('AL' + (Data.length + 3));
    cellF.value = { formula: `SUM(AK${Data.length + 3}/B${Data.length + 3}*100)`, result: undefined };

    worksheet1.views = [
        { state: 'frozen', xSplit: 1, ySplit: 2 }
    ];

    workbook.xlsx.writeBuffer().then(function (buffer) {
        let xlsData = Buffer.concat([buffer]);
        res
            .writeHead(200, {
                "Content-Length": Buffer.byteLength(xlsData),
                "Content-Type": "application/vnd.ms-excel",
                "Content-disposition": "attachment;filename=CompareView.xlsx",
            })
            .end(xlsData);
    });
}

exports.HeliumHisView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('C_CODE', sql.Int, parseInt(req.body.C_CODE))
                request.input('Q_CODE', sql.Int, parseInt(req.body.Q_CODE))
                request.input('CUT_CODE', sql.Int, parseInt(req.body.CUT_CODE))
                request.input('POL_CODE', sql.Int, parseInt(req.body.POL_CODE))
                request.input('SYM_CODE', sql.Int, parseInt(req.body.SYM_CODE))
                request.input('F_CARAT', sql.Decimal(10, 3), parseFloat(req.body.F_CARAT));
                request.input('T_CARAT', sql.Decimal(10, 3), parseFloat(req.body.T_CARAT));
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)

                request = await request.execute('usp_VWFrmHeliumHisView');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.MakHistoryUnLock = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('ISLOCK', sql.Bit, req.body.ISLOCK)

                request = await request.execute('usp_MakHistoryUnLock');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.VWGalFileView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('CHK_CODE', sql.VarChar(5), req.body.CHK_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)

                request = await request.execute('USP_VWGalFileView');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.VWGalFileViewTotal = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('CHK_CODE', sql.VarChar(5), req.body.CHK_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)

                request = await request.execute('USP_VWGalFileViewTotal');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.VWFrmDeleteView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('DEPT', sql.Bit, req.body.DEPT)

                request = await request.execute('usp_VWFrmDeleteView');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.BrkViewDisp = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7999), req.body.L_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('Srno', sql.Int, parseInt(req.body.Srno))
                request.input('SrnoTo', sql.Int, parseInt(req.body.SrnoTo))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('R_CODE', sql.Int, parseInt(req.body.R_CODE))
                request.input('TYPE', sql.VarChar(5), req.body.TYPE)
                if (req.body.I_DATE) { request.input('I_DATE', sql.DateTime2, new Date(req.body.I_DATE)) }
                if (req.body.DateTo) { request.input('DateTo', sql.DateTime2, new Date(req.body.DateTo)) }
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('CEMP_CODE', sql.VarChar(8), req.body.CEMP_CODE)
                request.input('MEMP_CODE', sql.VarChar(8), req.body.MEMP_CODE)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))

                request = await request.execute('VW_BrkViewDisp');
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

exports.Brkconfirm = async (req, res) => {
    // jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    //     if (err) {
    //         res.sendStatus(401);
    //     } else {
    //         const TokenData = await authData;
    //         let RapArr = req.body;
    //         console.log(RapArr)
    //         let status = true;
    //         let ErrorRap = []
    //         for (let i = 0; i < RapArr.length; i++) {
    //             try {

    //                 var request = new sql.Request();

    //                 request.input('L_CODE', sql.VarChar(8), RapArr[i].L_CODE)
    //                 request.input('SRNO', sql.Int, parseInt(RapArr[i].SRNO))
    //                 request.input('BTAG', sql.VarChar(2), RapArr[i].BTAG)
    //                 request.input('DETID', sql.Int, parseInt(RapArr[i].DETID))
    //                 request.input('IS_CNF', sql.Bit, RapArr[i].IS_CNF)
    //                 request.input('CF_USER', sql.VarChar(10), RapArr[i].CF_USER)
    //                 request.input('CF_COMP', sql.VarChar(10), RapArr[i].CF_COMP)
    //                 request.input('PARTY_CNF', sql.Bit, RapArr[i].PARTY_CNF)
    //                 request.input('CHK_CNF', sql.Bit, RapArr[i].CHK_CNF)
    //                 request.input('IUSER', sql.VarChar(10), RapArr[i].IUSER)
    //                 request.input('BRKCARAT', sql.Decimal(10, 3), parseFloat(req.body.BRKCARAT))
    //                 request.input('DEPT_CODE', sql.VarChar(5), RapArr[i].DEPT_CODE)

    //                 request = await request.execute('usp_Brkconfirm');

    //             } catch (err) {
    //                 status = false;
    //                 ErrorRap.push(RapArr[i])
    //             }
    //         }
    //         res.json({ success: status, data: ErrorRap })
    //     }
    // });
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(8), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('BTAG', sql.VarChar(2), req.body.BTAG)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                request.input('IS_CNF', sql.Bit, req.body.IS_CNF)
                request.input('CF_USER', sql.VarChar(10), req.body.CF_USER)
                request.input('CF_COMP', sql.VarChar(10), req.body.CF_COMP)
                request.input('PARTY_CNF', sql.Bit, req.body.PARTY_CNF)
                request.input('CHK_CNF', sql.Bit, req.body.CHK_CNF)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('BRKCARAT', sql.Decimal(10, 3), parseFloat(req.body.BRKCARAT))
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)

                request = await request.execute('usp_Brkconfirm');

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

exports.GrdOutEntBcode = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('IP', sql.VarChar(25), IP)

                request = await request.execute('USP_GrdOutEntBcode');

                if (request.recordset) {
                    let text = '';
                    for (var i = 0; i < request.recordset.length; i++) {
                        text = text + request.recordset[i].BCODE + '\n';
                    }
                    // let filePath = path.join(__dirname, `../../Public/file/BarCode/BARCODE-${Date.now()}.txt`);
                    let timeStemp = Date.now();
                    let filePath = `${process.env.BARCODE_PATH}/BARCODE-${timeStemp}.txt`;
                    let PrinterPath = request.recordsets[1][0] ? request.recordsets[1][0].PRN_NAME : '';
                    let message = '';

                    if (!PrinterPath) {
                        message = "Printer Path not found."
                    }
                    let printPath = `${process.env.PRINT_BARCODE_PATH}\\BARCODE-${timeStemp}.txt`
                    let printCmd = `Type ${printPath} > ${PrinterPath}`;
                    await fs.writeFile(filePath, text, async (err) => {
                        if (err)
                            console.log("File write error", err);
                        else {
                            console.log(printCmd)
                            await exec(printCmd, async (error, stdout, stderr) => {
                                if (error) {
                                    console.log("error ::::", error)
                                    message = error.message;
                                }
                                if (stderr) {
                                    console.log("stderr ::::", stderr)

                                    message = stderr;
                                }
                                console.log("stdout ::::", stdout)

                                message = stdout;
                                fs.unlinkSync(filePath);
                            });
                        }
                    });

                    res.json({ success: 1, data: request.recordset, message: message })
                } else {
                    res.json({ success: 1, data: "Not Found" })
                }

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.VWGrdView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('P_CODE', sql.VarChar(7991), req.body.P_CODE)
                request.input('TYP', sql.VarChar(10), req.body.TYP)
                if (req.body.DATE) { request.input('DATE', sql.DateTime2, new Date(req.body.DATE)) }
                if (req.body.DATETO) { request.input('DATETO', sql.DateTime2, new Date(req.body.DATETO)) }
                request.input('E_CODE', sql.VarChar(7999), req.body.E_CODE)
                request.input('COMP', sql.VarChar(10), req.body.COMP)
                request.input('CEMP_CODE', sql.VarChar(7991), req.body.CEMP_CODE)

                request = await request.execute('usp_VWGrdView');

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

exports.VWLockView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('PRC_CODE', sql.VarChar(5), req.body.PRC_CODE)
                request.input('F_SRNO', sql.Int, parseInt(req.body.F_SRNO))
                request.input('T_SRNO', sql.Int, parseInt(req.body.T_SRNO))
                request.input('TYPE', sql.VarChar(10), req.body.TYPE)
                request.input('ISLOCK', sql.Bit, req.body.ISLOCK)


                if (req.body.EMP_CODE) { request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE) }

                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                if (req.body.SELECTMODE) { request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE) }
                if (req.body.UL_USER) { request.input('UL_USER', sql.VarChar(15), req.body.UL_USER) }
                if (req.body.UL_COMP) { request.input('UL_COMP', sql.VarChar(15), req.body.UL_COMP) }
                if (req.body.PEMP_CODE) { request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE) }
                if (req.body.GRP) { request.input('GRP', sql.VarChar(5), req.body.GRP) }

                request = await request.execute('USP_VWLockView');

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.VWFrmLockViewPktLock = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('F_SRNO', sql.Int, parseInt(req.body.F_SRNO))
                request.input('T_SRNO', sql.Int, parseInt(req.body.T_SRNO))
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('PRC_CODE', sql.VarChar(5), req.body.PRC_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)

                request = await request.execute('Usp_VWLockViewPktLock');
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

exports.PrdViewDisp = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                request.input('L_CODE', sql.VarChar(7999), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('PRDTYPE', sql.VarChar(5), req.body.PRDTYPE)
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)
                request.input('DISPTYP', sql.VarChar(10), req.body.DISPTYP)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('TYP', sql.VarChar(5), req.body.TYP)
                request.input('F_CARAT', sql.Numeric(10, 3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10, 3), req.body.T_CARAT)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('GRP', sql.VarChar(10), req.body.GRP)

                request = await request.execute('VW_PrdViewDisp');
                if (request.recordsets) {
                    res.json({ success: 1, data: request.recordsets })
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

exports.PrdViewSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            let RapArr = req.body;
            let status = true;
            let ErrorRap = []
            let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
            for (let i = 0; i < RapArr.length; i++) {
                try {
                    var request = new sql.Request();

                    request.input('L_CODE', sql.VarChar(10), RapArr[i].L_CODE)
                    request.input('SRNO', sql.Int, parseInt(RapArr[i].SRNO))
                    request.input('TAG', sql.VarChar(5), RapArr[i].TAG)
                    request.input('PTAG', sql.VarChar(5), RapArr[i].PTAG)
                    request.input('PRDTYPE', sql.VarChar(5), RapArr[i].PRDTYPE)
                    request.input('PLANNO', sql.Int, parseInt(RapArr[i].PLANNO))
                    request.input('EMP_CODE', sql.VarChar(5), RapArr[i].EMP_CODE)
                    request.input('FPLNSEL', sql.Bit, RapArr[i].FPLNSEL)
                    request.input('ISLOCK', sql.Bit, RapArr[i].ISLOCK)
                    request.input('IUSER', sql.VarChar(10), RapArr[i].IUSER)
                    request.input('IP', sql.VarChar(25), IP)

                    request = await request.execute('usp_PrdViewSave');
                } catch (err) {
                    status = false;
                    ErrorRap.push(err)
                }
            }
            res.json({ success: status, data: ErrorRap })

        }
    });
}

exports.PrdViewPrint = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input("L_CODE", sql.VarChar(7999), req.body.L_CODE)
                request.input("SRNO", sql.Int, req.body.SRNO)
                request.input("TAG", sql.VarChar(5), req.body.TAG)
                request.input("EMP_CODE", sql.VarChar(5), req.body.EMP_CODE)
                request.input("PRDTYPE", sql.VarChar(5), req.body.PRDTYPE)
                request.input("PNT", sql.Int, req.body.PNT)
                request.input("TYP", sql.VarChar(5), req.body.TYP)
                request.input("F_CARAT", sql.Numeric(10, 3), req.body.F_CARAT)
                request.input("T_CARAT", sql.Numeric(10, 3), req.body.T_CARAT)
                request.input("S_CODE", sql.VarChar(5), req.body.S_CODE)
                request.input('GRP', sql.VarChar(10), req.body.GRP)

                request = await request.execute('USP_PrdViewPrint');

                if (request.recordsets) {
                    res.json({ success: 1, data: request.recordsets })
                } else {
                    res.json({ success: 0, data: '' })
                }

            } catch (err) {
                res.json({ success: 0, data: err })

            }
        }
    })
}

exports.PrdViewUnLock = async (req, res) => {

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

                request = await request.execute('VW_PrdViewUnLock');
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

exports.InnPrcTallyFGrdView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('L_CODE', sql.VarChar(5), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('PRC_CODE', sql.VarChar(7), req.body.PRC_CODE)
                request.input('M_CODE', sql.VarChar(7), req.body.M_CODE)
                request.input('EMP_CODE', sql.VarChar(7), req.body.EMP_CODE)
                request.input('IUSER', sql.VarChar(20), req.body.IUSER)
                request.input('ICOMP', sql.VarChar(20), req.body.ICOMP)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request = await request.execute('usp_FrmInnerPrcTallyFGrdDisp');
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

exports.InnPrcTallySGrdView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(3), req.body.TAG)
                request.input('L_CODE1', sql.VarChar(7991), req.body.L_CODE1)
                request.input('SRNO1', sql.Int, parseInt(req.body.SRNO1))
                request.input('SRNOTO1', sql.Int, parseInt(req.body.SRNOTO1))
                request.input('TAG1', sql.VarChar(5), req.body.TAG1)
                request.input('PRC_CODE', sql.VarChar(7), req.body.PRC_CODE)
                request.input('M_CODE', sql.VarChar(7), req.body.M_CODE)
                request.input('EMP_CODE', sql.VarChar(7), req.body.EMP_CODE)
                request.input('IUSER', sql.VarChar(20), req.body.IUSER)
                request.input('ICOMP', sql.VarChar(20), req.body.ICOMP)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request = await request.execute('usp_FrmInnerPrcTallySGrdDisp');

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

exports.TallyTotal = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('TYPE', sql.VarChar(5), req.body.TYPE)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)

                request = await request.execute('usp_TallyTotal');

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

exports.InnerPrcTallyClear = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('ICOMP', sql.VarChar(50), req.body.ICOMP)
                request.input('IUSER', sql.VarChar(50), req.body.IUSER)
                request.input('TYPE', sql.VarChar(5), req.body.TYPE)
                request.input('PRC_CODE', sql.VarChar(10), req.body.PRC_CODE)
                request.input('M_CODE', sql.VarChar(10), req.body.M_CODE)
                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)

                request = await request.execute('usp_FrmInnerPrcTallyClear');

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.InnerPrcTally = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                request.input('IUSER', sql.VarChar(20), req.body.IUSER)
                request.input('ICOMP', sql.VarChar(20), req.body.ICOMP)
                request.input('DEPT_CODE', sql.VarChar(10), req.body.DEPT_CODE)

                request = await request.execute('usp_FrmInnerPrcTally');

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.InnerPrcChkPcTally = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(2), req.body.TAG)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('IUSER', sql.VarChar(20), req.body.IUSER)
                request.input('ICOMP', sql.VarChar(20), req.body.ICOMP)
                request.input('EMP_CODE', sql.VarChar(7), req.body.EMP_CODE)
                request.input('M_CODE', sql.VarChar(5), req.body.M_CODE)
                request.input('PRC_CODE', sql.VarChar(5), req.body.PRC_CODE)

                request = await request.execute('ufn_FrmInnerPrcChkPcTally');

                if (request.output) {
                    res.json({ success: 1, data: request.output[''] })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                Decimal
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DimPktViewDisp = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)
                request.input('F_SRNO', sql.Int, parseInt(req.body.F_SRNO))
                request.input('T_SRNO', sql.Int, parseInt(req.body.T_SRNO))
                request.input('F_CARAT', sql.Decimal(10, 3), parseFloat(req.body.F_CARAT))
                request.input('T_CARAT', sql.Decimal(10, 3), parseFloat(req.body.T_CARAT))
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)

                request = await request.execute('usp_VWDimPktViewDisp');

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

exports.DimPktViewDispTotal = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)
                request.input('F_SRNO', sql.Int, parseInt(req.body.F_SRNO))
                request.input('T_SRNO', sql.Int, parseInt(req.body.T_SRNO))
                request.input('F_CARAT', sql.Decimal(10, 3), parseFloat(req.body.F_CARAT))
                request.input('T_CARAT', sql.Decimal(10, 3), parseFloat(req.body.T_CARAT))
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)

                request = await request.execute('usp_VWDimPktViewDispTotal');

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

exports.DimPktViewDispPrint = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('P_CODE', sql.VarChar(5), req.body.P_CODE)
                request.input('F_SRNO', sql.Int, parseInt(req.body.F_SRNO))
                request.input('T_SRNO', sql.Int, parseInt(req.body.T_SRNO))
                request.input('F_CARAT', sql.Decimal(10, 3), parseFloat(req.body.F_CARAT))
                request.input('T_CARAT', sql.Decimal(10, 3), parseFloat(req.body.T_CARAT))
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)

                request = await request.execute('usp_VWDimPktViewDispPrint');

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

exports.InnerPrcPktViewPrint = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('INO', sql.Int, parseInt(req.body.INO))
                request.input('INOTO', sql.Int, parseInt(req.body.INOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('M_CODE', sql.VarChar(5), req.body.M_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('usp_VWInnerPrcPktViewPrint');

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

exports.InnerPrcPrcViewPrint = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('INO', sql.Int, parseInt(req.body.INO))
                request.input('INOTO', sql.Int, parseInt(req.body.INOTO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('M_CODE', sql.VarChar(5), req.body.M_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('OEMP_CODE', sql.VarChar(5), req.body.OEMP_CODE)
                request.input('PRC_TYP', sql.VarChar(5), req.body.PRC_TYP)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)

                request = await request.execute('usp_VWInnerPrcPrcViewPrint');

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

exports.GalStkView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('F_SRNO', sql.Int, parseInt(req.body.F_SRNO))
                request.input('T_SRNO', sql.Int, parseInt(req.body.T_SRNO))
                request.input('F_INO', sql.Int, parseInt(req.body.F_INO))
                request.input('T_INO', sql.Int, parseInt(req.body.T_INO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) } else { request.input('F_DATE', sql.DateTime2, null) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) } else { request.input('T_DATE', sql.DateTime2, null) }
                request.input('F_TIME', sql.Time, req.body.F_TIME)
                request.input('T_TIME', sql.Time, req.body.T_TIME)
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('TYPE', sql.VarChar(10), req.body.TYPE)

                request = await request.execute('USP_VWFrmGalStkView');

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

exports.GalStkViewTotal = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('F_SRNO', sql.Int, parseInt(req.body.F_SRNO))
                request.input('T_SRNO', sql.Int, parseInt(req.body.T_SRNO))
                request.input('F_INO', sql.Int, parseInt(req.body.F_INO))
                request.input('T_INO', sql.Int, parseInt(req.body.T_INO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) } else { request.input('F_DATE', sql.DateTime2, null) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) } else { request.input('T_DATE', sql.DateTime2, null) }
                request.input('F_TIME', sql.Time, req.body.F_TIME)
                request.input('T_TIME', sql.Time, req.body.T_TIME)
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('TYPE', sql.VarChar(10), req.body.TYPE)

                request = await request.execute('USP_VWFrmGalStkViewTotal');

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

exports.GalStkUpd = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('TYP', sql.VarChar(10), req.body.TYP)
                request.input('ISDBC', sql.Bit, req.body.ISDBC)
                request.input('ISOK', sql.Bit, req.body.ISOK)
                if (req.body.I_DATE) { request.input('I_DATE', sql.DateTime2, new Date(req.body.I_DATE)) } else { request.input('I_DATE', sql.DateTime2, null) }
                request.input('INO', sql.Int, parseInt(req.body.INO))
                request.input('RNO', sql.Int, parseInt(req.body.RNO))
                request.input('COL', sql.VarChar(10), req.body.COL)

                request = await request.execute('USP_GalStkUpd');

                res.json({ success: 1, data: '' })

            } catch (err) {

                res.json({ success: 0, data: err })
            }
        }
    });
}


exports.DiscussView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('FINO', sql.Int, parseInt(req.body.FINO))
                request.input('TINO', sql.Int, parseInt(req.body.TINO))
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) } else { request.input('F_DATE', sql.DateTime2, null) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) } else { request.input('T_DATE', sql.DateTime2, null) }
                request.input('F_TIME', sql.Time, req.body.F_TIME)
                request.input('T_TIME', sql.Time, req.body.T_TIME)
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)
                request.input('TPROC_CODE', sql.VarChar(5), req.body.TPROC_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)

                request = await request.execute('USP_VWFrmOpPrcView');

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

exports.DiscussViewTot = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('FINO', sql.Int, parseInt(req.body.FINO))
                request.input('TINO', sql.Int, parseInt(req.body.TINO))
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) } else { request.input('F_DATE', sql.DateTime2, null) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) } else { request.input('T_DATE', sql.DateTime2, null) }
                request.input('F_TIME', sql.Time, req.body.F_TIME)
                request.input('T_TIME', sql.Time, req.body.T_TIME)
                request.input('SELECTMODE', sql.VarChar(10), req.body.SELECTMODE)
                request.input('TPROC_CODE', sql.VarChar(5), req.body.TPROC_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)

                request = await request.execute('USP_VWFrmOpPrcViewTot');

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

exports.CompEntView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('ID', sql.Int, parseInt(req.body.ID))
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('MACH_NAME', sql.VarChar(50), req.body.MACH_NAME)
                request.input('COMPLAIN', sql.VarChar(200), req.body.COMPLAIN)
                request.input('STATUS', sql.VarChar(5), req.body.STATUS)
                request.input('EUSER', sql.VarChar(10), req.body.EUSER)
                request.input('ECOMP', sql.VarChar(30), req.body.ECOMP)
                request.input('TYPE', sql.VarChar(10), req.body.TYPE)

                request = await request.execute('usp_VWFrmITCompEntView');

                if (req.body.TYPE == "SHOW") {
                    if (request.recordset) {
                        res.json({ success: 1, data: request.recordset })
                    } else {
                        res.json({ success: 0, data: "Not Found" })
                    }
                } else {
                    res.json({ success: 1, data: '' })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.FrmAutoTrfAct = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('TYPE', sql.VarChar(5), req.body.TYPE)
                if (req.body.I_DATE) { request.input('I_DATE', sql.DateTime2, new Date(req.body.I_DATE)) } else { request.input('I_DATE', sql.DateTime2, null) }
                request.input('PRC_CODE', sql.VarChar(10), req.body.PRC_CODE)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('ISLEAVE', sql.Bit, req.body.ISLEAVE)

                request = await request.execute('usp_FrmAutoTrfAct');

                if (req.body.TYPE == "SHOW" || req.body.TYPE == "CMB") {
                    if (request.recordset) {
                        res.json({ success: 1, data: request.recordset })
                    } else {
                        res.json({ success: 0, data: "Not Found" })
                    }
                } else {
                    res.json({ success: 1, data: '' })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.InnerLivePktView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('M_CODE', sql.VarChar(5), req.body.M_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(5), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(5), req.body.Q_CODE)
                request.input('F_CARAT', sql.Decimal(10, 3), parseFloat(req.body.F_CARAT))
                request.input('T_CARAT', sql.Decimal(10, 3), parseFloat(req.body.T_CARAT))
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) } else { request.input('F_DATE', sql.DateTime2, null) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) } else { request.input('T_DATE', sql.DateTime2, null) }
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('CUT_CODE', sql.Int, parseInt(req.body.CUT_CODE))
                request.input('POL_CODE', sql.Int, parseInt(req.body.POL_CODE))
                request.input('SYM_CODE', sql.Int, parseInt(req.body.SYM_CODE))
                request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('TYP', sql.VarChar(5), req.body.TYP)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)

                request = await request.execute('usp_VWFrmInnerLivePktView');

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

exports.InnerLivePktViewLock = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('M_CODE', sql.VarChar(5), req.body.M_CODE)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(5), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(5), req.body.Q_CODE)
                request.input('F_CARAT', sql.Numeric(10, 3), parseFloat(req.body.F_CARAT));
                request.input('T_CARAT', sql.Numeric(10, 3), parseFloat(req.body.T_CARAT));
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('CUT_CODE', sql.Int, parseInt(req.body.CUT_CODE))
                request.input('POL_CODE', sql.Int, parseInt(req.body.POL_CODE))
                request.input('SYM_CODE', sql.Int, parseInt(req.body.SYM_CODE))
                request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE)
                request.input('PEMP_CODE', sql.VarChar(5), req.body.PEMP_CODE)
                request.input('TYP', sql.VarChar(5), req.body.TYP)
                request.input('ISLOCK', sql.Bit, req.body.ISLOCK)
                request.input('UL_USER', sql.VarChar(10), req.body.UL_USER)
                request.input('UL_COMP', sql.VarChar(10), req.body.UL_COMP)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)

                request = await request.execute('usp_InnerLivePktViewLock');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.Pktchk = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseFloat(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseFloat(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)

                request = await request.execute('USP_PKTCHK');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }
            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    })
}

exports.CommonView = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData;
            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)

                request = await request.execute('usp_VWFrmCommonView')

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

exports.VWFrmCommonViewUPD = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                request.input('ISCOMMON', sql.Bit, req.body.ISCOMMON)
                request.input('CUSER', sql.VarChar(10), req.body.CUSER)
                request.input('CCOMP', sql.VarChar(30), req.body.CCOMP)

                request = await request.execute('USP_VWFrmCommonViewUPD')
                console.log(request)
                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.VWFrmValMsgView = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData
            try {
                var request = new sql.Request();

                request.input("L_CODE", sql.VarChar(7991), req.body.L_CODE)
                request.input("SRNO", sql.Int, parseInt(req.body.SRNO))
                request.input("SRNOTO", sql.Int, parseInt(req.body.SRNOTO))
                request.input("EMP_CODE", sql.VarChar(5), req.body.EMP_CODE)
                request.input("TAG", sql.VarChar(5), req.body.TAG)
                if (req.body.M_DATE != '') { request.input('M_DATE', sql.DateTime2, new Date(req.body.M_DATE)) }
                if (req.body.M_DATETO != '') { request.input('M_DATETO', sql.DateTime2, new Date(req.body.M_DATETO)) }
                request.input("PROC_CODE", sql.VarChar(5), req.body.PROC_CODE)
                request.input("GRP", sql.VarChar(5), req.body.GRP)

                request = await request.execute('USP_VWFrmValMsgView')

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

exports.NoticeView = async (req, res) => {
    let filenameArr = []
    var multer = require('multer');
    const storage = multer.diskStorage({
        destination: "public/images/display",
        filename: function (req, file, cb) {
            // console.log(file, cb)
            const orgFileName = file.originalname
            const [ext] = orgFileName.split('.').reverse();

            // let fileName = 'NOTIC.' + ext;
            filenameArr.push(orgFileName)
            cb(null, orgFileName)
        }
    });
    // console.log(filenameArr)
    const upload = multer({ storage: storage }).fields([{ name: "NoticeImage", maxCount: 10000 }]);

    upload(req, res, (err) => {
        // console.log(req, res)

        res.json(filenameArr)
        if (err) throw err; console.log(err)
    });
}

exports.backView = async (req, res) => {
    let filenameArr = []
    var multer = require('multer');
    const storage = multer.diskStorage({
        destination: "/back",
        filename: function (req, file, cb) {
            // console.log(file)
            const orgFileName = file.originalname
            const [ext] = orgFileName.split('.').reverse();

            // let fileName = 'NOTIC.' + ext;
            filenameArr.push(orgFileName)
            cb(null, orgFileName)
        }
    });
    // console.log(filenameArr)
    const upload = multer({ storage: storage }).fields([{ name: "NoticeImage", maxCount: 10000 }]);

    upload(req, res, (err) => {
        // console.log(req, res)

        res.json(filenameArr)
        if (err) throw err; console.log(err)
    });
}


exports.backFill = async (req, res) => {
    const fs = require('fs');

    fs.readdir('public/images/background', (err, files) => {
        // }
        // console.log(files)
        res.json({ success: 1, data: files })
    })

}

exports.NoticeFill = async (req, res) => {
    const fs = require('fs');

    fs.readdir('public/images/display', (err, files) => {
        if (err) {
            console.log(err)
        }
        // console.log(files)
        res.json({ success: 1, data: files })
    })

}


exports.NoticeDelete = async (req, res) => {
    const fs = require('fs')

    // console.log(req.body.img.split("/")[5])
    DelImg = req.body.img.split("/")[5]
    fs.unlink(`public/images/display/${DelImg}`, (err) => {
        if (err) {
            console.log(err)
            res.json({ success: 0, data: err })
            return;
        }
        res.json({ success: 1, Data: 'Deleted' })
        // console.log("Deleted")
    })

}

exports.VWFrmAssSDiffView = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData
            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7999), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO)),
                    request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO)),
                    request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('TYPE', sql.VarChar(5), req.body.TYPE)
                request.input('DIF', sql.Numeric(10, 3), req.body.DIF)
                request.input('TODIF', sql.Numeric(10, 3), req.body.TODIF)

                request = await request.execute('usp_VWFrmAssDiffView')

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 0, dat: 'No Data' })
                }

            } catch (Err) {
                res.json({ success: 0, data: Err })
                console.log(Err)
            }
        }
    })
}

exports.VWFrmAssDiffViewClick = async (req, res) => {
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

                request = await request.execute('usp_VWFrmAssDiffViewClick')

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    })
}

exports.VWHeadChkDiffViewClick = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData
            try {
                var request = new sql.Request();

                request.input("L_CODE", sql.VarChar(7999), req.body.L_CODE)
                request.input("SRNO", sql.Int, parseInt(req.body.SRNO))
                request.input("SRNOTO", sql.Int, parseInt(req.body.SRNOTO))
                request.input("TAG", sql.VarChar(2), req.body.TAG)
                if (req.body.I_DATE) { request.input('I_DATE', sql.DateTime2, new Date(req.body.I_DATE)) }
                if (req.body.DATETO) { request.input('DATETO', sql.DateTime2, new Date(req.body.DATETO)) }
                request.input("MEMP_CODE", sql.VarChar(10), req.body.MEMP_CODE)

                // console.log(req.body)

                request = await request.execute("usp_VWHeadChkDiffViewClick")

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.VWHeadChkDiffView = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData
            try {

                var request = new sql.Request();

                request.input("L_CODE", sql.VarChar(7999), req.body.L_CODE)
                request.input("SRNO", sql.Int, parseInt(req.body.SRNO))
                request.input("SRNOTO", sql.Int, parseInt(req.body.SRNOTO))
                request.input("TAG", sql.VarChar(2), req.body.TAG)
                if (req.body.I_DATE) { request.input('I_DATE', sql.DateTime, new Date(req.body.I_DATE)) }
                if (req.body.DATETO) { request.input('DATETO', sql.DateTime, new Date(req.body.DATETO)) }
                request.input("MEMP_CODE", sql.VarChar(10), req.body.MEMP_CODE)
                request.input("SELECTMODE", sql.VarChar(5), req.body.SELECTMODE)

                request = await request.execute("usp_VWHeadChkDiffView")

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 0, dat: 'No Data' })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
                // console.log(err)
            }
        }
    })
}

exports.VWHeadChkDiffViewClickDet = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData;
            try {
                var request = new sql.Request();

                request.input("L_Code", sql.VarChar(10), req.body.L_CODE)
                request.input("SRNO", sql.Int, parseInt(req.body.SRNO))
                request.input("DETNO", sql.Int, parseInt(req.body.DETNO))
                request.input("TAG", sql.VarChar(2), req.body.TAG)
                request.input("BTAG", sql.VarChar(5), req.body.BTAG)

                request = await request.execute("usp_VWHeadChkDiffViewClickDet")

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                // console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    })
}

exports.VWFrmLockViewClick = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData;
            try {
                var request = new sql.Request();

                request.input("L_CODE", sql.VarChar(10), req.body.L_CODE)
                request.input("SRNO", sql.Int, parseInt(req.body.SRNO))
                request.input("TAG", sql.VarChar(5), req.body.TAG)

                request = await request.execute('USP_VWFrmLockViewClick')

                res.json({ success: 1, data: request.recordset })


            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.RapDisCompView = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                const pool = await RapServerConn();
                let request = await pool.request()

                request.input("RAPTYPE", sql.VarChar(5), req.body.RAPTYPE)
                request.input("RTYPE", sql.VarChar(5), req.body.RTYPE)
                request.input("S_Code", sql.VarChar(5), req.body.S_Code)
                request.input("SIZE", sql.Int, parseInt(req.body.SIZE))
                request.input("V_CODE", sql.Int, parseInt(req.body.V_CODE))

                request = await request.execute('usp_RapDisCompView')

                res.json({ success: 1, data: request.recordsets })

            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.AttUpdView = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('USERID', sql.Int, parseInt(req.body.USERID))
                if (req.body.SDATE) { request.input('SDATE', sql.DateTime, new Date(req.body.SDATE)) }
                if (req.body.TDATE) { request.input('TDATE', sql.DateTime, new Date(req.body.TDATE)) }
                request.input('TYP', sql.VarChar(10), req.body.TYP)

                request = await request.execute('usp_AttUpdView')

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.LeaveEntGrpEmpFil = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401)
        } else {
            const TokenData = await authData;
            try {
                var request = new sql.Request();
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('PROC_CODE', sql.Int, parseInt(req.body.PROC_CODE))
                request.input('CAT', sql.VarChar(5), req.body.CAT)
                request.input('TYPE', sql.VarChar(5), req.body.TYPE)
                request.input('EGRP', sql.VarChar(5), req.body.EGRP)
                request = await request.execute("usp_LeaveEntGrpEmpFil")

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.LeaveEntSave = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                request.input('ID', sql.Int, parseInt(req.body.ID))
                request.input('GRP', sql.VarChar(10), req.body.GRP)
                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('REASON', sql.VarChar(500), req.body.REASON)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime, new Date(req.body.T_DATE)) }
                request.input('STATUS', sql.VarChar(10), req.body.STATUS)
                request.input('TYP', sql.VarChar(5), req.body.TYP)
                request.input('F_TIME', sql.VarChar, req.body.F_TIME)
                request.input('T_TIME', sql.VarChar, req.body.T_TIME)
                console.log
                // if (req.body.F_TIME) { request.input('F_TIME', sql.DateTime, new Date(req.body.F_TIME)) }
                // if (req.body.T_TIME) { request.input('T_TIME', sql.DateTime, new Date(req.body.T_TIME)) }
                request.input('IUSER', sql.VarChar(20), req.body.IUSER)
                request.input('ICOMP', sql.VarChar(20), req.body.ICOMP)
                request.input('TRPID', sql.Int, parseInt(req.body.TRPID))

                // console.log(req.body)
                request = await request.execute('usp_LeaveEntSave')
                // console.log(request)
                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.VWLeaveEntDisp = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input("USER_ID", sql.VarChar(7), req.body.USER_ID)
                request.input("GRP", sql.VarChar(7), req.body.GRP)
                if (req.body.DATE) { request.input('DATE', sql.DateTime, new Date(req.body.DATE)) }
                request.input("ID", sql.Int, parseInt(req.body.ID))
                // console.log(req.body)
                request = await request.execute('usp_VWLeaveEntDisp')
                // console.log(request)

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.LeaveEntUpdate = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            try {
                var request = new sql.Request();

                request.input('ID', sql.Int, parseInt(req.body.ID))
                request.input('STATUS', sql.VarChar(10), req.body.STATUS)
                request.input('TYP', sql.VarChar(5), req.body.TYP)
                request.input('IUSER', sql.VarChar(20), req.body.IUSER)
                request.input('ICOMP', sql.VarChar(20), req.body.ICOMP)
                request.input('TRPID', sql.Int, parseInt(req.body.TRPID))

                request = await request.execute('usp_LeaveEntUpdate')

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.LeaveEntDelete = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            try {
                var request = new sql.Request();

                request.input('ID', sql.Int, parseInt(req.body.ID))

                request = await request.execute('usp_LeaveEntDelete')

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.VWFrmRouReqView = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            try {
                var request = new sql.Request();

                request.input("GRP", sql.VarChar(10), req.body.GRP)

                request = await request.execute('usp_VWFrmRouReqView')

                res.json({ success: 1, data: request.recordset })
            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    })
}

exports.VWFrmRouReqEntSave = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {

            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('GRP', sql.VarChar(10), req.body.GRP)
                request.input('MRK', sql.Int, parseInt(req.body.MRK))
                request.input('OLIMIT', sql.Int, parseInt(req.body.OLIMIT))
                request.input('REQ', sql.Int, parseInt(req.body.REQ))
                request.input('AVG_PKT', sql.Int, parseInt(req.body.AVG_PKT))
                request.input('T_1', sql.Int, parseInt(req.body.T_1))
                request.input('PLUS', sql.Int, parseInt(req.body.PLUS))
                request.input('COLNAME', sql.VarChar(2), req.body.COLNAME)

                request = request.execute('usp_VWFrmRouReqEntSave')

                res.json({ success: 1, data: '' })


            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.VWFrmRouReqPrint = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            try {

                const TokenData = await authData;

                var request = new sql.Request();

                request.input("GRP", sql.VarChar(10), req.body.GRP)

                request = await request.execute('usp_VWFrmRouReqPrint')

                res.json({ success: 1, data: request.recordset })
            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    })
}



exports.VWFrmGrdLivePktView = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            try {

                const TokenData = await authData;

                var request = new sql.Request();

                request.input("L_CODE", sql.VarChar(7991), req.body.L_CODE)

                request = await request.execute('USP_VWFrmGrdLivePktView')
                // console.log(request)

                res.json({ success: 1, data: request.recordsets })
            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.EmpLabourCertiPen = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            try {

                const TokenData = await authData;

                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime, new Date(req.body.T_DATE)) }
                request.input('P_CODE', sql.VarChar(7991), req.body.P_CODE)
                request.input('GHATM_CODE', sql.VarChar(7991), req.body.GHATM_CODE)
                request.input('GHATEMP_CODE', sql.VarChar(7991), req.body.GHATEMP_CODE)
                request.input('GHATPRC_CODE', sql.VarChar(7991), req.body.GHATPRC_CODE)
                request.input('ENT_TYPE', sql.VarChar(10), req.body.ENT_TYPE)
                request = await request.execute('RP_EmpLabourCertiPen')

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })
}

exports.VWFrmAssEntView = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            try {

                const TokenData = await authData;

                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('GRP', sql.VarChar(10), req.body.GRP)
                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime, new Date(req.body.T_DATE)) }
                request.input('SELECT', sql.VarChar(10), req.body.SELECT)
                request.input("ISCOMP", sql.Bit, req.body.ISCOMP)
                // console.log(req.body)
                request = await request.execute('USP_VWFrmAssEntView')

                // console.log(JSON.stringify(request.recordsets))
                res.json({ success: 1, data: request.recordsets })

            } catch (Err) {
                res.json({ success: 0, data: err })
            }
        }
    })
}

exports.VWEmpLabourCertiPenTot = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            try {
                const TokenData = await authData;

                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('P_CODE', sql.VarChar(7991), req.body.P_CODE)
                request.input('GHATM_CODE', sql.VarChar(7991), req.body.GHATM_CODE)
                request.input('GHATEMP_CODE', sql.VarChar(7991), req.body.GHATEMP_CODE)
                request.input('GHATPRC_CODE', sql.VarChar(7991), req.body.GHATPRC_CODE)
                request.input('ENT_TYPE', sql.VarChar(10), req.body.ENT_TYPE)

                request = await request.execute('USP_VWEmpLabourCertiPenTot')

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    })
}
exports.VWEmpLabourCertiPenPrn = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            try {
                const TokenData = await authData;

                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime, new Date(req.body.T_DATE)) }
                request.input('P_CODE', sql.VarChar(7991), req.body.P_CODE)
                request.input('GHATM_CODE', sql.VarChar(7991), req.body.GHATM_CODE)
                request.input('GHATEMP_CODE', sql.VarChar(7991), req.body.GHATEMP_CODE)
                request.input('GHATPRC_CODE', sql.VarChar(7991), req.body.GHATPRC_CODE)
                request.input('ENT_TYPE', sql.VarChar(10), req.body.ENT_TYPE)

                request = await request.execute('USP_VWEmpLabourCertiPenPrn')

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    })
}

exports.VWFrmAssEntViewUpd = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            try {
                const TokenData = await authData;

                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(10), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(10), req.body.TAG)
                request.input('ISCHECK', sql.Bit, req.body.ISCHECK)

                request = await request.execute('USP_VWFrmAssEntViewUpd')
                res.json({ success: 1, data: '' })

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }

    })

}


exports.VWEmpLabourCertiPenDetPrint = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            req.sendStatus(401)
        } else {
            try {
                const TokenData = await authData;

                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(5), req.body.TAG)
                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime, new Date(req.body.T_DATE)) }
                request.input('P_CODE', sql.VarChar(7991), req.body.P_CODE)
                request.input('GHATM_CODE', sql.VarChar(7991), req.body.GHATM_CODE)
                request.input('GHATEMP_CODE', sql.VarChar(7991), req.body.GHATEMP_CODE)
                request.input('GHATPRC_CODE', sql.VarChar(7991), req.body.GHATPRC_CODE)
                request.input('ENT_TYPE', sql.VarChar(10), req.body.ENT_TYPE)

                console.log(req.body)
                request = await request.execute('USP_VWEmpLabourCertiPenDetPrint')

                res.json({ success: 1, data: request.recordsets })

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    })
}

exports.GrdAnalazerView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('FSRNO', sql.Int, parseInt(req.body.FSRNO))
                request.input('TSRNO', sql.Int, parseInt(req.body.TSRNO))
                request.input('TAG', sql.VarChar(2), req.body.TAG)
                request.input('TYP', sql.VarChar(10), req.body.TYP)
                if (req.body.F_DATE != '') { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE != '') { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('PNT', sql.Int, parseInt(req.body.PNT))


                request = await request.execute('usp_VWGrdAnalazerView');
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

exports.BrkHistoryDisp = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('TAG', sql.VarChar(2), req.body.TAG)


                request = await request.execute('usp_BrkHistoryDisp');
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

exports.PrdViewExport = async (req, res) => {
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
                request.input('TYP', sql.VarChar(5), req.body.TYP)
                request.input('F_CARAT', sql.Decimal(10, 3), parseFloat(req.body.F_CARAT));
                request.input('T_CARAT', sql.Decimal(10, 3), parseFloat(req.body.T_CARAT));
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('GRP', sql.VarChar(10), req.body.GRP)

                // console.log(req.body)
                request = await request.execute('USP_PrdViewExport');
                // console.log(request.recordsets)
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 1, data: "" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });

}

exports.PredictionViewExcelDownload = async (req, res) => {

    let ColData = [];
    let Data = JSON.parse(req.body.DataRow);
    ColData = JSON.parse(req.body.ColData);
    const Excel = require('exceljs');

    const workbook = new Excel.Workbook();
    const worksheet1 = workbook.addWorksheet("Summary");

    worksheet1.columns = [
        { key: "dummy" },
        { key: "Lot" },
        { key: "SrNo" },
        { key: "Emp" },
        { key: "PTag" },
        { key: "I.Crt" },
        { key: "Shp" },
        { key: "Qua" },
        { key: "Col" },
        { key: "Mc1" },
        { key: "Mc2" },
        { key: "Mc3" },
        { key: "Crt" },
        { key: "Cut" },
        { key: "Pol" },
        { key: "Sym" },
        { key: "Flo" },
        { key: "Inc" },
        { key: "Shd" },
        { key: "LB/LC" },
        { key: "Dis" },
        { key: "Rap" },
        { key: "Rate" },
        { key: "Amt" },
    ];

    worksheet1.getRow(2).values = [
        null,
        "Lot",
        "SrNo",
        "Emp",
        "PTag",
        "I.Crt",
        "Shp",
        "Qua",
        "Col",
        "Mc1",
        "Mc2",
        "Mc3",
        "Crt",
        "Cut",
        "Pol",
        "Sym",
        "Flo",
        "Inc",
        "Shd",
        "LB/LC",
        "Dis",
        "Rap",
        "Rate",
        "Amt",
    ];
    worksheet1.getColumn("Shp").width = 5;
    worksheet1.getColumn("Qua").width = 5;
    worksheet1.getColumn("Col").width = 5;
    worksheet1.getColumn("Mc1").width = 5;
    worksheet1.getColumn("Mc2").width = 5;
    worksheet1.getColumn("Mc3").width = 5;
    worksheet1.getColumn("Cut").width = 5;
    worksheet1.getColumn("Pol").width = 5;
    worksheet1.getColumn("Sym").width = 5;

    Data.forEach(function (record, index) {
        // console.log(record["Lot"],)
        worksheet1.addRow({
            Lot: record["Lot"],
            SrNo: record["SrNo"],
            Emp: record["Emp"],
            PTag: record["PTag"],
            "I.Crt": record["I.Crt"],
            Shp: record["Shp"],
            Qua: record["Qua"],
            Col: record["Col"],
            Mc1: record["MC1"],
            Mc2: record["MC2"],
            Mc3: record["MC3"],
            Crt: record["Crt"],
            Cut: record["Cut"],
            Pol: record["Pol"],
            Sym: record["Sym"],
            Flo: record["Flo"],
            Inc: record["Inc"],
            Shd: record["Shd"],
            "LB/LC": record["LB/LC"],
            Dis: record["Dis"],
            Rap: record["Rap"],
            Rate: record["Rate"],
            Amt: record["Amt"]
        });
    });

    worksheet1.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            if (colNumber !== 1) {
                cell.border = {
                    top: { style: "thin" },
                    left: { style: "thin" },
                    bottom: { style: "thin" },
                    right: { style: "thin" },
                };
            }
        });
    });

    let arr = []

    worksheet1.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        let lotCell = row.getCell("Lot");
        let srNoCell = row.getCell("SrNo");

        if (lotCell.value === null && srNoCell.value === null) {
            arr.push(rowNumber)
        }
    })
    for (let i = 0; i < arr.length; i++) {
        if (arr[i] !== 1) {
            worksheet1.mergeCells("N" + arr[i] + ":W" + arr[i]);
            worksheet1.mergeCells("B" + arr[i] + ":E" + arr[i]);
            worksheet1.mergeCells("G" + arr[i] + ":L" + arr[i]);
        }
    }
    let newarr = []
    let crtarr = []
    let Rarr = []
    Data.forEach(function (record, index) {

        if (record["Lot"] === null && record["SrNo"] === null) {
            worksheet1.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                if (rowNumber > 1) {
                    let lotCell = row.getCell("Lot");
                    let srNoCell = row.getCell("SrNo");
                    let icrtCell = row.getCell("I.Crt");
                    let CrtCell = row.getCell("Crt");
                    let AmtCell = row.getCell("Amt");

                    if (lotCell.value === null && srNoCell.value === null) {
                        var icrtCellValue = icrtCell.value;
                        var amtCellValue = AmtCell.value;
                        var rCell = row.getCell("R");
                        rCell.value = Math.floor(amtCellValue / icrtCellValue)
                        Rarr.push(rCell.value)
                        icrtCell.fill = {
                            type: "pattern",
                            pattern: "solid",
                            fgColor: { argb: "faf364" }
                        };
                        CrtCell.fill = {
                            type: "pattern",
                            pattern: "solid",
                            fgColor: { argb: "faf364" }
                        }
                        AmtCell.fill = {
                            type: "pattern",
                            pattern: "solid",
                            fgColor: { argb: "faf364" }
                        }

                        rCell.fill = {
                            type: "pattern",
                            pattern: "solid",
                            fgColor: { argb: "faf364" }
                        }
                        rCell.alignment = {
                            vertical: "middle",
                            horizontal: "center",
                        }
                    }
                }
            });



        }
    });

    let Amtarr = []
    Data.forEach(function (record, index) {
        if (record["Lot"] === null && record["SrNo"] === null) {
            crtarr.push(record["Crt"])
            newarr.push(record["I.Crt"])
            Amtarr.push(record["Amt"])
        }
    })
    const sum = newarr.reduce(function (a, b) {
        return a + b;
    });
    const Csum = crtarr.reduce(function (a, b) {
        return a + b;
    });
    const Asum = Amtarr.reduce(function (a, b) {
        return a + b;
    });
    // console.log(Asum)
    const RDIV = Math.floor(Asum / sum)
    // const RDiv

    const dataLength = Data.length;
    const rowNumber = dataLength + 3;


    worksheet1.getCell("F" + (Data.length + 3)).value = sum;
    worksheet1.getCell("F" + (Data.length + 3)).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "faf364" }
    }
    worksheet1.getCell("F" + (Data.length + 3)).alignment = {
        vertical: "middle",
        horizontal: "center",
    }
    worksheet1.getCell("M" + (Data.length + 3)).value = Csum;
    worksheet1.getCell("M" + (Data.length + 3)).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "faf364" }
    }
    worksheet1.getCell("M" + (Data.length + 3)).alignment = {
        vertical: "middle",
        horizontal: "center",
    }
    worksheet1.getCell("X" + (Data.length + 3)).value = Asum;
    worksheet1.getCell("X" + (Data.length + 3)).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "faf364" }
    }
    worksheet1.getCell("X" + (Data.length + 3)).alignment = {
        vertical: "middle",
        horizontal: "center",
    }
    worksheet1.getCell("N" + (Data.length + 3)).value = RDIV;
    worksheet1.getCell("N" + (Data.length + 3)).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "faf364" }
    }
    worksheet1.getCell("N" + (Data.length + 3)).alignment = {
        vertical: "middle",
        horizontal: "center",
    }

    const row = worksheet1.getRow(rowNumber);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber !== 1) {
            cell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                bottom: { style: "thin" },
                right: { style: "thin" },
            };
        }
        row.height = 30;
    });
    worksheet1.mergeCells("N" + rowNumber + ":W" + rowNumber);
    worksheet1.mergeCells("B" + rowNumber + ":E" + rowNumber);
    worksheet1.mergeCells("G" + rowNumber + ":L" + rowNumber);

    workbook.xlsx.writeBuffer().then(function (buffer) {
        let xlsData = Buffer.concat([buffer]);
        res
            .writeHead(200, {
                "Content-Length": Buffer.byteLength(xlsData),
                "Content-Type": "application/vnd.ms-excel",
                "Content-disposition": "attachment;filename=PredictionView.xlsx",
            })
            .end(xlsData);
    });
};
