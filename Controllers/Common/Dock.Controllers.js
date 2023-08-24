const sql = require("mssql");
const jwt = require('jsonwebtoken');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

// -------------------------- For Group ANd Emp Fill ----------------------------

exports.USP_SELECT = async (req, res) => {

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
                    res.json({ success: 2, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.GrpOrEmpFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('USERID', sql.VarChar(10), req.body.USERID)
                request.input('CAT', sql.VarChar(5), req.body.CAT)
                request.input('PROC_CODE', sql.VarChar(5), req.body.PROC_CODE)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('TYP', sql.VarChar(10), req.body.TYP)

                request = await request.execute('USP_GrpOrEmpFill');
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 2, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

// ---------------------------------------------------------------------------------

exports.DockPrcPen = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(15), req.body.DEPT_CODE)
                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('ISURGENT', sql.Bit, req.body.ISURGENT)
                request.input('ISMAN', sql.Bit, req.body.ISMAN)
                request.input('ISEMP', sql.Bit, req.body.ISEMP)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)

                request = await request.execute('usp_DockPrcPen');
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 2, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockProduction = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('GRP', sql.VarChar(7991), req.body.GRP)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))

                request = await request.execute('usp_DockProduction');
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 2, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockCurRoll = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(20), req.body.EMP_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))

                request = await request.execute('usp_DockCurRoll');
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 2, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockPrcPenDet = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(15), req.body.DEPT_CODE)
                request.input('PRC_CODE', sql.VarChar(15), req.body.PRC_CODE)
                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('ISURGENT', sql.Bit, req.body.ISURGENT)
                request.input('ISMAN', sql.Bit, req.body.ISMAN)
                request.input('ISEMP', sql.Bit, req.body.ISEMP)
                request.input('ORDTYP', sql.VarChar(5), req.body.ORDTYP)

                request = await request.execute('usp_DockPrcPenDet');
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 2, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockCurRollDet = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(20), req.body.EMP_CODE)
                request.input('PRC', sql.VarChar(10), req.body.PRC)
                request.input('PNT', sql.Int, parseInt(req.body.PNT))

                request = await request.execute('usp_DockCurRollDet');
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 2, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockDimPartyStk = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(15), req.body.DEPT_CODE)
                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('TYPE', sql.VarChar(7991), req.body.TYPE)

                request = await request.execute('usp_DockDimPartyStk');
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 2, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockGrdStk = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.VarChar(15), req.body.DEPT_CODE)
                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('PEMP_CODE', sql.VarChar(10), req.body.PEMP_CODE)
                request.input('EMP_CODE', sql.VarChar(7991), req.body.EMP_CODE)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))

                request = await request.execute('usp_DockGrdStk');
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 2, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockChkStockTaly = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('GRP', sql.VarChar(7991), req.body.GRP)
                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('CPCS', sql.Int, parseInt(req.body.CPCS))
                request.input('ISCHK', sql.Bit, req.body.ISCHK)
                request.input('TYPE', sql.VarChar(10), req.body.TYPE)
                request.input('ISALL', sql.Bit, req.body.ISALL)
                request.input('CAT', sql.VarChar(5), req.body.CAT)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)

                request = await request.execute('USP_DockChkStockTaly');
                // if (request.recordset) {
                res.json({ success: 1, data: request.recordset })
                // } else {
                //     res.json({ success: 2, data: "Not Found" })
                // }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockRequirement = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(8), req.body.EMP_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input('REQ', sql.Int, parseInt(req.body.REQ))
                request.input('PNT', sql.Int, parseInt(req.body.PNT))
                request.input('TYPE', sql.VarChar(10), req.body.TYPE)

                request = await request.execute('usp_DockRequirement');

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockPelProd = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('PEMP_CODE', sql.VarChar(10), req.body.PEMP_CODE)
                request.input('M_CODE', sql.VarChar(10), req.body.M_CODE)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))

                request = await request.execute('usp_DockPelProd');
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordsets })
                } else {
                    res.json({ success: 2, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockPrdLock = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('GRP', sql.VarChar(7991), req.body.GRP)
                request.input('EMP_CODE', sql.VarChar(100), req.body.EMP_CODE)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))

                request = await request.execute('USP_DockPrdLock');
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 2, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });

}

exports.Dock2nDUpd = async (req, res) => {

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
                request.input('IS_2ND', sql.Bit, req.body.IS_2ND)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))

                request = await request.execute('usp_Dock2nDUpd');

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.Dock2ND = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('EMP_CODE', sql.VarChar(20), req.body.EMP_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))

                request = await request.execute('usp_Dock2ND');
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 2, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockMrkStockTaly = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('GRP', sql.VarChar(7991), req.body.GRP)
                request.input('EMP_CODE', sql.VarChar(10), req.body.EMP_CODE)
                request.input('CPCS', sql.Int, parseInt(req.body.CPCS))
                request.input('ISCHK', sql.Bit, req.body.ISCHK)
                request.input('TYPE', sql.VarChar(10), req.body.TYPE)
                request.input('CAT', sql.VarChar(5), req.body.CAT)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))

                request = await request.execute('usp_DockMrkStockTaly');
                // if (request.recordset) {
                res.json({ success: 1, data: request.recordset })
                // } else {
                //     res.json({ success: 2, data: "Not Found" })
                // }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockFactory = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('PEMP_CODE', sql.VarChar(10), req.body.PEMP_CODE)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('SELECT', sql.VarChar(10), req.body.SELECT)
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))
                request = await request.execute('usp_DockFactory');

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockLotStatus = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('ISURGENT', sql.Bit, req.body.ISURGENT)
                request.input('P_TYP', sql.VarChar(10), req.body.P_TYP)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))
                request = await request.execute('usp_DockLotStatus');

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.DockLotStatusDet = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('ISURGENT', sql.Bit, req.body.ISURGENT)
                request.input('P_TYP', sql.VarChar(10), req.body.P_TYP)
                request.input('COLNAME', sql.VarChar(10), req.body.COLNAME)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))
                request = await request.execute('usp_DockLotStatusDet');

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.ClvToGrdDue = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input("PNT", sql.Int, parseInt(req.body.PNT))
                
                request = await request.execute('usp_ClvToGrdDue');

                res.json({ success: 1, data: request.recordset })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}
