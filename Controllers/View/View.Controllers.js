const sql = require("mssql");
const jwt = require('jsonwebtoken');
const Excel = require('exceljs');
const fs = require('fs');
var path = require('path');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.PricingWrk = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('S_CODE', sql.VarChar(7991), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(7991), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(7991), req.body.Q_CODE)
                request.input('F_CARAT', sql.Numeric(10, 3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10, 3), req.body.T_CARAT)
                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)

                request = await request.execute('VW_PricingWrk');

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

exports.ParcelgWrk = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if (req.body.F_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                if (req.body.T_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                request.input('S_CODE', sql.VarChar(7991), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(7991), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(7991), req.body.Q_CODE)
                request.input('F_CARAT', sql.Numeric(10, 3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10, 3), req.body.T_CARAT)
                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)

                request = await request.execute('VW_ParcelgWrk');

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

exports.PricingWrkDisp = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if (req.body.F_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                if (req.body.T_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                request.input('S_CODE', sql.VarChar(7991), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(7991), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(7991), req.body.Q_CODE)
                request.input('F_CARAT', sql.Numeric(10, 3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10, 3), req.body.T_CARAT)
                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, req.body.DETID)


                request = await request.execute('VW_PricingWrkDisp');

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

exports.ParcelWrkDisp = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                if (req.body.F_DATE) { request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE)) }
                if (req.body.T_DATE) { request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE)) }
                request.input('S_CODE', sql.VarChar(7991), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(7991), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(7991), req.body.Q_CODE)
                request.input('F_CARAT', sql.Numeric(10, 3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10, 3), req.body.T_CARAT)
                request.input('DETID', sql.Int, req.body.DETID)
                if (req.body.COMP_CODE) { request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE) }


                request = await request.execute('VW_ParcelWrkDisp');

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

exports.PricingWrkMperSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;

                request.input('MPER', sql.Numeric(10, 3), req.body.MPER)
                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('PLANNO', sql.Int, parseInt(req.body.PLANNO))
                request.input('PTAG', sql.VarChar(2), req.body.PTAG)
                request.input('MUSER', sql.VarChar(20), req.body.MUSER)
                request.input('MCOMP', sql.VarChar(30), IP)


                request = await request.execute('VW_PricingWrkMperSave');


                res.json({ success: 1, data: '' })
            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.BVView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))

                request = await request.execute('VW_BVView');

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

exports.BidDataView = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))

                request = await request.execute('VW_BidDataView');

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

exports.ColAnalysis = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))
                request.input('S_CODE', sql.VarChar(7991), req.body.S_CODE)
                request.input('C_CODE', sql.VarChar(7991), req.body.C_CODE)
                request.input('Q_CODE', sql.VarChar(7991), req.body.Q_CODE)
                request.input('CT_CODE', sql.VarChar(7991), req.body.CT_CODE)
                request.input('F_CARAT', sql.Numeric(10, 3), req.body.F_CARAT)
                request.input('T_CARAT', sql.Numeric(10, 3), req.body.T_CARAT)
                request.input('FINAL', sql.VarChar(7991), req.body.FINAL)
                request.input('RESULT', sql.VarChar(7991), req.body.RESULT)
                request.input('FLNO', sql.VarChar(7991), req.body.FLNO)
                request.input('MACHFL', sql.VarChar(7991), req.body.MACHFL)
                request.input('COMENT', sql.VarChar(7991), req.body.COMENT)
                request.input('LAB', sql.VarChar(10), req.body.LAB)

                request = await request.execute('VW_ColAnalysis');

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

exports.TenderWin = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('COMP_CODE', sql.VarChar(10), req.body.COMP_CODE)
                request.input('DETID', sql.Int, parseInt(req.body.DETID))

                request = await request.execute('VW_TenderWin');

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

exports.TendarWinSheet = async (req, res) => {

    let Data = JSON.parse(req.body.DataRow);
    let HeaderData = JSON.parse(req.body.HeaderDataRow);
    const Excel = require('exceljs');
    const workbook = new Excel.Workbook();
    const worksheet1 = workbook.addWorksheet("Summary");
    // console.log(HeaderData[0])
    worksheet1.columns = [
        { key: "Blank" },
        { key: "LOT" },
        { key: "RWEIGHT" },
        { key: "SHAP" },
        { key: "CUT" },
        { key: "SIZE" },
        { key: "COLAR" },
        { key: "CLIYARITY" },
        { key: "INC" },
        { key: "FLO" },
        { key: "RAP" },
        { key: "DIS" },
        { key: "AMOUNT" },
        { key: "TAMOUNT" },
        { key: "PAY" },
        { key: "PAYPERCERET" },
        { key: "PER" },
        { key: "LABOUR" },
        { key: "LAB_NAME" },
        { key: "tenshan" },
        { key: "1" },
        { key: "2" },
        { key: "3" },
        { key: "4t" },
        { key: "DN" },
        { key: "USER1" },
        { key: "USER2" },
        { key: "USER3" },
    ];

    const columnA = worksheet1.getColumn('A');
    const columnB = worksheet1.getColumn('B');
    const columnC = worksheet1.getColumn('C');
    const columnD = worksheet1.getColumn('D');
    const columnE = worksheet1.getColumn('E');
    const columnF = worksheet1.getColumn('F');
    const columnG = worksheet1.getColumn('G');
    const columnH = worksheet1.getColumn('H');
    const columnI = worksheet1.getColumn('I');
    const columnJ = worksheet1.getColumn('J');
    const columnK = worksheet1.getColumn('K');
    const columnL = worksheet1.getColumn('L');
    const columnM = worksheet1.getColumn('M');
    const columnN = worksheet1.getColumn('N');
    const columnO = worksheet1.getColumn('O');
    const columnP = worksheet1.getColumn('P');
    const columnQ = worksheet1.getColumn('Q');
    const columnR = worksheet1.getColumn('R');
    const columnS = worksheet1.getColumn('S');
    const columnT = worksheet1.getColumn('T');
    const columnU = worksheet1.getColumn('U');
    const columnV = worksheet1.getColumn('V');
    const columnW = worksheet1.getColumn('W');
    const columnX = worksheet1.getColumn('X');
    const columnY = worksheet1.getColumn('Y');
    const columnZ = worksheet1.getColumn('Z');
    const columnAA = worksheet1.getColumn('AA');
    const columnAB = worksheet1.getColumn('AB');

    columnA.width = 3.71;
    columnB.width = 16;
    columnC.width = 14;
    columnD.width = 7.71;
    columnE.width = 7.71;
    columnF.width = 8.71;
    columnG.width = 9.30;
    columnH.width = 13.71;
    columnI.width = 5.15;
    columnJ.width = 13.71;
    columnK.width = 8.71;
    columnL.width = 5.44;
    columnM.width = 13;
    columnN.width = 15.57;
    columnO.width = 14;
    columnP.width = 20.30;
    columnQ.width = 11.57;
    columnR.width = 17.71;
    columnS.width = 17.71;
    columnT.width = 11.57;
    columnU.width = 11.57;
    columnV.width = 9.15;
    columnW.width = 7.30;
    columnX.width = 11.30;
    columnY.width = 7.17;
    columnZ.width = 9.43;
    columnAA.width = 7.71;
    columnAB.width = 8.43;

    columnC.alignment = { horizontal: 'center', vertical: 'middle' };
    columnB.alignment = { horizontal: 'center', vertical: 'middle' };
    columnD.alignment = { horizontal: 'center', vertical: 'middle' };
    columnT.alignment = { horizontal: 'center', vertical: 'middle' };
    columnS.alignment = { horizontal: 'center', vertical: 'middle' };
    columnN.alignment = { horizontal: 'center', vertical: 'middle' };
    columnO.alignment = { horizontal: 'center', vertical: 'middle' };
    columnP.alignment = { horizontal: 'center', vertical: 'middle' };
    columnQ.alignment = { horizontal: 'center', vertical: 'middle' };
    columnR.alignment = { horizontal: 'center', vertical: 'middle' };
    columnY.alignment = { horizontal: 'center', vertical: 'middle' };
    columnZ.alignment = { horizontal: 'center', vertical: 'middle' };
    columnAA.alignment = { horizontal: 'center', vertical: 'middle' };
    columnAB.alignment = { horizontal: 'center', vertical: 'middle' };

    worksheet1.mergeCells('D2:S6');
    worksheet1.getCell('D2').value = HeaderData[0].HEADER
    worksheet1.getCell('D2').font = { size: 48, bold: true }

    worksheet1.mergeCells('T2:AB6');
    worksheet1.getCell('T2').value = 'COLOR MACHINE COMMENT'
    worksheet1.getCell('T2').font = { size: 24, bold: true }
    worksheet1.getCell('T2').border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
    };

    worksheet1.getRow(3).height = 36;
    worksheet1.getCell('B3').value = 'R-WIGHT'
    worksheet1.getCell('B3').font = { size: 16, bold: true }
    worksheet1.getCell('B3').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8ff2b' }
    };
    worksheet1.getCell('B3').border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
    };

    worksheet1.getCell('C3').value = HeaderData[0].I_CARAT
    worksheet1.getCell('C3').font = { size: 16, bold: true }
    worksheet1.getCell('C3').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8ff2b' }
    };
    worksheet1.getRow(4).height = 36;
    worksheet1.getCell('C3').border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
    };

    worksheet1.getCell('B4').value = 'T-AMOUNT'
    worksheet1.getCell('B4').font = { size: 16, bold: true }
    worksheet1.getCell('B4').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8ff2b' }
    };
    worksheet1.getCell('B4').border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
    };

    worksheet1.getCell('C4').value = HeaderData[0].FAMT
    worksheet1.getCell('C4').font = { size: 16, bold: true }
    worksheet1.getCell('C4').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8ff2b' }
    };
    worksheet1.getCell('C4').border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
    };
    worksheet1.getRow(5).height = 15.75;
    worksheet1.getRow(6).height = 15.75;
    worksheet1.mergeCells('B5:B6');
    worksheet1.getCell('B5').value = 'P-CTS'
    worksheet1.getCell('B5').font = { size: 14, bold: true }
    worksheet1.getCell('B5').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8ff2b' }
    };
    worksheet1.getCell('B5').border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
    };
    worksheet1.mergeCells('C5:C6');
    worksheet1.getCell('C5').value = HeaderData[0].PCRT
    worksheet1.getCell('C5').font = { size: 14, bold: true }
    worksheet1.getCell('C5').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8ff2b' }
    };
    worksheet1.getCell('C5').border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
    };

    let headerRow = worksheet1.getRow(2)
    for (let col = 2; col <= 12; col++) {
        const cell = headerRow.getCell(col);
        cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } },
        };
    }

    worksheet1.getRow(7).values = [
        "",
        "LOT.",
        "R.WEIGHT.",
        "SHAP",
        "CUT",
        "SIZE",
        "COLAR",
        "CLIYARITY",
        "INC",
        "FLO",
        "RAP",
        "DIS",
        "AMOUNT",
        "T-AMOUNT",
        "PAY",
        "PAY PER CERET",
        "PER",
        "LABOUR",
        "GIAIGIHRD",
        "tenshan%",
        "1",
        "2",
        "3",
        "4",
        "DN",
        "VHB",
        "YOGESH",
        "PS",
        "",
    ];
    let header = worksheet1.getRow(7)
    for (let col = 2; col <= 19; col++) {
        const cell = header.getCell(col);
        cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } },
        };
        cell.font = { size: 16, bold: true }
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    }
    for (let col = 20; col <= 28; col++) {
        const cell = header.getCell(col);
        cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } },
        };
        cell.font = { size: 16, bold: false }
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    }

    Data.forEach(function (record, index) {
        let Lot = ''
        if(record["I_CARAT"]){
            Lot = record["SRNO"]
        }else{
            Lot = ''
        }
        worksheet1.addRow({
            Blank: '',
            LOT: Lot,
            RWEIGHT: record["I_CARAT"],
            SHAP: record["S_NAME"],
            CUT: record["CT_NAME"],
            SIZE: record["CARAT"],
            COLAR: record["C_NAME"],
            CLIYARITY: record["Q_NAME"],
            INC: record["IN_NAME"],
            FLO: record["FL_NAME"],
            RAP: record["ORAP"],
            DIS: record["DIS"],
            AMOUNT: record["AMT"],
            TAMOUNT: record["MAMT"],
            PAY: record["FAMT"],
            PAYPERCERET: record["FBID"],
            PER: record["PER"],
            LABOUR: record["PERPAY"],
            LAB_NAME: record["LAB_NAME"],
            Tenshan: record["T_NAME"],
            1: record["FLAT1"],
            2: record["FLAT2"],
            3: record["MED"],
            4: record["HIGH"],
            DN: record["DN"],
            USER1: record["U1"],
            USER2: record["U2"],
            USER3: record["U3"],
        });

    })

    worksheet1.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber > 7) {
                row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    if (colNumber !== 1 && row.values[3]) {
                        cell.border = {
                            top: { style: 'thin', color: { argb: 'FF000000' } },
                            left: { style: 'thin', color: { argb: 'FF000000' } },
                            bottom: { style: 'thin', color: { argb: 'FF000000' } },
                            right: { style: 'thin', color: { argb: 'FF000000' } },
                        };
                    }
                });
            
        }
    })
    function mergeCellsBasedOnCondition(worksheet1, columnName, Datanumber) {
        let startRow = 8;
        let endRow = 8;
        
        worksheet1.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            const value = row.values[Datanumber];
            const nextValue = worksheet1.getCell(`${columnName}${rowNumber + 1}`).value;

            if (value === nextValue) {
                endRow = rowNumber + 1;
            } else {
                if (endRow > startRow) {
                    worksheet1.mergeCells(`${columnName}${startRow}:${columnName}${endRow}`);
                }
                startRow = rowNumber + 1;
                endRow = rowNumber + 1;
            }
        });

        if (endRow > startRow) {
            worksheet1.mergeCells(`${columnName}${startRow}:${columnName}${endRow}`);
        }

    }

    
    mergeCellsBasedOnCondition(worksheet1, 'N', 14);
    mergeCellsBasedOnCondition(worksheet1, 'O', 15);
    mergeCellsBasedOnCondition(worksheet1, 'P', 16);
    mergeCellsBasedOnCondition(worksheet1, 'Q', 17);
    mergeCellsBasedOnCondition(worksheet1, 'R', 18);
    mergeCellsBasedOnCondition(worksheet1, 'S', 19);
    mergeCellsBasedOnCondition(worksheet1, 'T', 20);
    mergeCellsBasedOnCondition(worksheet1, 'Y', 25);
    mergeCellsBasedOnCondition(worksheet1, 'Z', 26);
    mergeCellsBasedOnCondition(worksheet1, 'AA', 27);
    mergeCellsBasedOnCondition(worksheet1, 'AB', 28);

    let value = 0
    let SumOfvalue = 0

    let valueOfO = 0
    let SumOfO = 0

    let valueOfR = 0
    let SumOfR = 0
    worksheet1.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if(rowNumber > 7){
            if(value != row.getCell(`N`).value){
                value = row.getCell(`N`).value
                SumOfvalue += value
            }

            if(valueOfO != row.getCell(`O`).value){
                valueOfO = row.getCell(`O`).value
                SumOfO += valueOfO
            }

            if(valueOfR != row.getCell(`R`).value){
                valueOfR = row.getCell(`R`).value
                SumOfR+= valueOfR
            }
        }
    })

    const dataLength = Data.length;
    let rowNumber =0;
    for (let i = 0; i < dataLength + 1; i++) {
        rowNumber = i + 8
    }
    worksheet1.mergeCells(`C${rowNumber}:C${rowNumber + 1}`);
    worksheet1.mergeCells(`N${rowNumber}:N${rowNumber + 1}`);
    worksheet1.mergeCells(`O${rowNumber}:O${rowNumber + 1}`);
    worksheet1.mergeCells(`R${rowNumber}:R${rowNumber + 1}`);
    worksheet1.getCell(`R${rowNumber}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8ff2b' }
    };
    worksheet1.getCell(`N${rowNumber}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8ff2b' }
    };
    worksheet1.getCell(`O${rowNumber}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8ff2b' }
    };
    worksheet1.getCell(`C${rowNumber}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8ff2b' }
    };
    worksheet1.getCell(`C${rowNumber}`).value = HeaderData[0].I_CARAT;
    worksheet1.getCell(`N${rowNumber}`).value = SumOfvalue.toFixed(0);
    worksheet1.getCell(`O${rowNumber}`).value = SumOfO.toFixed(0);
    worksheet1.getCell(`R${rowNumber}`).value = SumOfR.toFixed(0);

    let row = worksheet1.getRow(rowNumber)
    let row1 = worksheet1.getRow(rowNumber + 1)

    for (let col = 2; col <= 28; col++) {
        const cell = row.getCell(col);
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } },
        };
        const cell1 = row1.getCell(col);
        cell1.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } },
        };
      }

    workbook.xlsx.writeBuffer().then(function (buffer) {
        let xlsData = Buffer.concat([buffer]);
        res
            .writeHead(200, {
                "Content-Length": Buffer.byteLength(xlsData),
                "Content-Type": "application/vnd.ms-excel",
                "Content-disposition": "attachment;filename=Tendar-Win.xlsx",
            })
            .end(xlsData);
    });
}
