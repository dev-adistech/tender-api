const sql = require("mssql");
const jwt = require('jsonwebtoken');
const Excel = require('exceljs');

var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

var { RapServerConn } = require('../../config/db/rapsql');

exports.RapParamDiscFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                const pool = await RapServerConn();
                let request = await pool.request()

                request.input('RAPTYPE', sql.VarChar(5), req.body.RAPTYPE)
                request.input('RTYPE', sql.VarChar(2), req.body.RTYPE)
                request.input('SZ_CODE', sql.Int, parseInt(req.body.SZ_CODE))
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('PARAM_NAME', sql.VarChar(50), req.body.PARAM_NAME)
                request.input('FMETER', sql.Decimal(10, 2), parseFloat(req.body.FMETER));
                request.input('TMETER', sql.Decimal(10, 2), parseFloat(req.body.TMETER));
                request.input('CT_CODE', sql.Int, parseInt(req.body.CT_CODE))
                request.input('FL_CODE', sql.Int, parseInt(req.body.FL_CODE))

                request = await request.execute('USP_FrmRapParamDisFill');

                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
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

exports.RapParamDiscSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            try {
                const TokenData = await authData;
                let RapArr = req.body;
                let status = true;
                let ErrorRap = []
                const pool = await RapServerConn();
                let request = await pool.request();

                request.input('RAPTYPE', sql.VarChar(5), req.body.RAPTYPE)
                request.input('RTYPE', sql.VarChar(2), req.body.RTYPE)
                request.input('SZ_CODE', sql.Int, parseInt(req.body.SZ_CODE))
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('C_CODE', sql.Int, parseInt(req.body.C_CODE))
                request.input('PARAM_NAME', sql.VarChar(50), req.body.PARAM_NAME)
                request.input('TMETER', sql.Decimal(10, 2), parseFloat(req.body.TMETER));
                request.input('FMETER', sql.Decimal(10, 2), parseFloat(req.body.FMETER));
                request.input('FL_CODE', sql.Int, parseInt(req.body.FL_CODE))
                request.input('CT_CODE', sql.Int, parseInt(req.body.CT_CODE))
                request.input('Q1', sql.Decimal(10, 2), parseFloat(req.body.Q1 == null ? 0 : req.body.Q1))
                request.input('Q2', sql.Decimal(10, 2), parseFloat(req.body.Q2 == null ? 0 : req.body.Q2))
                request.input('Q3', sql.Decimal(10, 2), parseFloat(req.body.Q3 == null ? 0 : req.body.Q3))
                request.input('Q4', sql.Decimal(10, 2), parseFloat(req.body.Q4 == null ? 0 : req.body.Q4))
                request.input('Q5', sql.Decimal(10, 2), parseFloat(req.body.Q5 == null ? 0 : req.body.Q5))
                request.input('Q6', sql.Decimal(10, 2), parseFloat(req.body.Q6 == null ? 0 : req.body.Q6))
                request.input('Q7', sql.Decimal(10, 2), parseFloat(req.body.Q7 == null ? 0 : req.body.Q7))
                request.input('Q8', sql.Decimal(10, 2), parseFloat(req.body.Q8 == null ? 0 : req.body.Q8))
                request.input('Q9', sql.Decimal(10, 2), parseFloat(req.body.Q9 == null ? 0 : req.body.Q9))
                request.input('Q10', sql.Decimal(10, 2), parseFloat(req.body.Q10 == null ? 0 : req.body.Q10))
                request.input('Q11', sql.Decimal(10, 2), parseFloat(req.body.Q11 == null ? 0 : req.body.Q11))
                request.input('Q12', sql.Decimal(10, 2), parseFloat(req.body.Q12 == null ? 0 : req.body.Q12))
                request.input('Q13', sql.Decimal(10, 2), parseFloat(req.body.Q13 == null ? 0 : req.body.Q13))
                request.input('Q14', sql.Decimal(10, 2), parseFloat(req.body.Q14 == null ? 0 : req.body.Q14))
                request.input('Q15', sql.Decimal(10, 2), parseFloat(req.body.Q15 == null ? 0 : req.body.Q15))
                request.input('Q16', sql.Decimal(10, 2), parseFloat(req.body.Q16 == null ? 0 : req.body.Q16))
                request.input('Q17', sql.Decimal(10, 2), parseFloat(req.body.Q17 == null ? 0 : req.body.Q17))
                request.input('F_CARAT', sql.Decimal(10, 2), parseFloat(req.body.F_CARAT))
                request.input('T_CARAT', sql.Decimal(10, 2), parseFloat(req.body.T_CARAT))

                request = await request.execute('USP_FrmRapParamDisSave');

                res.json({ success: 1, data: request.recordset })
            } catch (err) {
                res.json({ success: 0, data: err })
            }

        }
    });
}

exports.RapParamDiscImport = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            let RapArr = req.body;
            let status = true;
            let ErrorRap = []
            for (let shapeIndex = 0; shapeIndex < RapArr.length; shapeIndex++) {
                for (let i = 0; i < RapArr[shapeIndex].length; i++) {
                    for (let j = 0; j < RapArr[shapeIndex][i].C_NAME.length; j++) {
                        try {

                            const pool = await RapServerConn();
                            let request = await pool.request()

                            request.input('RAPTYPE', sql.VarChar(5), RapArr[shapeIndex][i].RAPTYPE)
                            request.input('RTYPE', sql.VarChar(2), RapArr[shapeIndex][i].RTYPE)
                            request.input('S_NAME', sql.VarChar(25), RapArr[shapeIndex][i].S_NAME)
                            request.input('C_NAME', sql.VarChar(5), RapArr[shapeIndex][i].C_NAME[j])
                            request.input('PARAM_NAME', sql.VarChar(50), RapArr[shapeIndex][i].PARAM_NAME)
                            request.input('FMETER', sql.Numeric(10, 2), RapArr[shapeIndex][i].FMETER)
                            request.input('TMETER', sql.Numeric(10, 2), RapArr[shapeIndex][i].TMETER)
                            request.input('CT_NAME', sql.VarChar(20), RapArr[shapeIndex][i].CT_NAME)
                            request.input('FL_NAME', sql.VarChar(20), RapArr[shapeIndex][i].FL_NAME)

                            request.input('Q1', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q1[j])
                            request.input('Q2', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q2[j])
                            request.input('Q3', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q3[j])
                            request.input('Q4', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q4[j])
                            request.input('Q5', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q5[j])
                            request.input('Q6', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q6[j])
                            request.input('Q7', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q7[j])
                            request.input('Q8', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q8[j])
                            request.input('Q9', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q9[j])
                            request.input('Q10', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q10[j])
                            request.input('Q11', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q11[j])
                            request.input('Q12', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q12[j])
                            request.input('Q13', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q13[j])
                            request.input('Q14', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q14[j])
                            request.input('Q15', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q15[j])
                            request.input('Q16', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q16[j])
                            request.input('Q17', sql.Numeric(10, 2), RapArr[shapeIndex][i].Q17[j])

                            request.input('F_CARAT', sql.Numeric(10, 3), parseFloat(RapArr[shapeIndex][i].F_CARAT))
                            request.input('T_CARAT', sql.Numeric(10, 3), parseFloat(RapArr[shapeIndex][i].T_CARAT))

                            request = await request.execute('USP_FrmRapParamDisImport');

                        } catch (err) {
                            status = false;
                            ErrorRap.push(err)
                        }
                    }
                }
            }
            res.json({ success: status, data: ErrorRap })
        }
    });
}

exports.RapParamDiscExport = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                const pool = await RapServerConn();
                let request = await pool.request()
                request.input('RAPTYPE', sql.VarChar(5), req.body.RAPTYPE)
                request.input('RTYPE', sql.VarChar(2), req.body.RTYPE)
                request.input('PARAM_NAME', sql.VarChar(50), req.body.PARAM_NAME)
                request.input('S_CODE', sql.VarChar(5), req.body.S_CODE)
                request.input('SZ_CODE', sql.VarChar(5), req.body.SZ_CODE)
                request.input('FL_CODE', sql.VarChar(5), req.body.FL_CODE)

                request = await request.execute('USP_FrmRapParamDisExport');

                res.json({ success: 1, data: request.recordset, SZ_NAME: req.body.SZ_NAME, SZ_CODE: req.body.SZ_CODE, TMETER: req.body.TMETER, FMETER: req.body.FMETER })

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.RapParamDiscExcelDownload = async (req, res) => {

    let DataToExport = [];
    let CellIndex = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

    DataToExport = JSON.parse(req.body.DataRow);
    QuaData = JSON.parse(req.body.QuaData)
    shape = req.body.SHAPE
    RTYPE = req.body.RTYPE

    let RObj = [
        { code: 'H', name: 'HRD' },
        { code: 'R', name: 'LOOSE' },
        { code: 'I', name: 'IGI' },
        { code: 'S', name: 'GIA' }
    ]
    let workbook = new Excel.Workbook();

    if (req.body.shapeCount) {
        shapeArr = req.body.shapeArr
        shapeArr = shapeArr.split(",");
        for (let shapeIndex = 0; shapeIndex < DataToExport.length; shapeIndex++) {
            // for (let z = 0; z < req.body.shapeCount; z++) {

            let worksheet = workbook.addWorksheet(shapeArr[shapeIndex])


            //TODAY DATE AND TIME
            var date_ob = new Date();
            var day = ("0" + date_ob.getDate()).slice(-2);
            var month = ("0" + (date_ob.getMonth() + 1)).slice(-2);
            var year = date_ob.getFullYear();
            var hours = date_ob.getHours();
            var minutes = date_ob.getMinutes();
            var dateTime = year + "-" + month + "-" + day + " " + hours + ":" + minutes;

            //set header
            if (RTYPE == 'H') {
                worksheet.mergeCells('B1:R1');
            } else {
                worksheet.mergeCells('B1:R1');
            }
            worksheet.getCell('B1').value = dateTime;
            worksheet.getCell('B1').alignment = { horizontal: 'center', bgColor: { argb: 'FF00FF00' } };
            worksheet.getCell('B1').fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'A0A0A0' },
                font: { bold: true }
            };
            worksheet.getCell('B1').font = {
                bold: true
            };
            worksheet.getCell('B1').border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            if (RTYPE == 'H') {
                worksheet.mergeCells('B2:R2');
            } else {
                worksheet.mergeCells('B2:R2');
            }
            worksheet.getCell('B2').value = `${RObj.find((r) => r.code == req.body.RTYPE ? r.name : '').name}_${req.body.RAPTYPE}`;
            worksheet.getCell('B2').alignment = { horizontal: 'center', bgColor: { argb: 'FF00FF00' } };
            worksheet.getCell('B2').fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'A0A0A0' },
                font: { bold: true }
            };
            worksheet.getCell('B2').font = {
                bold: true
            };
            worksheet.getCell('B2').border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            if (RTYPE == 'H') {
                worksheet.mergeCells('B3:R3');
            } else {
                worksheet.mergeCells('B3:R3');
            }
            worksheet.getCell('B3').value = `${req.body.S_CODE}-${shapeArr[shapeIndex]}-${req.body.FL_NAME}`;
            worksheet.getCell('B3').alignment = { horizontal: 'center', bgColor: { argb: 'FF00FF00' } };
            worksheet.getCell('B3').fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'A0A0A0' },
                font: { bold: true }
            };
            worksheet.getCell('B3').font = {
                bold: true
            };
            worksheet.getCell('B3').border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            if (RTYPE == 'H') {

                let SizeIndex = 5;
                for (let tableIndex = 0; tableIndex < DataToExport[shapeIndex].length; tableIndex++) {

                    worksheet.mergeCells('B' + SizeIndex + ':R' + SizeIndex);
                    worksheet.getCell('B' + SizeIndex).value = `${DataToExport[shapeIndex][tableIndex].SZ_NAME}-${DataToExport[shapeIndex][tableIndex].FMETER}-${DataToExport[shapeIndex][tableIndex].TMETER}`;
                    worksheet.getCell('B' + SizeIndex).alignment = { horizontal: 'center' };
                    worksheet.getCell('B' + SizeIndex).fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'A0A0A0' },
                        font: { bold: true }
                    };
                    worksheet.getCell('B' + SizeIndex).font = {
                        bold: true
                    };

                    SizeIndex = SizeIndex + DataToExport[shapeIndex][0].data.length + 3
                }
                let xlsQua = Object.keys(DataToExport[shapeIndex][0].data[0]).filter((item) => item.startsWith('Q'))
                xlsQua = QuaData.filter((item) => xlsQua.some((d) => item.code == d))
                let rowNameIndex = 6;
                for (let tableIndex = 0; tableIndex < DataToExport[shapeIndex].length; tableIndex++) {
                    for (let i = 0; i < xlsQua.length; i++) {

                        worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value = xlsQua[i].name;
                        worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
                            bold: true
                        };
                        worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = { horizontal: 'center' };
                    }
                    rowNameIndex = rowNameIndex + DataToExport[shapeIndex][0].data.length + 3
                }
                for (let i = 0; i < xlsQua.length; i++) {
                    let rowNameIndex = 6;
                    worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value = xlsQua[i].name;
                    worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
                        bold: true
                    };
                    worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = { horizontal: 'center' };
                    rowNameIndex = rowNameIndex + DataToExport[shapeIndex][0].data.length + 3
                }
                let colmnNameIndex = 7;
                for (let tableIndex = 0; tableIndex < DataToExport[shapeIndex].length; tableIndex++) {
                    for (let i = 0; i < DataToExport[shapeIndex][0].data.length; i++) {
                        worksheet.getCell('A' + (i + colmnNameIndex)).value = DataToExport[shapeIndex][0].data[i].C_NAME;
                        worksheet.getCell('A' + (i + colmnNameIndex)).font = {
                            bold: true
                        };
                        worksheet.getCell('A' + (i + colmnNameIndex)).alignment = { horizontal: 'center' };
                        for (var key in DataToExport[shapeIndex][0].data[i]) {
                            var value = DataToExport[shapeIndex][0].data[i][key];
                        }
                    }
                    colmnNameIndex = colmnNameIndex + DataToExport[shapeIndex][0].data.length + 3
                }
                let a = 0;
                let temp = 0;
                for (let tableIndex = 0; tableIndex < DataToExport[shapeIndex].length; tableIndex++) {
                    // a = a + 4
                    for (let j = 0; j < 26; j++) {
                        if (a == 0) {
                            a = a + 7;
                        } else {
                            a = a + 1
                        }
                        for (let i = 1; i < xlsQua.length + 1; i++) {

                            switch (i) {
                                case 1:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q1 ? DataToExport[shapeIndex][tableIndex].data[i].Q1 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 2:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q2 ? DataToExport[shapeIndex][tableIndex].data[0].Q2 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 3:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q3 ? DataToExport[shapeIndex][tableIndex].data[0].Q3 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 4:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q4 ? DataToExport[shapeIndex][tableIndex].data[0].Q4 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 5:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q5 ? DataToExport[shapeIndex][tableIndex].data[0].Q5 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 6:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q6 ? DataToExport[shapeIndex][tableIndex].data[0].Q6 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 7:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q7 ? DataToExport[shapeIndex][tableIndex].data[0].Q7 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 8:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q8 ? DataToExport[shapeIndex][tableIndex].data[0].Q8 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 9:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q9 ? DataToExport[shapeIndex][tableIndex].data[0].Q9 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 10:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q10 ? DataToExport[shapeIndex][tableIndex].data[0].Q10 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 11:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q11 ? DataToExport[shapeIndex][tableIndex].data[0].Q11 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 12:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q12 ? DataToExport[shapeIndex][tableIndex].data[0].Q12 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 13:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q13 ? DataToExport[shapeIndex][tableIndex].data[0].Q13 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 14:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q14 ? DataToExport[shapeIndex][tableIndex].data[0].Q14 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 15:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q15 ? DataToExport[shapeIndex][tableIndex].data[0].Q15 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 16:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q16 ? DataToExport[shapeIndex][tableIndex].data[0].Q16 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 17:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[j].Q17 ? DataToExport[shapeIndex][tableIndex].data[0].Q17 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                            }
                            worksheet.getCell(CellIndex[2] + a).protection = { locked: false };
                            worksheet.protect();
                        }
                    }
                    // console.log(CellIndex[0] + a);
                    a = a + 3
                }

            } else {

                let SizeIndex = 5;

                for (let tableIndex = 0; tableIndex < DataToExport[shapeIndex].length; tableIndex++) {

                    worksheet.mergeCells('B' + SizeIndex + ':R' + SizeIndex);
                    worksheet.getCell('B' + SizeIndex).value = `${DataToExport[shapeIndex][tableIndex].SZ_NAME}-${DataToExport[shapeIndex][tableIndex].FMETER}-${DataToExport[shapeIndex][tableIndex].TMETER}`;
                    worksheet.getCell('B' + SizeIndex).alignment = { horizontal: 'center' };
                    worksheet.getCell('B' + SizeIndex).fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'A0A0A0' },
                        font: { bold: true }
                    };
                    worksheet.getCell('B' + SizeIndex).font = {
                        bold: true
                    };

                    SizeIndex = SizeIndex + DataToExport[shapeIndex][0].data.length + 3
                }

                let xlsQua = [
                    { code: 'D1', name: 'FL' },
                    { code: 'D2', name: 'IF' },
                    { code: 'D3', name: 'VVS1' },
                    { code: 'D4', name: 'VVS2' },
                    { code: 'D5', name: 'VS1' },
                    { code: 'D6', name: 'VS2' },
                    { code: 'D7', name: 'SI1' },
                    { code: 'D8', name: 'SI2' },
                    { code: 'D9', name: 'SI3' },
                    { code: 'D10', name: 'I1' },
                    { code: 'D11', name: 'I2' },
                    { code: 'D12', name: 'I3' },
                    { code: 'D13', name: 'I4' },
                    { code: 'D14', name: 'I5' },
                    { code: 'D15', name: 'I6' },
                    { code: 'D16', name: 'I7' },
                    { code: 'D17', name: 'I8' },
                ]
                let rowNameIndex = 6;
                for (let tableIndex = 0; tableIndex < DataToExport[shapeIndex].length; tableIndex++) {
                    for (let i = 0; i < xlsQua.length; i++) {

                        worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value = xlsQua[i].name;
                        worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
                            bold: true
                        };
                        worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = { horizontal: 'center' };
                    }
                    rowNameIndex = rowNameIndex + DataToExport[shapeIndex][0].data.length + 3
                }
                for (let i = 0; i < xlsQua.length; i++) {
                    let rowNameIndex = 6;
                    worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value = xlsQua[i].name;
                    worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = { horizontal: 'center' };
                    rowNameIndex = rowNameIndex + DataToExport[shapeIndex][0].data.length + 3
                }
                let colmnNameIndex = 7;
                for (let tableIndex = 0; tableIndex < DataToExport[shapeIndex].length; tableIndex++) {
                    for (let i = 0; i < DataToExport[shapeIndex][0].data.length; i++) {
                        worksheet.getCell('A' + (i + colmnNameIndex)).value = DataToExport[shapeIndex][0].data[i].C_NAME;
                        worksheet.getCell('A' + (i + colmnNameIndex)).font = {
                            bold: true
                        };
                        worksheet.getCell('A' + (i + colmnNameIndex)).alignment = { horizontal: 'center' };
                        for (var key in DataToExport[shapeIndex][0].data[i]) {
                            var value = DataToExport[shapeIndex][0].data[i][key];
                        }
                    }
                    colmnNameIndex = colmnNameIndex + DataToExport[shapeIndex][0].data.length + 3
                }

                let a = 0;
                let temp = 0;
                for (let tableIndex = 0; tableIndex < DataToExport[shapeIndex].length; tableIndex++) {
                    // a = a + 4
                    for (let dataRowIndex = 0; dataRowIndex < DataToExport[shapeIndex][tableIndex].data.length; dataRowIndex++) {
                        if (a == 0) {
                            a = a + 7;
                        } else {
                            a = a + 1
                        }
                        for (let i = 1; i < xlsQua.length + 1; i++) {
                            switch (i) {
                                case 1:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q1 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q1 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 2:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q2 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q2 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 3:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q3 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q3 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 4:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q4 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q4 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 5:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q5 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q5 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 6:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q6 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q6 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 7:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q7 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q7 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 8:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q8 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q8 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 9:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q9 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q9 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 10:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q10 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q10 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 11:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q11 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q11 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 12:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q12 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q12 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 13:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q13 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q13 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 14:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q14 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q14 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 15:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q15 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q15 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 16:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q16 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q16 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 17:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q17 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q17 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                                case 18:
                                    worksheet.getCell(CellIndex[i] + a).value = DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q18 ? DataToExport[shapeIndex][tableIndex].data[dataRowIndex].Q18 : 0;
                                    worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                    worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                    break;
                            }
                            worksheet.getCell(CellIndex[2] + a).protection = { locked: false };
                            worksheet.protect();
                        }
                    }

                    a = a + 3
                }
            }
        }

    } else {

        let worksheet = workbook.addWorksheet(req.body.SheetName);

        //Today date and time
        var date_ob = new Date();
        var day = ("0" + date_ob.getDate()).slice(-2);
        var month = ("0" + (date_ob.getMonth() + 1)).slice(-2);
        var year = date_ob.getFullYear();
        var hours = date_ob.getHours();
        var minutes = date_ob.getMinutes();
        var dateTime = year + "-" + month + "-" + day + " " + hours + ":" + minutes;
        var shape = req.body.S_CODE + "_" + req.body.SheetName + "_" + req.body.FL_NAME

        //set header
        if (RTYPE == 'H') {
            worksheet.mergeCells('B1:R1');
        } else {
            worksheet.mergeCells('B1:R1');
        }
        worksheet.getCell('B1').value = dateTime;
        worksheet.getCell('B1').alignment = { horizontal: 'center', bgColor: { argb: 'FF00FF00' } };
        worksheet.getCell('B1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'A0A0A0' },
            font: { bold: true }
        };
        worksheet.getCell('B1').font = {
            bold: true
        };
        worksheet.getCell('B1').border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        if (RTYPE == 'H') {
            worksheet.mergeCells('B2:R2');
        } else {
            worksheet.mergeCells('B2:R2');
        }

        worksheet.getCell('B2').value = `${RObj.find((r) => r.code == req.body.RTYPE ? r.name : '').name}_${req.body.RAPTYPE}`;
        worksheet.getCell('B2').alignment = { horizontal: 'center', bgColor: { argb: 'FF00FF00' } };
        worksheet.getCell('B2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'A0A0A0' },
            font: { bold: true }
        };
        worksheet.getCell('B2').font = {
            bold: true
        };
        worksheet.getCell('B2').border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        if (RTYPE == 'H') {
            worksheet.mergeCells('B3:R3');
        } else {
            worksheet.mergeCells('B3:R3');
        }

        worksheet.getCell('B3').value = shape;
        worksheet.getCell('B3').alignment = { horizontal: 'center', bgColor: { argb: 'FF00FF00' } };
        worksheet.getCell('B3').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'A0A0A0' },
            font: { bold: true }
        };
        worksheet.getCell('B3').font = {
            bold: true
        };
        worksheet.getCell('B3').border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        if (RTYPE == 'H') {
            let SizeIndex = 5;
            for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {

                worksheet.mergeCells('B' + SizeIndex + ':R' + SizeIndex);
                worksheet.getCell('B' + SizeIndex).value = `${DataToExport[tableIndex].SZ_NAME}-${DataToExport[tableIndex].FMETER}-${DataToExport[tableIndex].TMETER}`;
                worksheet.getCell('B' + SizeIndex).alignment = { horizontal: 'center' };
                worksheet.getCell('B' + SizeIndex).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'A0A0A0' },
                    font: { bold: true }
                };
                worksheet.getCell('B' + SizeIndex).font = {
                    bold: true
                };
                SizeIndex = SizeIndex + DataToExport[0].data.length + 3
            }
            let xlsQua = Object.keys(DataToExport[0].data[0]).filter((item) => item.startsWith('Q'))
            xlsQua = QuaData.filter((item) => xlsQua.some((d) => item.code == d))

            let rowNameIndex = 6;
            for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
                for (let i = 0; i < xlsQua.length; i++) {

                    worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value = xlsQua[i].name;
                    worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
                        bold: true
                    };
                    worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = { horizontal: 'center' };
                }
                rowNameIndex = rowNameIndex + DataToExport[0].data.length + 3
            }
            for (let i = 0; i < xlsQua.length; i++) {
                let rowNameIndex = 6;
                worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value = xlsQua[i].name;
                worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
                    bold: true
                };
                worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = { horizontal: 'center' };
                rowNameIndex = rowNameIndex + DataToExport[0].data.length + 3
            }
            let colmnNameIndex = 7;
            for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
                for (let i = 0; i < DataToExport[0].data.length; i++) {
                    worksheet.getCell('A' + (i + colmnNameIndex)).value = DataToExport[0].data[i].C_NAME;
                    worksheet.getCell('A' + (i + colmnNameIndex)).font = {
                        bold: true
                    };
                    worksheet.getCell('A' + (i + colmnNameIndex)).alignment = { horizontal: 'center' };
                    for (var key in DataToExport[0].data[i]) {
                        var value = DataToExport[0].data[i][key];
                    }
                }
                colmnNameIndex = colmnNameIndex + DataToExport[0].data.length + 3
            }

            let a = 0;
            let temp = 0;
            for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
                // a = a + 4
                for (let dataRowIndex = 0; dataRowIndex < DataToExport[tableIndex].data.length; dataRowIndex++) {
                    if (a == 0) {
                        a = a + 7;
                    } else {
                        a = a + 1
                    }

                    for (let i = 1; i < xlsQua.length + 1; i++) {
                        switch (i) {
                            case 1:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q1 ? DataToExport[tableIndex].data[dataRowIndex].Q1 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 2:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q2 ? DataToExport[tableIndex].data[dataRowIndex].Q2 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 3:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q3 ? DataToExport[tableIndex].data[dataRowIndex].Q3 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 4:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q4 ? DataToExport[tableIndex].data[dataRowIndex].Q4 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 5:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q5 ? DataToExport[tableIndex].data[dataRowIndex].Q5 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 6:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q6 ? DataToExport[tableIndex].data[dataRowIndex].Q6 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 7:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q7 ? DataToExport[tableIndex].data[dataRowIndex].Q7 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 8:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q8 ? DataToExport[tableIndex].data[dataRowIndex].Q8 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 9:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q9 ? DataToExport[tableIndex].data[dataRowIndex].Q9 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 10:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q10 ? DataToExport[tableIndex].data[dataRowIndex].Q1 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 11:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q11 ? DataToExport[tableIndex].data[dataRowIndex].Q11 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 12:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q12 ? DataToExport[tableIndex].data[dataRowIndex].Q12 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 13:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q13 ? DataToExport[tableIndex].data[dataRowIndex].Q13 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 14:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q14 ? DataToExport[tableIndex].data[dataRowIndex].Q14 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 15:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q15 ? DataToExport[tableIndex].data[dataRowIndex].Q15 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 16:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q16 ? DataToExport[tableIndex].data[dataRowIndex].Q16 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 17:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q17 ? DataToExport[tableIndex].data[dataRowIndex].Q17 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                        }
                        worksheet.getCell(CellIndex[2] + a).protection = { locked: false };
                        worksheet.protect();
                    }
                }
                a = a + 3
            }
        } else {

            let SizeIndex = 5;
            for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
                worksheet.mergeCells('B' + SizeIndex + ':R' + SizeIndex);
                worksheet.getCell('B' + SizeIndex).value = `${DataToExport[tableIndex].SZ_NAME}-${DataToExport[tableIndex].FMETER}-${DataToExport[tableIndex].TMETER}`;
                worksheet.getCell('B' + SizeIndex).alignment = { horizontal: 'center' };
                worksheet.getCell('B' + SizeIndex).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'A0A0A0' },
                    font: { bold: true }
                };
                worksheet.getCell('B' + SizeIndex).font = {
                    bold: true
                };
                SizeIndex = SizeIndex + DataToExport[0].data.length + 3
            }

            let xlsQua = [
                { code: 'D1', name: 'FL' },
                { code: 'D2', name: 'IF' },
                { code: 'D3', name: 'VVS1' },
                { code: 'D4', name: 'VVS2' },
                { code: 'D5', name: 'VS1' },
                { code: 'D6', name: 'VS2' },
                { code: 'D7', name: 'SI1' },
                { code: 'D8', name: 'SI2' },
                { code: 'D9', name: 'SI3' },
                { code: 'D10', name: 'I1' },
                { code: 'D11', name: 'I2' },
                { code: 'D12', name: 'I3' },
                { code: 'D13', name: 'I4' },
                { code: 'D14', name: 'I5' },
                { code: 'D15', name: 'I6' },
                { code: 'D16', name: 'I7' },
                { code: 'D17', name: 'I8' }
            ]
            let rowNameIndex = 6;
            for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
                for (let i = 0; i < xlsQua.length; i++) {
                    worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value = xlsQua[i].name;
                    worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
                        bold: true
                    };
                }
                rowNameIndex = rowNameIndex + DataToExport[0].data.length + 3

            }
            for (let i = 0; i < xlsQua.length; i++) {
                let rowNameIndex = 6;
                worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value = xlsQua[i].name;
                worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
                    bold: true
                };
                worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = { horizontal: 'center' };
                rowNameIndex = rowNameIndex + DataToExport[0].data.length + 3

            }
            let colmnNameIndex = 7;
            for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
                for (let i = 0; i < DataToExport[0].data.length; i++) {
                    worksheet.getCell('A' + (i + colmnNameIndex)).value = DataToExport[0].data[i].C_NAME;
                    worksheet.getCell('A' + (i + colmnNameIndex)).font = {
                        bold: true
                    };
                    worksheet.getCell('A' + (i + colmnNameIndex)).alignment = { horizontal: 'center' };
                    for (var key in DataToExport[0].data[i]) {
                        var value = DataToExport[0].data[i][key];
                    }
                }
                colmnNameIndex = colmnNameIndex + DataToExport[0].data.length + 3

            }
            let a = 0;
            let temp = 0;
            for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
                for (let dataRowIndex = 0; dataRowIndex < DataToExport[tableIndex].data.length; dataRowIndex++) {
                    if (a == 0) {
                        a = a + 7;
                    } else {
                        a = a + 1;
                    }
                    for (let i = 0; i < xlsQua.length + 1; i++) {
                        switch (i) {
                            case 1:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q1 ? DataToExport[tableIndex].data[dataRowIndex].Q1 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 2:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q2 ? DataToExport[tableIndex].data[dataRowIndex].Q2 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 3:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q3 ? DataToExport[tableIndex].data[dataRowIndex].Q3 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 4:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q4 ? DataToExport[tableIndex].data[dataRowIndex].Q4 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 5:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q5 ? DataToExport[tableIndex].data[dataRowIndex].Q5 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 6:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q6 ? DataToExport[tableIndex].data[dataRowIndex].Q6 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 7:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q7 ? DataToExport[tableIndex].data[dataRowIndex].Q7 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 8:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q8 ? DataToExport[tableIndex].data[dataRowIndex].Q8 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 9:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q9 ? DataToExport[tableIndex].data[dataRowIndex].Q9 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 10:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q10 ? DataToExport[tableIndex].data[dataRowIndex].Q10 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 11:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q11 ? DataToExport[tableIndex].data[dataRowIndex].Q11 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 12:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q12 ? DataToExport[tableIndex].data[dataRowIndex].Q12 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 13:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q13 ? DataToExport[tableIndex].data[dataRowIndex].Q13 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 14:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q14 ? DataToExport[tableIndex].data[dataRowIndex].Q14 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 15:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q15 ? DataToExport[tableIndex].data[dataRowIndex].Q15 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 16:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q16 ? DataToExport[tableIndex].data[dataRowIndex].Q16 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                            case 17:
                                worksheet.getCell(CellIndex[i] + a).value = DataToExport[tableIndex].data[dataRowIndex].Q17 ? DataToExport[tableIndex].data[dataRowIndex].Q17 : 0;
                                worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
                                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                                break;
                        }
                        worksheet.getCell(CellIndex[2] + a).protection = { locked: false };
                        worksheet.protect();
                    }
                }
                a = a + 3
            }
        }
    }


    workbook.xlsx.writeBuffer().then(function (buffer) {
        let xlsData = Buffer.concat([buffer]);
        res.writeHead(200, {
            'Content-Length': Buffer.byteLength(xlsData),
            'Content-Type': 'application/vnd.ms-excel',
            'Content-disposition': 'attachment;filename=RapParamDisc.xlsx',
        }).end(xlsData);
    });
}
