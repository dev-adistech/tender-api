const sql = require("mssql");
const jwt = require("jsonwebtoken");
const Excel = require("exceljs");

var { _tokenSecret } = require("../../Config/token/TokenConfig.json");

var { RapServerConn } = require("../../config/db/rapsql");
const { RapMastRapOrgRate } = require("./RapMast.Controllers");

exports.RapFloDiscFill = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {

        let request = new sql.Request()

        request.input("RTYPE", sql.VarChar(2), req.body.RTYPE);
        request.input("SZ_CODE", sql.Int, parseInt(req.body.SZ_CODE));
        request.input("S_CODE", sql.VarChar(5), req.body.S_CODE);
        request.input("C_CODE", sql.Int, parseInt(req.body.C_CODE));
        request.input("FL_CODE", sql.Int, parseInt(req.body.FL_CODE));
        request.input("CT_CODE", sql.Int, parseInt(req.body.CT_CODE));

        request = await request.execute("USP_FrmRapFlodFill");

        if (request.recordset) {
          res.json({ success: 1, data: request.recordset });
        } else {
          res.json({ success: 1, data: "Not Found" });
        }
      } catch (err) {
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.RapFloDiscSave = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      try {
        const TokenData = await authData;
        let RapArr = req.body;
        let status = true;
        let ErrorRap = [];
        let request = new sql.Request()

        request.input("RTYPE", sql.VarChar(2), req.body.RTYPE);
        request.input("SZ_CODE", sql.Int, parseInt(req.body.SZ_CODE));
        request.input("S_CODE", sql.VarChar(5), req.body.S_CODE);
        request.input("C_CODE", sql.Int, parseInt(req.body.C_CODE));
        request.input("FL_CODE", sql.Int, parseInt(req.body.FL_CODE));
        request.input("CT_CODE", sql.Int, parseInt(req.body.CT_CODE));
        request.input("Q1", sql.Decimal(10, 2), parseFloat(req.body.Q1));
        request.input("Q2", sql.Decimal(10, 2), parseFloat(req.body.Q2));
        request.input("Q3", sql.Decimal(10, 2), parseFloat(req.body.Q3));
        request.input("Q4", sql.Decimal(10, 2), parseFloat(req.body.Q4));
        request.input("Q5", sql.Decimal(10, 2), parseFloat(req.body.Q5));
        request.input("Q6", sql.Decimal(10, 2), parseFloat(req.body.Q6));
        request.input("Q7", sql.Decimal(10, 2), parseFloat(req.body.Q7));
        request.input("Q8", sql.Decimal(10, 2), parseFloat(req.body.Q8));
        request.input("Q9", sql.Decimal(10, 2), parseFloat(req.body.Q9));
        request.input("Q10", sql.Decimal(10, 2), parseFloat(req.body.Q10));
        request.input("Q11", sql.Decimal(10, 2), parseFloat(req.body.Q11));
        request.input("Q12", sql.Decimal(10, 2), parseFloat(req.body.Q12));
        request.input("Q13", sql.Decimal(10, 2), parseFloat(req.body.Q13));
        request.input("Q14", sql.Decimal(10, 2), parseFloat(req.body.Q14));
        request.input("Q15", sql.Decimal(10, 2), parseFloat(req.body.Q15));
        request.input("Q16", sql.Decimal(10, 2), parseFloat(req.body.Q16));
        request.input("Q17", sql.Decimal(10, 2), parseFloat(req.body.Q17));
        request.input(
          "F_CARAT",
          sql.Decimal(10, 2),
          parseFloat(req.body.F_CARAT)
        );
        request.input(
          "T_CARAT",
          sql.Decimal(10, 2),
          parseFloat(req.body.T_CARAT)
        );

        request = await request.execute("USP_FrmRapFlodSave");

        if (request.recordset) {
          res.json({ success: 1, data: request.recordset });
        } else {
          res.json({ success: 1, data: "Not Found" });
        }
      } catch (err) {
        res.json({ success: 1, data: err });
      }
    }
  });
};

exports.RapFloDiscImport = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;
      let RapArr = req.body;
      // console.log(RapArr[0])
      let status = true;
      let ErrorRap = [];
      let count = 0;

      try {
        for (let shapeIndex = 0; shapeIndex < RapArr.length; shapeIndex++) {
          for (let i = 0; i < RapArr[shapeIndex].length; i++) {
            for (let j = 0; j < RapArr[shapeIndex][i].C_NAME.length; j++) {
              try {
                let request = new sql.Request()

                request.input('RTYPE', sql.VarChar(2), RapArr[shapeIndex][i].RTYPE)
                request.input('S_NAME', sql.VarChar(25), RapArr[shapeIndex][i].S_NAME)
                request.input('C_NAME', sql.VarChar(5), RapArr[shapeIndex][i].C_NAME[j])
                request.input('CT_NAME', sql.VarChar(10), RapArr[shapeIndex][i].CT_NAME)
                request.input('FL_NAME', sql.VarChar(10), RapArr[shapeIndex][i].FL_NAME)
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
                request.input('F_CARAT', sql.Numeric(10, 3), RapArr[shapeIndex][i].F_CARAT)
                request.input('T_CARAT', sql.Numeric(10, 3), RapArr[shapeIndex][i].T_CARAT)

                // console.log('##', shapeIndex, i, j);
                request = await request.execute('USP_FrmRapFlodImport');
                count++
              } catch (err) {
                console.log(err)
                status = false;
                ErrorRap.push(RapArr[i])
                // console.log(i, ': ####### Error: ', err);
              }
            }
          }
        }
        res.json({ success: status, data: ErrorRap });

      } catch (err) {
        console.log(err)
      }


    }
  });
};

exports.RapFloDiscExport = async (req, res) => {
  jwt.verify(req.token, _tokenSecret, async (err, authData) => {
    if (err) {
      res.sendStatus(401);
    } else {
      const TokenData = await authData;

      try {
        let request = new sql.Request()

        request.input("RTYPE", sql.VarChar(2), req.body.RTYPE);
        request.input("S_CODE", sql.VarChar(5), req.body.S_CODE);
        request.input("SZ_CODE", sql.VarChar(5), req.body.SZ_CODE);
        request.input("FL_CODE", sql.VarChar(5), req.body.FL_CODE);
        request.input("CT_CODE", sql.VarChar(5), req.body.CT_CODE);

        request = await request.execute("USP_FrmRapFlodExport");
        // console.log("ddd", request)
        res.json({
          success: 1,
          data: request.recordset,
          SZ_NAME: req.body.SZ_NAME,
        });
      } catch (err) {
        console.log(err);
        res.json({ success: 0, data: err });
      }
    }
  });
};

exports.DisFloExcelDownload = async (req, res) => {
  let DataToExport = [];
  let CellIndex = [
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
  ];

  DataToExport = JSON.parse(req.body.DataRow);
  QuaData = JSON.parse(req.body.QuaData);
  shape = req.body.SHAPE;
  RTYPE = req.body.RTYPE;
  // console.log("201", DataToExport.Data)
  let workbook = new Excel.Workbook();

  let RObj = [
    { code: "H", name: "HRD" },
    { code: "R", name: "LOOSE" },
    { code: "I", name: "IGI" },
    { code: "G", name: "GIA" },
  ];

  if (req.body.shapeCount) {
    shapeArr = req.body.shapeArr;
    shapeArr = shapeArr.split(",");
    for (let shapeIndex = 0; shapeIndex < DataToExport.length; shapeIndex++) {
      // for (let z = 0; z < req.body.shapeCount; z++) {

      let worksheet = workbook.addWorksheet(shapeArr[shapeIndex]);

      //TODAY DATE AND TIME
      var date_ob = new Date();
      var day = ("0" + date_ob.getDate()).slice(-2);
      var month = ("0" + (date_ob.getMonth() + 1)).slice(-2);
      var year = date_ob.getFullYear();
      var hours = date_ob.getHours();
      var minutes = date_ob.getMinutes();
      var dateTime =
        year + "-" + month + "-" + day + " " + hours + ":" + minutes;

      //set header
      worksheet.mergeCells("B1:R1");

      worksheet.getCell("B1").value = dateTime;
      worksheet.getCell("B1").alignment = {
        horizontal: "center",
        bgColor: { argb: "FF00FF00" },
      };
      worksheet.getCell("B1").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "A0A0A0" },
        font: { bold: true },
      };
      worksheet.getCell("B1").font = {
        bold: true,
      };
      worksheet.getCell("B1").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.mergeCells("B2:R2");
      worksheet.getCell("B2").value = `${RObj.find((r) => (r.code == req.body.RTYPE ? r.name : "")).name
        }`;
      worksheet.getCell("B2").alignment = {
        horizontal: "center",
        bgColor: { argb: "FF00FF00" },
      };
      worksheet.getCell("B2").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "A0A0A0" },
        font: { bold: true },
      };
      worksheet.getCell("B2").font = {
        bold: true,
      };
      worksheet.getCell("B2").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet.mergeCells("B3:R3");

      worksheet.getCell(
        "B3"
      ).value = `${req.body.S_CODE}-${req.body.CT_CODE}-${shapeArr[shapeIndex]}`;
      worksheet.getCell("B3").alignment = {
        horizontal: "center",
        bgColor: { argb: "FF00FF00" },
      };
      worksheet.getCell("B3").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "A0A0A0" },
        font: { bold: true },
      };
      worksheet.getCell("B3").font = {
        bold: true,
      };
      worksheet.getCell("B3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      // worksheet.mergeCells('B5:R5');
      // worksheet.getCell('B5').value =  DataToExport[0].SIZE;
      // worksheet.getCell('B5').alignment = { horizontal:'center'} ;
      // if (RTYPE == 'R') {

      let SizeIndex = 5;
      for (
        let tableIndex = 0;
        tableIndex < DataToExport[shapeIndex].length;
        tableIndex++
      ) {
        worksheet.mergeCells("B" + SizeIndex + ":R" + SizeIndex);
        worksheet.getCell("B" + SizeIndex).value = DataToExport[shapeIndex][
          tableIndex
        ].SIZE
          ? DataToExport[shapeIndex][tableIndex].SIZE
          : 0;
        worksheet.getCell("B" + SizeIndex).alignment = { horizontal: "center" };
        worksheet.getCell("B" + SizeIndex).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "A0A0A0" },
          font: { bold: true },
        };
        worksheet.getCell("B" + SizeIndex).font = {
          bold: true,
        };

        SizeIndex = SizeIndex + DataToExport[shapeIndex][0].Data.length + 3;
      }

      // let xlsQua = Object.keys(DataToExport[shapeIndex][0].Data[0]).filter((item) => item.startsWith('C_N'))
      // console.log('313', xlsQua)
      // console.log('313', QuaData)
      // xlsQua = QuaData.filter((item) => xlsQua.some((d) => item.code == d))
      let xlsQua = [
        { code: "Q1", name: "FL" },
        { code: "Q2", name: "IF" },
        { code: "Q3", name: "VVS1" },
        { code: "Q4", name: "VVS2" },
        { code: "Q5", name: "VS1" },
        { code: "Q6", name: "VS2" },
        { code: "Q7", name: "SI1" },
        { code: "Q8", name: "SI2" },
        { code: "Q9", name: "SI3" },
        { code: "Q10", name: "I1" },
        { code: "Q11", name: "I2" },
        { code: "Q12", name: "I3" },
        { code: "Q13", name: "I4" },
        { code: "Q14", name: "I5" },
        { code: "Q15", name: "I6" },
        { code: "Q16", name: "I7" },
        { code: "Q17", name: "I8" },
      ];

      // console.log("313", xlsQua)

      let rowNameIndex = 6;
      for (
        let tableIndex = 0;
        tableIndex < DataToExport[shapeIndex].length;
        tableIndex++
      ) {
        for (let i = 0; i < xlsQua.length; i++) {
          worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value =
            xlsQua[i].name;
          worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
            bold: true,
          };
          worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = {
            horizontal: "center",
          };
        }
        rowNameIndex =
          rowNameIndex + DataToExport[shapeIndex][0].Data.length + 3;
      }
      for (let i = 0; i < xlsQua.length; i++) {
        let rowNameIndex = 6;
        worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value =
          xlsQua[i].name;

        worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = {
          horizontal: "center",
        };
        rowNameIndex =
          rowNameIndex + DataToExport[shapeIndex][0].Data.length + 3;
      }
      let colmnNameIndex = 7;
      for (
        let tableIndex = 0;
        tableIndex < DataToExport[shapeIndex].length;
        tableIndex++
      ) {
        for (let i = 0; i < DataToExport[shapeIndex][0].Data.length; i++) {
          worksheet.getCell("A" + (i + colmnNameIndex)).value =
            DataToExport[shapeIndex][0].Data[i].C_NAME;
          worksheet.getCell("A" + (i + colmnNameIndex)).font = {
            bold: true,
          };
          worksheet.getCell("A" + (i + colmnNameIndex)).alignment = {
            horizontal: "center",
          };
          for (var key in DataToExport[shapeIndex][0].Data[i]) {
            var value = DataToExport[shapeIndex][0].Data[i][key];
          }
        }
        colmnNameIndex =
          colmnNameIndex + DataToExport[shapeIndex][0].Data.length + 3;
      }

      let a = 0;
      let temp = 0;
      for (
        let tableIndex = 0;
        tableIndex < DataToExport[shapeIndex].length;
        tableIndex++
      ) {
        // console.log("374", DataToExport[shapeIndex].length)
        // console.log("375", DataToExport[shapeIndex][tableIndex].Data)

        // a = a + 4
        for (let j = 0; j < 26; j++) {
          if (a == 0) {
            a = a + 7;
          } else {
            a = a + 1;
          }
          // console.log('383', xlsQua.length)
          for (let i = 1; i < xlsQua.length + 1; i++) {
            switch (i) {
              case 1:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q1 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q1;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 2:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q2 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q2;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 3:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q3 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q3;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 4:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q4 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q4;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 5:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q5 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q5;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 6:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q6 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q6;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 7:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q7 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q7;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 8:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q8 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q8;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 9:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q9 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q9;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 10:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q10 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q10;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 11:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q11 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q11;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 12:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q12 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q12;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 13:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q13 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q13;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 14:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q14 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q14;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 15:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q15 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q15;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 16:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q16 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q16;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 17:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[shapeIndex][tableIndex].Data[j].Q17 == null
                    ? 0
                    : DataToExport[shapeIndex][tableIndex].Data[j].Q17;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
            }
            worksheet.getCell(CellIndex[2] + a).protection = { locked: false };
            worksheet.protect();
          }
        }
        // console.log(CellIndex[0] + a);
        a = a + 3;
      }

      // } else {
      //     console.log('462: R Else')
      //     let SizeIndex = 5;

      //     for (let tableIndex = 0; tableIndex < DataToExport[shapeIndex].length; tableIndex++) {

      //         worksheet.mergeCells('B' + SizeIndex + ':R' + SizeIndex);
      //         worksheet.getCell('B' + SizeIndex).value = DataToExport[shapeIndex][tableIndex].SIZE;
      //         worksheet.getCell('B' + SizeIndex).alignment = { horizontal: 'center' };
      //         worksheet.getCell('B' + SizeIndex).fill = {
      //             type: 'pattern',
      //             pattern: 'solid',
      //             fgColor: { argb: 'A0A0A0' },
      //             font: { bold: true }
      //         };
      //         worksheet.getCell('B' + SizeIndex).font = {
      //             bold: true
      //         };

      //         SizeIndex = SizeIndex + DataToExport[shapeIndex][0].Data.length + 3
      //     }

      //     // let xlsQua = Object.keys(DataToExport[shapeIndex][0].Data[0]).filter((item) => item.startsWith('Q'))
      //     // xlsQua = QuaData.filter((item) => xlsQua.some((d) => item.code == d))
      //     // console.log(xlsQua);

      //     let xlsQua = [
      //         { code: 'D1', name: 'FL' },
      //         { code: 'D2', name: 'IF' },
      //         { code: 'D3', name: 'VVS1' },
      //         { code: 'D4', name: 'VVS2' },
      //         { code: 'D5', name: 'VS1' },
      //         { code: 'D6', name: 'VS2' },
      //         { code: 'D7', name: 'SI1' },
      //         { code: 'D8', name: 'SI2' },
      //         { code: 'D9', name: 'SI3' },
      //         { code: 'D10', name: 'I1' },
      //         { code: 'D11', name: 'I2' },
      //         { code: 'D12', name: 'I3' },
      //         { code: 'D13', name: 'I4' },
      //         { code: 'D14', name: 'I5' },
      //         { code: 'D15', name: 'I6' },
      //         { code: 'D16', name: 'I7' },
      //         { code: 'D17', name: 'I8' },
      //     ]
      //     let rowNameIndex = 6;
      //     for (let tableIndex = 0; tableIndex < DataToExport[shapeIndex].length; tableIndex++) {
      //         for (let i = 0; i < xlsQua.length; i++) {

      //             worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value = xlsQua[i].name;
      //             worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
      //                 bold: true
      //             };
      //             worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = { horizontal: 'center' };
      //         }
      //         rowNameIndex = rowNameIndex + DataToExport[shapeIndex][0].Data.length + 3
      //     }
      //     for (let i = 0; i < xlsQua.length; i++) {
      //         let rowNameIndex = 6;
      //         worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value = xlsQua[i].name;
      //         worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = { horizontal: 'center' };
      //         rowNameIndex = rowNameIndex + DataToExport[shapeIndex][0].Data.length + 3
      //     }
      //     let colmnNameIndex = 7;
      //     for (let tableIndex = 0; tableIndex < DataToExport[shapeIndex].length; tableIndex++) {
      //         for (let i = 0; i < DataToExport[shapeIndex][0].Data.length; i++) {
      //             worksheet.getCell('A' + (i + colmnNameIndex)).value = DataToExport[shapeIndex][0].Data[i].C_NAME;
      //             worksheet.getCell('A' + (i + colmnNameIndex)).font = {
      //                 bold: true
      //             };
      //             worksheet.getCell('A' + (i + colmnNameIndex)).alignment = { horizontal: 'center' };
      //             for (var key in DataToExport[shapeIndex][0].Data[i]) {
      //                 var value = DataToExport[shapeIndex][0].Data[i][key];
      //             }
      //         }
      //         colmnNameIndex = colmnNameIndex + DataToExport[shapeIndex][0].Data.length + 3
      //     }

      //     let a = 0;
      //     let temp = 0;
      //     for (let tableIndex = 0; tableIndex < DataToExport[shapeIndex].length; tableIndex++) {
      //         // a = a + 4
      //         for (let dataRowIndex = 0; dataRowIndex < DataToExport[shapeIndex][tableIndex].Data.length; dataRowIndex++) {
      //             if (a == 0) {
      //                 a = a + 7;
      //             } else {
      //                 a = a + 1
      //             }
      //             for (let i = 1; i < xlsQua.length + 1; i++) {
      //                 // switch (i) {
      //                 //     case 1:
      //                 //         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D1 ? 0 : DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D1;
      //                 //         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                 //         break;
      //                 //     case 2:
      //                 //         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D2 ? 0 : DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D2;
      //                 //         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                 //         break;
      //                 //     case 3:
      //                 //         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D3 ? 0 : DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D3;
      //                 //         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                 //         break;
      //                 //     case 4:
      //                 //         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D4 ? 0 : DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D4;
      //                 //         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                 //         break;
      //                 //     case 5:
      //                 //         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D5 ? 0 : DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D5;
      //                 //         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                 //         break;
      //                 //     case 6:
      //                 //         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D6 ? 0 : DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D6;
      //                 //         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                 //         break;
      //                 //     case 7:
      //                 //         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D7 ? 0 : DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D7;
      //                 //         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                 //         break;
      //                 //     case 8:
      //                 //         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D8 ? 0 : DataToExport[shapeIndex][tableIndex].Data[dataRowIndex].D8;
      //                 //         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                 //         break;

      //                 // }
      //                 switch (i) {
      //                     case 1:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q1 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q1;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 2:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q2 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q2;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 3:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q3 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q3;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 4:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q4 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q4;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 5:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q5 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q5;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 6:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q6 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q6;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 7:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q7 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q7;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 8:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q8 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q8;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 9:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q9 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q9;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 10:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q10 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q10;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 11:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q11 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q11;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 12:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q12 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q12;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 13:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q13 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q13;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 14:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q14 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q14;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 15:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q15 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q15;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 16:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q16 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q16;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 17:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[shapeIndex][tableIndex].Data[j].Q17 ? 0 : DataToExport[shapeIndex][tableIndex].Data[j].Q17;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;

      //                 }

      //             }
      //         }
      //         // console.log(CellIndex[0] + a);
      //         a = a + 3
      //     }
      // }
    }
    // }
  } else {
    try {
      let worksheet = workbook.addWorksheet(req.body.SHAPE);

      //Today date and time
      var date_ob = new Date();
      var day = ("0" + date_ob.getDate()).slice(-2);
      var month = ("0" + (date_ob.getMonth() + 1)).slice(-2);
      var year = date_ob.getFullYear();
      var hours = date_ob.getHours();
      var minutes = date_ob.getMinutes();
      var dateTime =
        year + "-" + month + "-" + day + " " + hours + ":" + minutes;

      //set header
      worksheet.mergeCells("B1:R1");
      worksheet.getCell("B1").value = dateTime;
      worksheet.getCell("B1").alignment = {
        horizontal: "center",
        bgColor: { argb: "FF00FF00" },
      };
      worksheet.getCell("B1").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "A0A0A0" },
        font: { bold: true },
      };
      worksheet.getCell("B1").font = {
        bold: true,
      };
      worksheet.getCell("B1").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet.mergeCells("B2:R2");

      worksheet.getCell("B2").value = `${RObj.find((r) => (r.code == req.body.RTYPE ? r.name : "")).name
        }`;
      worksheet.getCell("B2").alignment = {
        horizontal: "center",
        bgColor: { argb: "FF00FF00" },
      };
      worksheet.getCell("B2").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "A0A0A0" },
        font: { bold: true },
      };
      worksheet.getCell("B2").font = {
        bold: true,
      };
      worksheet.getCell("B2").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.mergeCells("B3:R3");

      worksheet.getCell(
        "B3"
      ).value = `${req.body.S_CODE}-${req.body.CT_CODE}-${req.body.SHAPE}`;
      worksheet.getCell("B3").alignment = {
        horizontal: "center",
        bgColor: { argb: "FF00FF00" },
      };
      worksheet.getCell("B3").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "A0A0A0" },
        font: { bold: true },
      };
      worksheet.getCell("B3").font = {
        bold: true,
      };
      worksheet.getCell("B3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      // worksheet.mergeCells('B5:R5');
      // worksheet.getCell('B5').value =  DataToExport[0].SIZE;
      // worksheet.getCell('B5').alignment = { horizontal:'center'} ;
      // if (RTYPE == 'R') {
      let SizeIndex = 5;
      for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
        // console.log(DataToExport[tableIndex])
        worksheet.mergeCells("B" + SizeIndex + ":R" + SizeIndex);
        worksheet.getCell("B" + SizeIndex).value = DataToExport[tableIndex].SIZE
          ? DataToExport[tableIndex].SIZE
          : 0;
        worksheet.getCell("B" + SizeIndex).alignment = { horizontal: "center" };
        worksheet.getCell("B" + SizeIndex).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "A0A0A0" },
          font: { bold: true },
        };
        worksheet.getCell("B" + SizeIndex).font = {
          bold: true,
        };

        SizeIndex = SizeIndex + DataToExport[0].Data.length + 3;
      }

      // let xlsQua = Object.keys(DataToExport[0].Data[0]).filter((item) => item.startsWith('Q'))
      // xlsQua = QuaData.filter((item) => xlsQua.some((d) => item.code == d))

      let xlsQua = [
        { code: "D1", name: "FL" },
        { code: "D2", name: "IF" },
        { code: "D3", name: "VVS1" },
        { code: "D4", name: "VVS2" },
        { code: "D5", name: "VS1" },
        { code: "D6", name: "VS2" },
        { code: "D7", name: "SI1" },
        { code: "D8", name: "SI2" },
        { code: "D9", name: "SI3" },
        { code: "D10", name: "I1" },
        { code: "D11", name: "I2" },
        { code: "D12", name: "I3" },
        { code: "D13", name: "I4" },
        { code: "D14", name: "I5" },
        { code: "D15", name: "I6" },
        { code: "D16", name: "I7" },
        { code: "D17", name: "I8" },
      ];

      let rowNameIndex = 6;
      for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
        for (let i = 0; i < xlsQua.length; i++) {
          worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value =
            xlsQua[i].name;
          worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
            bold: true,
          };
          worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = {
            horizontal: "center",
          };
        }
        rowNameIndex = rowNameIndex + DataToExport[0].Data.length + 3;
      }
      for (let i = 0; i < xlsQua.length; i++) {
        let rowNameIndex = 6;
        worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value =
          xlsQua[i].name;
        // worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
        //     bold: true
        // };
        worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = {
          horizontal: "center",
        };
        rowNameIndex = rowNameIndex + DataToExport[0].Data.length + 3;
      }
      let colmnNameIndex = 7;
      for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
        for (let i = 0; i < DataToExport[0].Data.length; i++) {
          // console.log("817", DataToExport[0].Data.length);
          worksheet.getCell("A" + (i + colmnNameIndex)).value =
            DataToExport[0].Data[i].C_NAME == null
              ? DataToExport[0].Data[i].C_NAME
              : DataToExport[0].Data[i].C_NAME;
          worksheet.getCell("A" + (i + colmnNameIndex)).font = {
            bold: true,
          };
          worksheet.getCell("A" + (i + colmnNameIndex)).alignment = {
            horizontal: "center",
          };
          for (var key in DataToExport[0].Data[i]) {
            var value = DataToExport[0].Data[i][key];
          }
        }
        colmnNameIndex = colmnNameIndex + DataToExport[0].Data.length + 3;
      }

      let a = 0;
      let temp = 0;
      for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
        // a = a + 4
        for (
          let dataRowIndex = 0;
          dataRowIndex < DataToExport[tableIndex].Data.length;
          dataRowIndex++
        ) {

          if (a == 0) {
            a = a + 7;
          } else {
            a = a + 1;
          }

          for (let i = 1; i < xlsQua.length + 1; i++) {
            switch (i) {
              case 1:
                worksheet.getCell(CellIndex[i] + a).value = !DataToExport[
                  tableIndex
                ].Data[dataRowIndex].Q1
                  ? 0
                  : DataToExport[tableIndex].Data[dataRowIndex].Q1;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 2:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q2 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q2;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 3:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q3 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q3;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 4:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q4 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q4;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 5:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q5 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q5;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 6:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q6 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q6;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 7:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q7 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q7;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 8:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q8 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q8;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 9:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q9 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q9;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 10:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q10 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q10;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 11:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q11 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q11;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 12:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q12 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q12;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 13:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q13 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q13;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 14:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q14 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q14;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 15:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q15 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q15;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 16:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q16 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q16;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
              case 17:
                worksheet.getCell(CellIndex[i] + a).value =
                  DataToExport[tableIndex].Data[dataRowIndex].Q17 == null
                    ? 0
                    : DataToExport[tableIndex].Data[dataRowIndex].Q17;
                worksheet.getCell(CellIndex[i] + a).alignment = {
                  horizontal: "center",
                };
                worksheet.getCell(CellIndex[i] + a).protection = { locked: false };
                break;
            }
            worksheet.getCell(CellIndex[2] + a).protection = { locked: false };
            worksheet.protect();
          }
        }
        // console.log(CellIndex[0] + a);
        a = a + 3;
      }
      // } else {
      //     let SizeIndex = 5;
      //     for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {

      //         worksheet.mergeCells('B' + SizeIndex + ':R' + SizeIndex);
      //         worksheet.getCell('B' + SizeIndex).value = DataToExport[tableIndex].SIZE;
      //         worksheet.getCell('B' + SizeIndex).alignment = { horizontal: 'center' };
      //         worksheet.getCell('B' + SizeIndex).fill = {
      //             type: 'pattern',
      //             pattern: 'solid',
      //             fgColor: { argb: 'A0A0A0' },
      //             font: { bold: true }
      //         };
      //         worksheet.getCell('B' + SizeIndex).font = {
      //             bold: true
      //         };

      //         SizeIndex = SizeIndex + DataToExport[0].Data.length + 3
      //     }

      //     // let xlsQua = Object.keys(DataToExport[0].Data[0]).filter((item) => item.startsWith('Q'))
      //     // xlsQua = QuaData.filter((item) => xlsQua.some((d) => item.code == d))

      //     let xlsQua = [
      //         { code: 'D1', name: 'FL' },
      //         { code: 'D2', name: 'IF' },
      //         { code: 'D3', name: 'VVS1' },
      //         { code: 'D4', name: 'VVS2' },
      //         { code: 'D5', name: 'VS1' },
      //         { code: 'D6', name: 'VS2' },
      //         { code: 'D7', name: 'SI1' },
      //         { code: 'D8', name: 'SI2' },
      //     ]

      //     let rowNameIndex = 6;
      //     for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
      //         for (let i = 0; i < xlsQua.length; i++) {

      //             worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value = xlsQua[i].name;
      //             worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
      //                 bold: true
      //             };
      //         }
      //         rowNameIndex = rowNameIndex + DataToExport[0].Data.length + 3
      //     }
      //     for (let i = 0; i < xlsQua.length; i++) {
      //         let rowNameIndex = 6;
      //         worksheet.getCell(CellIndex[i + 1] + rowNameIndex).value = xlsQua[i].name;
      //         worksheet.getCell(CellIndex[i + 1] + rowNameIndex).font = {
      //             bold: true
      //         };
      //         worksheet.getCell(CellIndex[i + 1] + rowNameIndex).alignment = { horizontal: 'center' };
      //         rowNameIndex = rowNameIndex + DataToExport[0].Data.length + 3
      //     }
      //     let colmnNameIndex = 7;
      //     for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
      //         for (let i = 0; i < DataToExport[0].Data.length; i++) {
      //             worksheet.getCell('A' + (i + colmnNameIndex)).value = DataToExport[0].Data[i].C_NAME;
      //             worksheet.getCell('A' + (i + colmnNameIndex)).font = {
      //                 bold: true
      //             };
      //             worksheet.getCell('A' + (i + colmnNameIndex)).alignment = { horizontal: 'center' };

      //             for (var key in DataToExport[0].Data[i]) {
      //                 var value = DataToExport[0].Data[i][key];
      //             }
      //         }
      //         colmnNameIndex = colmnNameIndex + DataToExport[0].Data.length + 3
      //     }
      //     let a = 0;
      //     let temp = 0;
      //     for (let tableIndex = 0; tableIndex < DataToExport.length; tableIndex++) {
      //         // a = a + 4
      //         for (let dataRowIndex = 0; dataRowIndex < DataToExport[tableIndex].Data.length; dataRowIndex++) {
      //             if (a == 0) {
      //                 a = a + 7;
      //             } else {

      //                 a = a + 1
      //             }

      //             for (let i = 1; i < xlsQua.length + 1; i++) {
      //                 switch (i) {
      //                     case 1:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[tableIndex].Data[dataRowIndex].Q1 ? 0 : DataToExport[tableIndex].Data[dataRowIndex].Q1;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 2:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[tableIndex].Data[dataRowIndex].Q2 ? 0 : DataToExport[tableIndex].Data[dataRowIndex].Q2;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 3:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[tableIndex].Data[dataRowIndex].Q3 ? 0 : DataToExport[tableIndex].Data[dataRowIndex].Q3;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 4:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[tableIndex].Data[dataRowIndex].Q4 ? 0 : DataToExport[tableIndex].Data[dataRowIndex].Q4;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 5:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[tableIndex].Data[dataRowIndex].Q5 ? 0 : DataToExport[tableIndex].Data[dataRowIndex].Q5;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 6:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[tableIndex].Data[dataRowIndex].Q6 ? 0 : DataToExport[tableIndex].Data[dataRowIndex].Q6;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 7:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[tableIndex].Data[dataRowIndex].Q7 ? 0 : DataToExport[tableIndex].Data[dataRowIndex].Q7;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;
      //                     case 8:
      //                         worksheet.getCell(CellIndex[i] + a).value = !DataToExport[tableIndex].Data[dataRowIndex].Q8 ? 0 : DataToExport[tableIndex].Data[dataRowIndex].Q8;
      //                         worksheet.getCell(CellIndex[i] + a).alignment = { horizontal: 'center' };
      //                         break;

      //                 }
      //             }
      //         }
      //         // console.log(CellIndex[0] + a);
      //         a = a + 3
      //     }
      // }
    } catch (err) {
      console.log(err);
    }
  }
  workbook.xlsx.writeBuffer().then(function (buffer) {
    let xlsData = Buffer.concat([buffer]);
    res
      .writeHead(200, {
        "Content-Length": Buffer.byteLength(xlsData),
        "Content-Type": "application/vnd.ms-excel",
        "Content-disposition": "attachment;filename=RapFloDisc.xlsx",
      })
      .end(xlsData);
  });
};
