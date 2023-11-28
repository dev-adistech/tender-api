var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rapCalc = require('../../Controllers/Rap/RapCalc.Controllers');

router.post('/RapCalcDisp', verifyToken, _rapCalc.RapCalcDisp)
router.post('/FindRap', verifyToken, _rapCalc.FindRap)
router.post('/FindRapType', verifyToken, _rapCalc.FindRapType)
router.post('/RapSave', verifyToken, _rapCalc.RapSave)
router.post('/PrdTagUpd', verifyToken, _rapCalc.PrdTagUpd)
router.post('/PrdEmpFill', verifyToken, _rapCalc.PrdEmpFill)
router.post('/RapCalcCrtFill', verifyToken, _rapCalc.RapCalcCrtFill)
router.post('/RapCalcSaveValidation', verifyToken, _rapCalc.RapCalcSaveValidation)
router.post('/RapCalcSelectValidation', verifyToken, _rapCalc.RapCalcSelectValidation)
router.post('/RapPrint', verifyToken, _rapCalc.RapPrint)
router.post('/RapPrintOP', verifyToken, _rapCalc.RapPrintOP)
router.post('/RapCalPrdDel', verifyToken, _rapCalc.RapCalPrdDel)
router.post('/RapCalcTendarVidUpload', verifyToken, _rapCalc.RapCalcTendarVidUpload)
router.post('/RapCalcTendarVidDisp', verifyToken, _rapCalc.RapCalcTendarVidDisp)
router.post('/RapCalcTendarVidDelete', verifyToken, _rapCalc.RapCalcTendarVidDelete)
router.post('/TendarRapCalDisp', verifyToken, _rapCalc.TendarRapCalDisp)
router.post('/TendarRapCalPrdSave', verifyToken, _rapCalc.TendarRapCalPrdSave)
router.post('/TendarRapCalPrdDelete', verifyToken, _rapCalc.TendarRapCalPrdDelete)
router.post('/TendarExcel', verifyToken, _rapCalc.TendarExcel)
router.post('/TendarExcelDownload', _rapCalc.TendarExcelDownload)
router.post('/TendarExcelDownloadSKD', _rapCalc.TendarExcelDownloadSKD)
router.post('/RapCalExport', _rapCalc.RapCalExport)
router.post('/RapCalcPrdQuery', verifyToken, _rapCalc.RapCalcPrdQuery)
router.post('/RapCalcTendarDel', verifyToken, _rapCalc.RapCalcTendarDel)
router.post('/TenderLotDelete', verifyToken, _rapCalc.TendarLotDelete)
router.post('/RapCalcDoubleClick', verifyToken, _rapCalc.RapCalcDoubleClick)

module.exports = router;
