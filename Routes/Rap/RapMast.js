var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rapMast = require('../../Controllers/Rap/RapMast.Controllers');

router.post('/RapTypeFill', verifyToken, _rapMast.RapTypeFill)
router.post('/RapMastFill', verifyToken, _rapMast.RapMastFill)
router.post('/RapMastRapOrgRate', verifyToken, _rapMast.RapMastRapOrgRate)
router.post('/RapMastSave', verifyToken, _rapMast.RapMastSave)
router.post('/RapMastExport', verifyToken, _rapMast.RapMastExport)
router.post('/RapMastImport', verifyToken, _rapMast.RapMastImport)
router.post('/RapDisComp', verifyToken, _rapMast.RapDisComp)
router.post('/RapMastImportUpdate', verifyToken, _rapMast.RapMastImportUpdate)
// router.post('/RouTrnSummaryExcelDownload', verifyToken, _rapMast.RouTrnSummaryExcelDownload)
router.post('/RapMastExcelDownload', _rapMast.RapMastExcelDownload)

module.exports = router;
