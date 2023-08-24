var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rapCutDisc = require('../../Controllers/Rap/RapCutDisc.Controllers');

router.post('/RapCutDiscFill', verifyToken, _rapCutDisc.RapCutDiscFill)
router.post('/RapCutdSave', verifyToken, _rapCutDisc.RapCutdSave)
router.post('/RapCutInsertDis', verifyToken, _rapCutDisc.RapCutInsertDis)
router.post('/RapCutdImport', verifyToken, _rapCutDisc.RapCutdImport)
router.post('/RapCutdExport', _rapCutDisc.RapCutdExport)
router.post('/RapCutdExcelDownload', _rapCutDisc.RapCutdExcelDownload)

module.exports = router;
