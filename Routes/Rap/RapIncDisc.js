var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rapIncDisc = require('../../Controllers/Rap/RapIncDisc.Controllers');

router.post('/RapIncDiscFill', verifyToken, _rapIncDisc.RapIncDiscFill)
router.post('/RapIncDiscSave', verifyToken, _rapIncDisc.RapIncDiscSave)
router.post('/RapIncDiscImport', verifyToken, _rapIncDisc.RapIncDiscImport)
router.post('/RapIncDiscExport', verifyToken, _rapIncDisc.RapIncDiscExport)
router.post('/RapIncDiscExcelDownload', _rapIncDisc.RapIncDiscExcelDownload)

module.exports = router;
