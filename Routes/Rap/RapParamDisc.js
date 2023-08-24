var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rapParamDisc = require('../../Controllers/Rap/RapParamDisc.Controllers');

router.post('/RapParamDiscFill', verifyToken, _rapParamDisc.RapParamDiscFill)
router.post('/RapParamDiscSave', verifyToken, _rapParamDisc.RapParamDiscSave)
router.post('/RapParamDiscImport', verifyToken, _rapParamDisc.RapParamDiscImport)
router.post('/RapParamDiscExport', verifyToken, _rapParamDisc.RapParamDiscExport)
router.post('/RapParamDiscExcelDownload', _rapParamDisc.RapParamDiscExcelDownload)

module.exports = router;
