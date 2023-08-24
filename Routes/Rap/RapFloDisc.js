var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rapFloDisc = require('../../Controllers/Rap/RapFloDisc.Controllers');

router.post('/RapFloDiscFill', verifyToken, _rapFloDisc.RapFloDiscFill)
router.post('/RapFloDiscSave', verifyToken, _rapFloDisc.RapFloDiscSave)
router.post('/RapFloDiscImport', verifyToken, _rapFloDisc.RapFloDiscImport)
router.post('/RapFloDiscExport', verifyToken, _rapFloDisc.RapFloDiscExport)
router.post('/RapFloDiscExcelDownload', _rapFloDisc.DisFloExcelDownload)

module.exports = router;