var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rapShdDisc = require('../../Controllers/Rap/RapShdDisc.Controllers');

router.post('/RapShdDiscFill', verifyToken, _rapShdDisc.RapShdDiscFill)
router.post('/RapShdDiscSave', verifyToken, _rapShdDisc.RapShdDiscSave)
router.post('/RapShdDiscImport', verifyToken, _rapShdDisc.RapShdDiscImport)
router.post('/RapShdDiscExport', verifyToken, _rapShdDisc.RapShdDiscExport)
router.post('/RapOrdIsFill', verifyToken, _rapShdDisc.RapOrdIsFill)
router.post('/RapOrdIsDelete', verifyToken, _rapShdDisc.RapOrdIsDelete)
router.post('/RapOrdIsSave', verifyToken, _rapShdDisc.RapOrdIsSave)
router.post('/RapShdDiscExcelDownload', _rapShdDisc.RapShdDiscExcelDownload)

module.exports = router;
