var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rapOrg = require('../../Controllers/Rap/RapOrg.Controllers');

router.post('/RapOrgFill', verifyToken, _rapOrg.RapOrgFill)
router.post('/RapOrgSave', verifyToken, _rapOrg.RapOrgSave)
router.post('/RapDownload', verifyToken, _rapOrg.RapDownload)
router.post('/RapTrf', verifyToken, _rapOrg.RapTrf)

module.exports = router;
