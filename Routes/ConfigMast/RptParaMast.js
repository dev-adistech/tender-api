var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rptParaMast = require('../../Controllers/ConfigMast/RptParaMast.Controllers');

router.post('/RptParaMastFill', verifyToken, _rptParaMast.RptParaMastFill)
router.post('/RptParaMastSave', verifyToken, _rptParaMast.RptParaMastSave)
router.post('/RptParaMastDelete', verifyToken, _rptParaMast.RptParaMastDelete)

module.exports = router;