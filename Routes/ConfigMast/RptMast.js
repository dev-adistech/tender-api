var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rptMast = require('../../Controllers/ConfigMast/RptMast.Controllers');

router.post('/RptMastFill', verifyToken, _rptMast.RptMastFill)
router.post('/RptMastSave', verifyToken, _rptMast.RptMastSave)
router.post('/RptMastDelete', verifyToken, _rptMast.RptMastDelete)

module.exports = router;