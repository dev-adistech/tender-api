var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rptPerMast = require('../../Controllers/ConfigMast/RptPerMast.Controllers');

router.post('/RptPerMastFill', verifyToken, _rptPerMast.RptPerMastFill)
router.post('/RptPerMastSave', verifyToken, _rptPerMast.RptPerMastSave)
router.post('/RptPerMastDelete', verifyToken, _rptPerMast.RptPerMastDelete)
router.post('/RptPerMastCopy', verifyToken, _rptPerMast.RptPerMastCopy)

module.exports = router;