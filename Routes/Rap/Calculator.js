var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _Calculator = require('../../Controllers/Rap/Calculator.Controllers');

router.post('/RapDisMastFill', verifyToken, _Calculator.RapDisMastFill)
router.post('/FindORap', verifyToken, _Calculator.FindORap)
// router.post('/ORapParamDisSave', verifyToken, _rapaPortDia.ORapParamDisSave)
// router.post('/ORapParamDisDelete', verifyToken, _rapaPortDia.ORapParamDisDelete)

module.exports = router;
