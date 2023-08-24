var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rapaPortDia = require('../../Controllers/Rap/RapaportDia.Controllers');

router.post('/ORapParamDisFill', verifyToken, _rapaPortDia.ORapParamDisFill)
router.post('/ORapParamDisSave', verifyToken, _rapaPortDia.ORapParamDisSave)
router.post('/ORapParamDisDelete', verifyToken, _rapaPortDia.ORapParamDisDelete)

module.exports = router;
