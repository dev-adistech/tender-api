var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _LotMap = require('../../Controllers/Transaction/LotMap.Controllers');

router.post('/LotMappingFill', verifyToken, _LotMap.LotMappingFill)
router.post('/LotMappingSave', verifyToken, _LotMap.LotMappingSave)

module.exports = router;