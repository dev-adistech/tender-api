var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _common = require('../../Controllers/Common/Common.Controllers');

router.post('/BarCode', verifyToken, _common.BarCode)
router.post('/BGImageUpload', verifyToken, _common.BGImageUpload)
router.post('/PassUpdate', verifyToken, _common.PassUpdate)

module.exports = router;
