var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _lotUtility = require('../../Controllers/ConfigMast/LotUtility.Controllers');

router.post('/LotChange', verifyToken, _lotUtility.LotChange)
router.post('/RateUpdateUtility', verifyToken, _lotUtility.RateUpdateUtility)

module.exports = router;
