var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _rateupdate = require('../../Controllers/Utility/rateupdate.Controllers');

router.post('/TendarRateUpd', verifyToken, _rateupdate.TendarRateUpd)

module.exports = router;