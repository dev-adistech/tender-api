var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _makreturn = require('../../Controllers/ConfigMast/Makreturn.Controllers');

router.post('/makreturnFill', verifyToken, _makreturn.makreturnFill)

module.exports = router;
