var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _ExpDate = require('../../Controllers/Transaction/ExpDate.Controllers');

router.post('/TendarExpDateFill', verifyToken, _ExpDate.TendarExpDateFill)
router.post('/TendarExpDateSave', verifyToken, _ExpDate.TendarExpDateSave)

module.exports = router;