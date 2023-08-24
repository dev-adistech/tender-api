var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _TendarCom = require('../../Controllers/Transaction/TendarCom.Controller');

router.post('/TendarComMastFill', verifyToken, _TendarCom.TendarComMastFill)
router.post('/TendarComMastSave', verifyToken, _TendarCom.TendarComMastSave)
router.post('/TendarComMastDelete', verifyToken, _TendarCom.TendarComMastDelete)

module.exports = router;