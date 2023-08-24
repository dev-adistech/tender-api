var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _TendarMast = require('../../Controllers/Transaction/TendarMast.Controller');

router.post('/TendarMastFill', verifyToken, _TendarMast.TendarMastFill)
router.post('/TendarMastDelete', verifyToken, _TendarMast.TendarMastDelete)
router.post('/TendarMastSave', verifyToken, _TendarMast.TendarMastSave)
router.post('/TendarPktEntFill', verifyToken, _TendarMast.TendarPktEntFill)
router.post('/TendarPktEntSave', verifyToken, _TendarMast.TendarPktEntSave)
router.post('/TendarPktEntDelete', verifyToken, _TendarMast.TendarPktEntDelete)
router.post('/GetTendarNumber', verifyToken, _TendarMast.GetTendarNumber)
router.post('/ChkTendarNumber', verifyToken, _TendarMast.ChkTendarNumber)
router.post('/GetTendarPktNumber', verifyToken, _TendarMast.GetTendarPktNumber)
router.post('/ChkTendarPktEnt', verifyToken, _TendarMast.ChkTendarPktEnt)

module.exports = router;