var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _MilkyMast = require('../../Controllers/Master/MilkyMast.Controller');


router.post('/MilkyMastFill', verifyToken, _MilkyMast.MilkyMastFill)
router.post('/MilkyMastSave', verifyToken, _MilkyMast.MilkyMastSave)
router.post('/MilkyMastDelete', verifyToken, _MilkyMast.MilkyMastDelete)

module.exports = router;