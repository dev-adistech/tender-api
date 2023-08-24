var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _cutMast = require('../../Controllers/Master/CutMast.Controllers');

router.post('/CutMastFill', verifyToken, _cutMast.CutMastFill)
router.post('/CutMastSave', verifyToken, _cutMast.CutMastSave)
router.post('/CutMastDelete', verifyToken, _cutMast.CutMastDelete)

module.exports = router;