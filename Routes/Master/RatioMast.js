var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _RatioMast = require('../../Controllers/Master/RatioMast.Controller');

router.post('/RAtioMastFill', verifyToken, _RatioMast.RAtioMastFill)
router.post('/RatioMastSave', verifyToken, _RatioMast.RatioMastSave)
router.post('/RatioMastDelete', verifyToken, _RatioMast.RatioMastDelete)

module.exports = router;