var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _colMast = require('../../Controllers/Master/ColMast.Controllers');

router.post('/ColMastFill', verifyToken, _colMast.ColMastFill)
router.post('/ColMastSave', verifyToken, _colMast.ColMastSave)
router.post('/ColMastDelete', verifyToken, _colMast.ColMastDelete)

module.exports = router;