var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _shpMast = require('../../Controllers/Master/ShpMast.Controllers');

router.post('/ShpMastFill', verifyToken, _shpMast.ShpMastFill)
router.post('/ShpMastSave', verifyToken, _shpMast.ShpMastSave)
router.post('/ShpMastDelete', verifyToken, _shpMast.ShpMastDelete)

module.exports = router;