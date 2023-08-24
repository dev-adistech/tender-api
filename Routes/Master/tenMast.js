var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _tenMast = require('../../Controllers/Master/TenMast.Controllers');

router.post('/TenMastFill', verifyToken, _tenMast.TenMastFill)
router.post('/TenMastSave', verifyToken, _tenMast.TenMastSave)
router.post('/TenMastDelete', verifyToken, _tenMast.TenMastDelete)

module.exports = router;