var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _sizeMast = require('../../Controllers/Master/SizeMast.Controllers');

router.post('/SizeMastFill', verifyToken, _sizeMast.SizeMastFill)
router.post('/SizeMastSave', verifyToken, _sizeMast.SizeMastSave)
router.post('/SizeMastDelete', verifyToken, _sizeMast.SizeMastDelete)

module.exports = router;