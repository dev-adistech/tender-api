var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _incTypeMast = require('../../Controllers/Master/IncTypeMast.Controllers');

router.post('/IncTypeMastFill', verifyToken, _incTypeMast.IncTypeMastFill)
router.post('/IncTypeMastSave', verifyToken, _incTypeMast.IncTypeMastSave)
router.post('/IncTypeMastDelete', verifyToken, _incTypeMast.IncTypeMastDelete)

module.exports = router;