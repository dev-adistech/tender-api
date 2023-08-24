var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _quaMast = require('../../Controllers/Master/QuaMast.Controllers');

router.post('/QuaMastFill', verifyToken, _quaMast.QuaMastFill)
router.post('/QuaMastSave', verifyToken, _quaMast.QuaMastSave)
router.post('/QuaMastDelete', verifyToken, _quaMast.QuaMastDelete)

module.exports = router;