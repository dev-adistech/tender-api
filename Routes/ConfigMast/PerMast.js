var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _perMast = require('../../Controllers/ConfigMast/PerMast.Controllers');

router.post('/PerMastFill', verifyToken, _perMast.PerMastFill)
router.post('/PerMastSave', verifyToken, _perMast.PerMastSave)
router.post('/PerMastDelete', verifyToken, _perMast.PerMastDelete)
router.post('/PerMastCopy', verifyToken, _perMast.PerMastCopy)

module.exports = router;