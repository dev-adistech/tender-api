var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _ShdMast = require('../../Controllers/Master/ShdMast.Controllers');

router.post('/ShdMastFill', verifyToken, _ShdMast.ShdMastFill)
router.post('/ShdMastSave', verifyToken, _ShdMast.ShdMastSave)
router.post('/ShdMastDelete', verifyToken, _ShdMast.ShdMastDelete)

module.exports = router;