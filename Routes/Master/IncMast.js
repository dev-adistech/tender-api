var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _incMast = require('../../Controllers/Master/IncMast.Controllers');

router.post('/IncMastFill', verifyToken, _incMast.IncMastFill)
router.post('/IncMastSave', verifyToken, _incMast.IncMastSave)
router.post('/IncMastDelete', verifyToken, _incMast.IncMastDelete)
router.post('/IncMastGetMaxId', verifyToken, _incMast.IncMastGetMaxId)

module.exports = router;