var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _labMast = require('../../Controllers/Master/LabMast.Contoller');

router.post('/LabMastFill', verifyToken, _labMast.LabMastFill)
router.post('/LabMastSave', verifyToken, _labMast.LabMastSave)
router.post('/LabMastDelete', verifyToken, _labMast.LabMastDelete)

module.exports = router;