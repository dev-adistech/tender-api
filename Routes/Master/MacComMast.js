var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _MacComMast = require('../../Controllers/Master/MacComMast.Controller');


router.post('/MachComMastFill', verifyToken, _MacComMast.MachComMastFill)
router.post('/MachComMastSave', verifyToken, _MacComMast.MachComMastSave)
router.post('/MachComMastDelete', verifyToken, _MacComMast.MachComMastDelete)

module.exports = router;