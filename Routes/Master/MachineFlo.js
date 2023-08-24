var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _MacfloMast = require('../../Controllers/Master/MachineFlo.Controller');

router.post('/MachineFloMastFill', verifyToken, _MacfloMast.MachineFloMastFill)
router.post('/MachineFloMastSave', verifyToken, _MacfloMast.MachineFloMastSave)
router.post('/MachineFloMastDelete', verifyToken, _MacfloMast.MachineFloMastDelete)

module.exports = router;