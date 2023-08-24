var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _MacColMast = require('../../Controllers/Master/MacColMast.Controller');


router.post('/MachineColMastFill', verifyToken, _MacColMast.MachineColMastFill)
router.post('/MachineColMastSave', verifyToken, _MacColMast.MachineColMastSave)
router.post('/MachineColMastDelete', verifyToken, _MacColMast.MachineColMastDelete)

module.exports = router;