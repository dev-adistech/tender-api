var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _ParcelEnt = require('../../Controllers/Transaction/ParcelEnt.Controllers');

router.post('/TendarParcelEnt', verifyToken, _ParcelEnt.TendarParcelEnt)
router.post('/TendarResParcelSave', verifyToken, _ParcelEnt.TendarResParcelSave)
router.post('/TendarMastDisSave', verifyToken, _ParcelEnt.TendarMastDisSave)

module.exports = router;