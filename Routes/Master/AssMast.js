var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _assMast = require('../../Controllers/Master/AssMast.Controller');

router.post('/TendarAssMstSave', verifyToken, _assMast.TendarAssMstSave)
router.post('/TendarAssEntFill', verifyToken, _assMast.TendarAssEntFill)
router.post('/FindRapAss', verifyToken, _assMast.FindRapAss)
router.post('/TendarAssPrdSave', verifyToken, _assMast.TendarAssPrdSave)
router.post('/TendarAssTrnSave', verifyToken, _assMast.TendarAssTrnSave)

module.exports = router;