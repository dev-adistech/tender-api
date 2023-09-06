var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _floMast = require('../../Controllers/Master/FloMast.Controllers');

router.post('/FloMastFill', verifyToken, _floMast.FloMastFill)
router.post('/FloMastSave', verifyToken, _floMast.FloMastSave)
router.post('/FloMastDelete', verifyToken, _floMast.FloMastDelete)
router.post('/FloMastsmipleFill', verifyToken, _floMast.FloMastsmipleFill)
router.post('/FloMastsmipleSave', verifyToken, _floMast.FloMastsmipleSave)
router.post('/FloMastsmipleDelete', verifyToken, _floMast.FloMastsmipleDelete)

module.exports = router;