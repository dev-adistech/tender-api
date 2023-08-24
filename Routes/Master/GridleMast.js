var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _GridleMast = require('../../Controllers/Master/GridleMast.Controller');


router.post('/GirdleMastFill', verifyToken, _GridleMast.GirdleMastFill)
router.post('/GirdleMastSave', verifyToken, _GridleMast.GirdleMastSave)
router.post('/GirdleMastDelete', verifyToken, _GridleMast.GirdleMastDelete)

module.exports = router;