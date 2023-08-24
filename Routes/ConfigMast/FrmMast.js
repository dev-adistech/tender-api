var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _frmMast = require('../../Controllers/ConfigMast/FrmMast.Controllers');

router.post('/FrmMastFill', verifyToken, _frmMast.FrmMastFill)
router.post('/FrmMastSave', verifyToken, _frmMast.FrmMastSave)
router.post('/FrmMastDelete', verifyToken, _frmMast.FrmMastDelete)

module.exports = router;