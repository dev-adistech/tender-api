var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _userCatMast = require('../../Controllers/ConfigMast/UserCatMast.Controllers');

router.post('/UserCatMastFill', verifyToken, _userCatMast.UserCatMastFill)
router.post('/UserCatMastSave', verifyToken, _userCatMast.UserCatMastSave)
router.post('/UserCatMastDelete', verifyToken, _userCatMast.UserCatMastDelete)

module.exports = router;