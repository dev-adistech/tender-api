var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _userMast = require('../../Controllers/ConfigMast/UserMast.Controllers');

router.post('/UserMastFill', verifyToken, _userMast.UserMastFill)
router.post('/UserMastSave', verifyToken, _userMast.UserMastSave)
router.post('/UserMastDelete', verifyToken, _userMast.UserMastDelete)
router.post('/UserCreate', verifyToken, _userMast.UserCreate)

module.exports = router;
