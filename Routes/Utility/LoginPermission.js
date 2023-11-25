var express = require('express');
var router = express.Router();

var { verifyToken} = require('../../Config/verify/verify');

var _loginPermission = require('../../Controllers/Utility/LoginPermission.Controllers');

router.post('/CompPerMastIPFill',verifyToken,_loginPermission.CompPerMastIPFill);
router.post('/CompPerMastIPSave',verifyToken,_loginPermission.CompPerMastIPSave);
router.post('/CompPerMastIPDelete',verifyToken,_loginPermission.CompPerMastIPDelete);

module.exports = router;