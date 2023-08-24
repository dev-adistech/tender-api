var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');

var _login = require('../../Controllers/Utility/Login.Controllers');

router.post('/LoginAuthentication', _login.LoginAuthentication);
router.post('/UserFrmOpePer', verifyToken,_login.UserFrmOpePer);
router.post('/UserFrmPer', verifyToken,_login.UserFrmPer);

module.exports = router;