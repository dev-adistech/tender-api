var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');

var _login = require('../../Controllers/Utility/Login.Controllers');

router.post('/LoginAuthentication', _login.LoginAuthentication);
router.post('/EmailSendOTP', _login.EmailSendOTP);
router.post('/UserFrmOpePer', verifyToken,_login.UserFrmOpePer);
router.post('/UserFrmPer', verifyToken,_login.UserFrmPer);
router.post('/GetEmail',_login.GetEmail);

module.exports = router;