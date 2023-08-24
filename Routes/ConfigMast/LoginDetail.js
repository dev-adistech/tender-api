var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _loginDetail = require('../../Controllers/ConfigMast/LoginDetail.Controllers');

router.post('/LoginDetailFill', verifyToken, _loginDetail.LoginDetailFill)
router.post('/LoginDetailUpdate', _loginDetail.LoginDetailUpdate)
router.post('/LoginDetailSave', verifyToken, _loginDetail.LoginDetailSave)

module.exports = router;
