var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _nodemailer = require('../../Controllers/Common/nodemailer.Controllers');

router.post('/sendHtml', verifyToken, _nodemailer.sendHtml)

module.exports = router;

