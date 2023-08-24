var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _DepthMast = require('../../Controllers/Master/DeptMast.Controllers');

router.post('/DeptMastFill', verifyToken, _DepthMast.DeptMastFill)
router.post('/DeptMastSave', verifyToken, _DepthMast.DeptMastSave)
router.post('/DeptMastDelete', verifyToken, _DepthMast.DeptMastDelete)

module.exports = router;