var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _RefMast = require('../../Controllers/Master/RefMast.Controllers');

router.post('/ReflectionMastFill', verifyToken, _RefMast.ReflectionMastFill)
router.post('/ReflectionMastSave', verifyToken, _RefMast.ReflectionMastSave)
router.post('/ReflectionMastDelete', verifyToken, _RefMast.ReflectionMastDelete)

module.exports = router;