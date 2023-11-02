var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _selldaysMast = require('../../Controllers/Master/SelldaysMast.Controller');

router.post('/SellDaysMastFill', verifyToken, _selldaysMast.SellDaysMastFill)
router.post('/SellDaysMastSave', verifyToken, _selldaysMast.SellDaysMastSave)
router.post('/SellDaysMastDelete', verifyToken, _selldaysMast.SellDaysMastDelete)

module.exports = router;