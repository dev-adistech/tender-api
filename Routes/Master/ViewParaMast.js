var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _viewpara = require('../../Controllers/Master/ViewParaMast.Controllers');

router.post('/ViewParaFill', verifyToken, _viewpara.ViewParaFill)
router.post('/ViewParaComboFill', verifyToken, _viewpara.ViewParaComboFill)
router.post('/ViewParaSave', verifyToken, _viewpara.ViewParaSave)
router.post('/ViewParaDelete', verifyToken, _viewpara.ViewParaDelete)

module.exports = router;