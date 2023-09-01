var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _view = require('../../Controllers/View/View.Controllers');

router.post('/PricingWrk', verifyToken, _view.PricingWrk)
router.post('/PricingWrkDisp', verifyToken, _view.PricingWrkDisp)
router.post('/PricingWrkMperSave', verifyToken, _view.PricingWrkMperSave)
router.post('/BVView', verifyToken, _view.BVView)
router.post('/BidDataView', verifyToken, _view.BidDataView)

module.exports = router;
