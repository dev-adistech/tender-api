var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _dashboard = require('../../Controllers/Dashboard/Dashboard.Controllers');

router.post('/FillAllMaster', verifyToken, _dashboard.FillAllMaster)
router.post('/FillAllExpTendar', verifyToken, _dashboard.FillAllExpTendar)
router.post('/DashStockFill', verifyToken, _dashboard.DashStockFill)
router.post('/DashReceiveFill', verifyToken, _dashboard.DashReceiveFill)
router.post('/DashReceiveStockConfirm', verifyToken, _dashboard.DashReceiveStockConfirm)
router.post('/DashCurrentRollingFill', verifyToken, _dashboard.DashCurrentRollingFill)
router.post('/DashTodayMakeableFill', verifyToken, _dashboard.DashTodayMakeableFill)
router.post('/DashMakeableViewFill', verifyToken, _dashboard.DashMakeableViewFill)
router.post('/DashPredictionPendingFill', verifyToken, _dashboard.DashPredictionPendingFill)
router.post('/DashPredictionPendingFileDownload', _dashboard.DashPredictionPendingFileDownload)
router.post('/DashTop20MarkerFill', verifyToken, _dashboard.DashTop20MarkerFill)
router.post('/DashGrdView', verifyToken, _dashboard.DashGrdView)
router.post('/DashMrkAccView', verifyToken, _dashboard.DashMrkAccView)
router.post('/DashMrkAccViewDet', verifyToken, _dashboard.DashMrkAccViewDet)
router.post('/DashMrkPrdView', verifyToken, _dashboard.DashMrkPrdView)
router.post('/DashMrkPrdViewDet', verifyToken, _dashboard.DashMrkPrdViewDet)
router.post('/DashAttence', verifyToken, _dashboard.DashAttence)
router.post('/DashAttenceDet', verifyToken, _dashboard.DashAttenceDet)
router.post('/DashMrkAccDoubleClick', verifyToken, _dashboard.DashMrkAccDoubleClick)
router.post('/DashMrkAccPrdSave', verifyToken, _dashboard.DashMrkAccPrdSave)
router.post('/MrkAccDel', verifyToken, _dashboard.MrkAccDel)
router.post('/DashPrdPenUpd', verifyToken, _dashboard.DashPrdPenUpd)
router.post('/DashOTPrdPen', verifyToken, _dashboard.DashOTPrdPen)
router.post('/DashTop20Det', verifyToken, _dashboard.DashTop20Det)
router.post('/DashMrkAccDobDel', verifyToken, _dashboard.DashMrkAccDobDel)
router.post('/DashDailyPrc', verifyToken, _dashboard.DashDailyPrc)
router.post('/DashSarin', verifyToken, _dashboard.DashSarin)
router.post('/DashIssStock', verifyToken, _dashboard.DashIssStock)
router.post('/DashStockConfSave', verifyToken, _dashboard.DashStockConfSave)
router.post('/DashStkConfSave', verifyToken, _dashboard.DashStkConfSave)

module.exports = router;
