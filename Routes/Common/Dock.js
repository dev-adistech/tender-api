var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _dock = require('../../Controllers/Common/Dock.Controllers');

router.post('/USP_SELECT', verifyToken, _dock.USP_SELECT)
router.post('/GrpOrEmpFill', verifyToken, _dock.GrpOrEmpFill)
router.post('/DockPrcPen', verifyToken, _dock.DockPrcPen)
router.post('/DockPrcPenDet', verifyToken, _dock.DockPrcPenDet)
router.post('/DockProduction', verifyToken, _dock.DockProduction)
router.post('/DockCurRoll', verifyToken, _dock.DockCurRoll)
router.post('/DockCurRollDet', verifyToken, _dock.DockCurRollDet)
router.post('/DockDimPartyStk', verifyToken, _dock.DockDimPartyStk)
router.post('/DockGrdStk', verifyToken, _dock.DockGrdStk)
router.post('/DockChkStockTaly', verifyToken, _dock.DockChkStockTaly)
router.post('/DockRequirement', verifyToken, _dock.DockRequirement)
router.post('/DockPelProd', verifyToken, _dock.DockPelProd)
router.post('/DockPrdLock', verifyToken, _dock.DockPrdLock)
router.post('/Dock2nDUpd', verifyToken, _dock.Dock2nDUpd)
router.post('/Dock2ND', verifyToken, _dock.Dock2ND)
router.post('/DockMrkStockTaly', verifyToken, _dock.DockMrkStockTaly)
router.post('/DockFactory', verifyToken, _dock.DockFactory)
router.post('/DockLotStatus', verifyToken, _dock.DockLotStatus)
router.post('/DockLotStatusDet', verifyToken, _dock.DockLotStatusDet)
router.post('/ClvToGrdDue', verifyToken, _dock.ClvToGrdDue)

module.exports = router;
