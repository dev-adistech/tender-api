var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _TendarEst = require('../../Controllers/Transaction/TendarEst.Controller');


router.post('/TendarPrdDetDisp', verifyToken, _TendarEst.TendarPrdDetDisp)
router.post('/FindRap', verifyToken, _TendarEst.FindRap)
router.post('/TendarPrdDetSave', verifyToken, _TendarEst.TendarPrdDetSave)
router.post('/TendarVidUpload', verifyToken, _TendarEst.TendarVidUpload)
router.post('/TendarVidDisp', verifyToken, _TendarEst.TendarVidDisp)
router.post('/TendarVidDelete', verifyToken, _TendarEst.TendarVidDelete)
router.post('/TendarPrdDetDock', verifyToken, _TendarEst.TendarPrdDetDock)
router.post('/TendarResSave', verifyToken, _TendarEst.TendarResSave)
router.post('/TendarApprove', verifyToken, _TendarEst.TendarApprove)
router.post('/TendarVidUploadDisp', verifyToken, _TendarEst.TendarVidUploadDisp)

module.exports = router;