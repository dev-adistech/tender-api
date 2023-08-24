var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _versionmast = require('../../Controllers/Master/VersionMast.Controllers');

router.post('/VersionMastFill', verifyToken, _versionmast.VersionMastFill)
router.post('/VersionMastSave', verifyToken, _versionmast.VersionMastSave)
router.post('/VersionMastDelete', verifyToken, _versionmast.VersionMastDelete)

module.exports = router;