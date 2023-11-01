var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');

var _getcerti = require('../../Controllers/Utility/getcerti.Controllers');

router.post('/LabResultGet', verifyToken, _getcerti.LabResultGet)
router.post('/LabResultImportUpdate', verifyToken, _getcerti.LabResultImportUpdate)
router.post('/GetLot', verifyToken, _getcerti.GetLot)
router.post('/GetStonId', verifyToken, _getcerti.GetStonId)
router.post('/CertiResultSave', verifyToken, _getcerti.CertiResultSave)

module.exports = router;