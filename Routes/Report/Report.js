var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _report = require('../../Controllers/Report/Report.Controllers');

router.post('/ReportPrint', _report.ReportPrint)
router.post('/ReportViewer', _report.ReportViewer)
router.post('/FetchReportData', verifyToken, _report.FetchReportData)
router.post('/FillReportList', verifyToken, _report.FillReportList)
router.post('/GetReportParams', verifyToken, _report.GetReportParams)
router.post('/RPTDATA', verifyToken, _report.RPTDATA)
router.post('/FrmRptDataFill', verifyToken, _report.FrmRptDataFill)

module.exports = router;