var express = require('express');
var router = express.Router();

var Common = require('./Routes/Common/Common')
var nodemailer = require('./Routes/Common/nodemailer')
var Dock = require('./Routes/Common/Dock')

var FrmMast = require('./Routes/ConfigMast/FrmMast')
var Makreturn = require('./Routes/ConfigMast/Makreturn')
var LoginDetail = require('./Routes/ConfigMast/LoginDetail')
var PerMast = require('./Routes/ConfigMast/PerMast')
var RptMast = require('./Routes/ConfigMast/RptMast')
var RptParaMast = require('./Routes/ConfigMast/RptParaMast')
var RptPerMast = require('./Routes/ConfigMast/RptPerMast')
var UserCatMast = require('./Routes/ConfigMast/UserCatMast')
var UserMast = require('./Routes/ConfigMast/UserMast')
var LotUtility = require('./Routes/ConfigMast/LotUtility')

var Dashboard = require('./Routes/Dashboard/Dashboard')

var ColMast = require('./Routes/Master/ColMast')
var AssMast = require('./Routes/Master/AssMast')
var ModelMast = require('./Routes/Master/ModelMast')
var SelldaysMast = require('./Routes/Master/SelldaysMast')
var ShdMast = require('./Routes/Master/ShdMast')
var RefMast = require('./Routes/Master/RefMast')
var tenMast = require('./Routes/Master/tenMast')
var MacColMast = require('./Routes/Master/MacColMast')
var CutMast = require('./Routes/Master/CutMast')
var FloMast = require('./Routes/Master/FloMast')
var MachineFloMast = require('./Routes/Master/MachineFlo')
var MacComMast = require('./Routes/Master/MacComMast')
var IncMast = require('./Routes/Master/IncMast')
var IncTypeMast = require('./Routes/Master/IncTypeMast')
var QuaMast = require('./Routes/Master/QuaMast')
var ShpMast = require('./Routes/Master/ShpMast')
var SizeMast = require('./Routes/Master/SizeMast')
var LabMast = require('./Routes/Master/LabMast')
var ViewParaMast = require('./Routes/Master/ViewParaMast')
var MilkyMast = require('./Routes/Master/MilkyMast')
var GridleMast = require('./Routes/Master/GridleMast')
var DepthMast = require('./Routes/Master/DepthMast')
var RatioMast = require('./Routes/Master/RatioMast')

var RapCutDisc = require('./Routes/Rap/RapCutDisc')
var RapFloDisc = require('./Routes/Rap/RapFloDisc')
var RapIncDisc = require('./Routes/Rap/RapIncDisc')
var RapMast = require('./Routes/Rap/RapMast')
var RapOrg = require('./Routes/Rap/RapOrg')
var RapParamDisc = require('./Routes/Rap/RapParamDisc')
var RapShdDisc = require('./Routes/Rap/RapShdDisc')

var Report = require('./Routes/Report/Report')

var Login = require('./Routes/Utility/Login')
var getcerti = require('./Routes/Utility/getcerti')
var rateupdate = require('./Routes/Utility/rateupdate')

var View = require('./Routes/View/View')
var StoneDetailView = require('./Routes/View/StoneDetailView')
var TendarCom = require('./Routes/Transaction/TendarCom')
var TendarMast = require('./Routes/Transaction/TendarMast')
var TendarEst = require('./Routes/Transaction/TendarEst')
var LotMap = require('./Routes/Transaction/LotMap')
var ExpDate = require('./Routes/Transaction/ExpDate')
var ParcelEnt = require('./Routes/Transaction/ParcelEnt')
var LoginPermission = require('./Routes/Utility/LoginPermission')
var RapCalc = require('./Routes/Rap/RapCalc')

router.use('/Common', Common);
router.use('/nodemailer', nodemailer);
router.use('/Dock', Dock);

router.use('/FrmMast', FrmMast);
router.use('/Makreturn', Makreturn);
router.use('/LoginDetail', LoginDetail);
router.use('/PerMast', PerMast);
router.use('/RptMast', RptMast);
router.use('/RptParaMast', RptParaMast);
router.use('/RptPerMast', RptPerMast);
router.use('/UserCatMast', UserCatMast);
router.use('/UserMast', UserMast);
router.use('/LotUtility', LotUtility);

router.use('/Dashboard', Dashboard);

router.use('/ColMast', ColMast);
router.use('/AssMast', AssMast);
router.use('/ModelMast', ModelMast);
router.use('/SelldaysMast', SelldaysMast);
router.use('/ShdMast', ShdMast);
router.use('/RefMast', RefMast);
router.use('/tenMast', tenMast);
router.use('/CutMast', CutMast);
router.use('/FloMast', FloMast);
router.use('/MachineFloMast', MachineFloMast);
router.use('/IncMast', IncMast);
router.use('/IncTypeMast', IncTypeMast);
router.use('/QuaMast', QuaMast);
router.use('/ShpMast', ShpMast);
router.use('/SizeMast', SizeMast);
router.use('/LabMast', LabMast);
router.use('/ViewParaMast', ViewParaMast);
router.use('/MacColMast', MacColMast);
router.use('/MacComMast', MacComMast);
router.use('/MilkyMast', MilkyMast);
router.use('/GridleMast', GridleMast);
router.use('/DepthMast', DepthMast);
router.use('/RatioMast', RatioMast);

router.use('/RapCutDisc', RapCutDisc);
router.use('/RapFloDisc', RapFloDisc);
router.use('/RapIncDisc', RapIncDisc);
router.use('/RapMast', RapMast);
router.use('/RapOrg', RapOrg);
router.use('/RapParamDisc', RapParamDisc);
router.use('/RapShdDisc', RapShdDisc);
router.use('/RapCalc', RapCalc);

router.use('/Report', Report);
router.use('/TendarCom', TendarCom);
router.use('/TendarMast', TendarMast);
router.use('/TendarEst', TendarEst);
router.use('/LotMap', LotMap);
router.use('/ExpDate', ExpDate);
router.use('/ParcelEnt', ParcelEnt);

router.use('/Login', Login);
router.use('/getcerti', getcerti);
router.use('/rateupdate', rateupdate);
router.use('/LoginPermission',LoginPermission);

router.use('/View', View);
router.use('/StoneDetailView', StoneDetailView);

module.exports = router;
