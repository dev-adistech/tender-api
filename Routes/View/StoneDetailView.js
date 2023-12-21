var express = require('express');
var router = express.Router();

var { verifyToken} = require('../../Config/verify/verify');

var  _stoneDetail= require('../../Controllers/View/StoneDetailView.Controllers');

router.post('/StoneDetailView',verifyToken,_stoneDetail.StoneDetailView);
router.post('/PartyEmailIdList',verifyToken,_stoneDetail.PartyEmailIdList);
router.post('/PacketOfferView',verifyToken,_stoneDetail.PacketOfferView);
router.post('/OfferEntrySave',verifyToken,_stoneDetail.OfferEntrySave);
router.post('/DownloadDoc',_stoneDetail.DownloadDoc);

module.exports = router;