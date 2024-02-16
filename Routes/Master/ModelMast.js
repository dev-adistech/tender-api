var express = require('express');
var router = express.Router();

var { verifyToken } = require('../../Config/verify/verify');
var _modelMast = require('../../Controllers/Master/ModelMast.Controller');

router.post('/ModelMastFill', verifyToken, _modelMast.ModelMastFill)
router.post('/ModelMastSave', verifyToken, _modelMast.ModelMastSave)
router.post('/ModelMastDelete', verifyToken, _modelMast.ModelMastDelete)

module.exports = router;