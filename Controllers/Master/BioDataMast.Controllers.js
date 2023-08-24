const sql = require("mssql");
const jwt = require('jsonwebtoken');
var { _tokenSecret } = require('../../Config/token/TokenConfig.json');
const https = require('https');
const http = require('http');

exports.BioDataFill = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                // console.log(req.body)
                request.input('DEPT_CODE', sql.NVarChar(100), req.body.DEPT_CODE)
                request.input('IDCODE', sql.NVarChar(100), req.body.IDCODE)
                request.input('TRPID', sql.NVarChar(100), req.body.TRPID)


                request = await request.execute('USP_BioDataFill');
                // console.log(request)
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}


exports.BioDataSave = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.NVarChar(100), req.body.DEPT_CODE)
                request.input('ID', sql.Int, parseInt(req.body.ID))
                request.input('IDCODE', sql.NVarChar(100), req.body.IDCODE)
                request.input('NAME', sql.NVarChar(100), req.body.NAME)
                request.input('DEPT', sql.NVarChar(100), req.body.DEPT)
                request.input('BDATE', sql.DateTime2, new Date(req.body.BDATE))
                request.input('TRPID', sql.NVarChar(100), req.body.TRPID)
                request.input('CABIN', sql.NVarChar(100), req.body.CABIN)
                if (req.body.JDATE) { request.input('JDATE', sql.DateTime2, new Date(req.body.JDATE)) }
                if (req.body.LDATE) { request.input('LDATE', sql.DateTime2, new Date(req.body.LDATE)) }
                request.input('AGE', sql.NVarChar(100), req.body.AGE)
                request.input('MARRIED', sql.NVarChar(100), req.body.MARRIED)
                request.input('RENT', sql.NVarChar(100), req.body.RENT)

                request.input('ADDICTION', sql.NVarChar(100), req.body.ADDICTION)
                request.input('ADDRESS', sql.NVarChar(250), req.body.ADDRESS)
                request.input('AREA', sql.NVarChar(100), req.body.AREA)
                request.input('CITY', sql.NVarChar(100), req.body.CITY)
                request.input('HEIGHT', sql.NVarChar(100), req.body.HEIGHT)
                request.input('WEIGHT', sql.NVarChar(100), req.body.WEIGHT)
                request.input('BLOOD', sql.NVarChar(100), req.body.BLOOD)
                request.input('EDU', sql.NVarChar(100), req.body.EDU)
                request.input('PHNO1', sql.NVarChar(10), req.body.PHNO1)
                request.input('PHNO2', sql.NVarChar(10), req.body.PHNO2)
                request.input('EXP', sql.NVarChar(100), req.body.EXP)

                request.input('SKILL', sql.NVarChar(100), req.body.SKILL)
                request.input('HOBBY', sql.NVarChar(100), req.body.HOBBY)
                request.input('VEHICLE', sql.NVarChar(100), req.body.VEHICLE)
                request.input('IDENTITI', sql.NVarChar(100), req.body.IDENTITI)
                request.input('LSTPLACE', sql.NVarChar(100), req.body.LSTPLACE)
                request.input('WRK', sql.NVarChar(100), req.body.WRK)
                request.input('REASON', sql.NVarChar(100), req.body.REASON)
                request.input('SALARY', sql.NVarChar(10), req.body.SALARY)

                request.input('HUSWF', sql.NVarChar(100), req.body.HUSWF)
                if (req.body.HUSWFBDATE) { request.input('HUSWFBDATE', sql.DateTime2, new Date(req.body.HUSWFBDATE)) }
                request.input('HUSWFEDU', sql.NVarChar(100), req.body.HUSWFEDU)
                request.input('CHILD1', sql.NVarChar(100), req.body.CHILD1)
                if (req.body.CHILD1BDATE) { request.input('CHILD1BDATE', sql.DateTime2, new Date(req.body.CHILD1BDATE)) }
                request.input('CHILD1EDU', sql.NVarChar(100), req.body.CHILD1EDU)
                request.input('CHILD2', sql.NVarChar(100), req.body.CHILD2)
                if (req.body.CHILD2BDATE) { request.input('CHILD2BDATE', sql.DateTime2, new Date(req.body.CHILD2BDATE)) }
                request.input('CHILD2EDU', sql.NVarChar(100), req.body.CHILD2EDU)
                request.input('CHILD3', sql.NVarChar(100), req.body.CHILD3)
                if (req.body.CHILD3BDATE) { request.input('CHILD3BDATE', sql.DateTime2, new Date(req.body.CHILD3BDATE)) }
                request.input('CHILD3EDU', sql.NVarChar(100), req.body.CHILD3EDU)
                request.input('FNAME', sql.NVarChar(100), req.body.FNAME)
                request.input('FFNAME', sql.NVarChar(100), req.body.FFNAME)
                request.input('FSNAME', sql.NVarChar(100), req.body.FSNAME)
                request.input('FRELATIVE', sql.NVarChar(100), req.body.FRELATIVE)
                request.input('FBUSINESS', sql.NVarChar(100), req.body.FBUSINESS)


                request.input('FCAST', sql.NVarChar(100), req.body.FCAST)
                request.input('FNATIVE', sql.NVarChar(100), req.body.FNATIVE)
                request.input('FADDRESS', sql.NVarChar(250), req.body.FADDRESS)
                request.input('FPHNO', sql.NVarChar(10), req.body.FPHNO)
                request.input('MNAME', sql.NVarChar(100), req.body.MNAME)
                request.input('MFNAME', sql.NVarChar(100), req.body.MFNAME)
                request.input('MSNAME', sql.NVarChar(100), req.body.MSNAME)
                request.input('MRELATIVE', sql.NVarChar(100), req.body.MRELATIVE)
                request.input('MBUSINESS', sql.NVarChar(100), req.body.MBUSINESS)
                request.input('MCAST', sql.NVarChar(100), req.body.MCAST)
                request.input('MNATIVE', sql.NVarChar(100), req.body.MNATIVE)
                request.input('MADDRESS', sql.NVarChar(250), req.body.MADDRESS)
                request.input('MPHNO', sql.NVarChar(10), req.body.MPHNO)
                request.input('R1NAME', sql.NVarChar(100), req.body.R1NAME)
                request.input('R1FNAME', sql.NVarChar(100), req.body.R1FNAME)
                request.input('R1SNAME', sql.NVarChar(100), req.body.R1SNAME)
                request.input('R1RELATIVE', sql.NVarChar(100), req.body.R1RELATIVE)
                request.input('R1BUSINESS', sql.NVarChar(100), req.body.R1BUSINESS)
                request.input('R1CAST', sql.NVarChar(100), req.body.R1CAST)


                request.input('R1NATIVE', sql.NVarChar(100), req.body.R1NATIVE)
                request.input('R1ADDRESS', sql.NVarChar(250), req.body.R1ADDRESS)
                request.input('R1PHNO', sql.NVarChar(10), req.body.R1PHNO)
                request.input('R2NAME', sql.NVarChar(100), req.body.R2NAME)
                request.input('R2FNAME', sql.NVarChar(100), req.body.R2FNAME)
                request.input('R2SNAME', sql.NVarChar(100), req.body.R2SNAME)
                request.input('R2RELATIVE', sql.NVarChar(100), req.body.R2RELATIVE)
                request.input('R2BUSINESS', sql.NVarChar(100), req.body.R2BUSINESS)
                request.input('R2CAST', sql.NVarChar(100), req.body.R2CAST)
                request.input('R2NATIVE', sql.NVarChar(100), req.body.R2NATIVE)
                request.input('R2ADDRESS', sql.NVarChar(250), req.body.R2ADDRESS)
                request.input('R2PHNO', sql.NVarChar(10), req.body.R2PHNO)
                request.input('R3NAME', sql.NVarChar(100), req.body.R3NAME)
                request.input('R3FNAME', sql.NVarChar(100), req.body.R3FNAME)
                request.input('R3SNAME', sql.NVarChar(100), req.body.R3SNAME)
                request.input('R3RELATIVE', sql.NVarChar(100), req.body.R3RELATIVE)
                request.input('R3BUSINESS', sql.NVarChar(100), req.body.R3BUSINESS)
                request.input('R3CAST', sql.NVarChar(100), req.body.R3CAST)
                request.input('R3NATIVE', sql.NVarChar(100), req.body.R3NATIVE)
                request.input('R3ADDRESS', sql.NVarChar(250), req.body.R3ADDRESS)
                request.input('R3PHNO', sql.NVarChar(10), req.body.R3PHNO)
                request.input('R4NAME', sql.NVarChar(100), req.body.R4NAME)
                request.input('R4FNAME', sql.NVarChar(100), req.body.R4FNAME)
                request.input('R4SNAME', sql.NVarChar(100), req.body.R4SNAME)
                request.input('R4RELATIVE', sql.NVarChar(100), req.body.R4RELATIVE)

                request.input('R4BUSINESS', sql.NVarChar(100), req.body.R4BUSINESS)
                request.input('R4CAST', sql.NVarChar(100), req.body.R4CAST)
                request.input('R4NATIVE', sql.NVarChar(100), req.body.R4NATIVE)
                request.input('R4ADDRESS', sql.NVarChar(250), req.body.R4ADDRESS)
                request.input('R4PHNO', sql.NVarChar(10), req.body.R4PHNO)
                request.input('HASTKNAME', sql.NVarChar(100), req.body.HASTKNAME)
                request.input('HASTKFNAME', sql.NVarChar(100), req.body.HASTKFNAME)
                request.input('HASTKSNAME', sql.NVarChar(100), req.body.HASTKSNAME)
                request.input('HASTKRELATIVE', sql.NVarChar(100), req.body.HASTKRELATIVE)
                request.input('HASTKBUSINESS', sql.NVarChar(100), req.body.HASTKBUSINESS)
                request.input('HASTKCAST', sql.NVarChar(100), req.body.HASTKCAST)
                request.input('HASTKNATIVE', sql.NVarChar(100), req.body.HASTKNATIVE)
                request.input('HASTKADDRESS', sql.NVarChar(250), req.body.HASTKADDRESS)
                request.input('HASTKPHNO', sql.NVarChar(10), req.body.HASTKPHNO)
                request.input('SHRT_CODE', sql.NVarChar(10), req.body.SHRT_CODE)

                request.input('IMG', sql.VarChar(250), req.body.IMG)
                request.input('DOC1', sql.VarChar(250), req.body.DOC1)
                request.input('DOC2', sql.VarChar(250), req.body.DOC2)
                request.input('DOC3', sql.VarChar(250), req.body.DOC3)
                request.input('DESIG', sql.VarChar(50), req.body.DESIG)

                request = await request.execute('USP_FrmBioDataSave');

                res.json({ success: 1, data: '' })

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    });
}


exports.BioDataDelete = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();

                request.input('DEPT_CODE', sql.NVarChar(100), req.body.DEPT_CODE)
                request.input('IDCODE', sql.NVarChar(100), req.body.IDCODE)

                request = await request.execute('USP_BioDataDelete');

                res.json({ success: 1, data: '' })

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}


exports.BioDataImgFill = async (req, res) => {
    // console.log("210",req.body)
    const fs = require('fs');
    let files1 = '';
    fs.readdir('public/images/biodata/bio-data-images', (err, files) => {
        if (err) {
            console.log(err)
        }
        let fl = files
        // console.log(files)
        for (let i = 0; i < fl.length; i++) {
            // console.log(fl[i].split('.'))
            let fls = fl[i].split('.')[0]
            if (fls.includes(req.body.IMGID)) {
                files1 = fl[i]
                // console.log("225",fl)
                // console.log("226",files1)
            }
        }
        // console.log("228",files1)
        res.json({ success: 1, data: files1 })
    })
}

exports.BioDataDocument = async (req, res) => {
    const fs = require('fs');
    // console.log(req.body)
    let files1 = []
    fs.readdir('public/images/biodata/bio-data-documents', (err, files) => {
        if (err) {
            console.log(err)
        }
        let doc = files;
        for (let i = 0; i < doc.length; i++) {
            if (doc[i].includes(req.body.IMGDOC)) {
                // files1 = doc[i]
                files1.push(doc[i])
                // console.log("245",doc[i])
            }
        }
        // console.log(files)
        // console.log("249",files1)
        res.json({ success: 1, data: files1 })
    })
}

exports.FileSearch = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                var request = new sql.Request();
                // console.log(req.body)
                request.input('FILENAME', sql.VarChar(50), req.body.FILENAME)


                request = await request.execute('usp_FileSearch');
                // console.log(request)

                // const path = 'file://192.168.1.107/d/o__file/PC-115';

                // exec(`explorer ${path}`, (error, stdout, stderr) => {
                //     if (error) {
                //         console.error('Error:', error);
                //     }
                // })
                if (request.recordset) {
                    res.json({ success: 1, data: request.recordset })
                } else {
                    res.json({ success: 0, data: "Not Found" })
                }

            } catch (err) {
                res.json({ success: 0, data: err })
            }
        }
    });
}

const path = require('path');
const express = require('express');
const app = express();

app.use(express.json());

exports.FileExploreropen = async (req, res) => {
    try {
        const FINALFILENAME = req.body.F_NAME;
        const filePath = path.resolve(__dirname, FINALFILENAME);
        res.download(filePath, req.body.FILE, (err) => {
            if (err) {
                res.status(500).json('FILE NOT FOUND');
            }
        });

    } catch (err) {
        console.log(err)
        res.json({ success: 0, data: err });
    }
}



