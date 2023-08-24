const sql = require("mssql");
const jwt = require('jsonwebtoken');
var CryptoJS = require("crypto-js");
// const multer = req('multer')
const fs = require('fs');
var path = require('path');
const { exec } = require("child_process");


var { _tokenSecret } = require('../../Config/token/TokenConfig.json');

exports.BarCode = async (req, res) => {

    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;

            try {
                let IP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
                var request = new sql.Request();

                request.input('INO', sql.Int, parseInt(req.body.INO))
                request.input('TPROC_CODE', sql.VarChar(5), req.body.TPROC_CODE)
                request.input('IUSER', sql.VarChar(10), req.body.IUSER)
                request.input('PNT', sql.Int, parseInt(req.body.LPnt))
                request.input('TYPE', sql.VarChar(5), req.body.TYPE)
                if (req.body.I_DATE) request.input('I_DATE', sql.DateTime2, new Date(req.body.I_DATE))
                request.input('L_CODE', sql.VarChar(7991), req.body.L_CODE)
                request.input('TAG', sql.VarChar(2), req.body.TAG)
                request.input('SRNO', sql.Int, parseInt(req.body.SRNO))
                request.input('SRNOTO', sql.Int, parseInt(req.body.SRNOTO))
                request.input('EMP_CODE', sql.VarChar(5), req.body.EMP_CODE)
                request.input('GRP', sql.VarChar(5), req.body.GRP)
                if (req.body.F_DATE) request.input('F_DATE', sql.DateTime2, new Date(req.body.F_DATE))
                if (req.body.T_DATE) request.input('T_DATE', sql.DateTime2, new Date(req.body.T_DATE))
                request.input('GRD', sql.Int, parseInt(req.body.GRD))
                request.input('SELECTMODE', sql.VarChar(5), req.body.SELECTMODE)
                request.input('DEPT_CODE', sql.VarChar(5), req.body.DEPT_CODE)
                request.input('IP', sql.VarChar(25), IP)
                // console.log(IP)
                request = await request.execute('USP_BarCode');
                if (request.recordset) {
                    // console.log(request.recordset)
                    let text = '';
                    for (var i = 0; i < request.recordset.length; i++) {
                        text = text + request.recordset[i].BCODE + '\n';
                    }
                    // let filePath = path.join(__dirname, `../../Public/file/BarCode/BARCODE-${Date.now()}.txt`);
                    let timeStemp = Date.now();
                    console.log(process)
                    let filePath = `${process.env.BARCODE_PATH}/BARCODE-${timeStemp}.txt`;
                    let PrinterPath = request.recordsets[1][0] ? request.recordsets[1][0].PRN_NAME : '';
                    let message = '';

                    if (!PrinterPath) {
                        message = "Printer Path not found."
                    }
                    let printPath = `${process.env.PRINT_BARCODE_PATH}\\BARCODE-${timeStemp}.txt`
                    let printCmd = `Type ${printPath} > ${PrinterPath}`;
                    await fs.writeFile(filePath, text, async (err) => {
                        if (err)
                            console.log("File write error", err);
                        else {
                            console.log(printCmd)
                            await exec(printCmd, async (error, stdout, stderr) => {
                                if (error) {
                                    console.log("error ::::", error)
                                    message = error.message;
                                }
                                if (stderr) {
                                    console.log("stderr ::::", stderr)

                                    message = stderr;
                                }
                                console.log("stdout ::::", stdout)

                                message = stdout;
                                fs.unlinkSync(filePath);
                            });
                        }
                    });

                    res.json({ success: 1, data: request.recordset, message: message })
                } else {
                    res.json({ success: 1, data: "Not Found" })
                }

            } catch (err) {
                console.log(err)
                res.json({ success: 0, data: err })
            }
        }
    });
}

exports.BGImageUpload = async (req, res) => {

    let fileNameAry = [];
    var multer = require('multer');
    const storage = multer.diskStorage({
        destination: "public/images/bg",
        filename: function (req, file, cb) {

            const orgFileName = file.originalname
            const [ext] = orgFileName.split('.').reverse();

            let fileName = 'background.' + ext;
            fileNameAry.push(fileName)
            cb(null, fileName)
        }
    });

    const upload = multer({ storage: storage }).fields([{ name: "BGImage" }]);

    upload(req, res, (err) => {
        res.json(fileNameAry)
        if (err) throw err;
    });

}

exports.Export = () => {
    const watcher = chokidar.watch('D:/Export', {
        ignored: /(^|[\/\\])\../, // ignore dotfiles
        persistent: true
    });

    // const path = 'D:/Export';
    const log = console.log.bind(console);
    let enableWatcher = false

    watcher
        .on('ready', () => {
            enableWatcher = true
            log('Initial scan complete. Ready for changes')
        })
        .on('add', path => {
            log(`File ${path} has been added`)
            if (enableWatcher === true) {
                fs.readFile(path, 'utf8', function (err, data) {
                    // Display the file content
                    log(`-------------------------------`)
                    log(`Data from file ${path}: `)
                    log(`-------------------------------`)
                    console.log(data);
                });
            }
        })
        .on('change', path => log(`File ${path} has been changed`))
        .on('unlink', path => log(`File ${path} has been removed`));
}


exports.PassUpdate = async (req, res) => {
    jwt.verify(req.token, _tokenSecret, async (err, authData) => {
        if (err) {
            res.sendStatus(401);
        } else {
            const TokenData = await authData;
            try {
                var request = new sql.Request();

                request.input('USERID', sql.VarChar(10), req.body.USERID)
                request.input('U_PASS', sql.VarChar(256), CryptoJS.AES.encrypt(req.body.U_PASS, process.env.PASS_KEY).toString())

                request = await request.execute('USP_PassUpdate')

                res.json({ success: 1, data: '' })


            } catch (err) {
                res.json({ success: 0, data: err })
                console.log(err)
            }
        }
    })

}


const isFalsy = value => !value;
const isWhitespaceString = value =>
    typeof value === 'string' && /^\s*$/.test(value);
const isEmptyCollection = value =>
    (Array.isArray(value) || value === Object(value)) &&
    !Object.keys(value).length;
const isInvalidDate = value =>
    value instanceof Date && Number.isNaN(value.getTime());
const isEmptySet = value => value instanceof Set && value.size === 0;
const isEmptyMap = value => value instanceof Map && value.size === 0;

const isBlank = value => {
    if (isFalsy(value)) return true;
    if (isWhitespaceString(value)) return true;
    if (isEmptyCollection(value)) return true;
    if (isInvalidDate(value)) return true;
    if (isEmptySet(value)) return true;
    if (isEmptyMap(value)) return true;
    return false;
};

isBlank(null); // true
isBlank(undefined); // true
isBlank(0); // true
isBlank(false); // true
isBlank(''); // true
isBlank(' \r\n '); // true
isBlank(NaN); // true
isBlank([]); // true
isBlank({}); // true
isBlank(new Date('hello')); // true
isBlank(new Set()); // true
isBlank(new Map()); // true
