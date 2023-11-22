var nodemailer = require('nodemailer');
var path = require('path');
var fs = require('fs');
let ejs = require("ejs");

// var credential_transport = nodemailer.createTransport({
//     service: 'gmail',
//     auth: {
//         user: process.env.GMAIL_USER,
//         pass: process.env.GMAIL_PWD
//     }
// });

var credential_transport = nodemailer.createTransport({
    host: 'smtp.office365.com',
    port: 587,
    auth: {
        user: process.env.GMAIL_USER,
        pass: process.env.GMAIL_PWD
    }, 
    tls: {
        ciphers: 'SSLv3'
    }
});


function sendHtml(email, HtmlPath, templateData, subject) {

    return new Promise(function (resolve, reject) {

        return fs.readFile(HtmlPath, {
            encoding: 'utf-8'
        }, function (err, html) {
          
            if (!err) {

                var template = ejs.compile(html);

                var result = template(templateData);


                var mailOptions = {
                    from: process.env.GMAIL_USER,
                    to: email,
                    subject: subject,
                    html: result
                };
               
                return credential_transport.sendMail(mailOptions, function (error, info) {
                    if (error) {
                        resolve(false);
                    } else {
                        resolve(true);
                    }
                });
            } else {
                console.log(err)
                resolve(false);
            }
        });
    });
}

function sendSingle(mailOptions,transporter)    {
    return new Promise(function (resolve, reject) {
        return transporter.sendMail(mailOptions, function(error, info){
            if (error) {
                console.log(error)
                resolve(false);
            } else {
                resolve(true);
            }
        });
    });
}


module.exports = {sendSingle , sendHtml};