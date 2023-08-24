
const express = require("express")
const bodyparser = require("body-parser")
const cors = require("cors")
const sql = require("mssql")
const path = require("path")
const https = require('https');
const fs = require("fs");
var { Export } = require('./Config/Export');


// For Live
// const httpsOptions = {
// 	cert : fs.readFileSync('./cert/apitender.cert') , 
// 	key : fs.readFileSync('./cert/apitender.pem'),
// 	ca :fs.readFileSync('./cert/apitenderc.pem')
//   };
  

require('dotenv').config({
	path: path.resolve(__dirname, `.env`)
});

const port = process.env.MAIN_PORT

var conn = require("./Config/db/sql")
var app = express();

// For Live
// const httpsServer = https.createServer(httpsOptions, app);

var routes = require("./routes")

app.use(cors())
app.use(bodyparser.json({ limit: "100mb" }))
app.use(bodyparser.urlencoded({ limit: "100mb", extended: false, parameterLimit: 100000 }))

app.set("view engine", "ejs")
app.set("views", path.join(__dirname, "/Public/views"))
app.set("report", path.join(__dirname, "/Public/report"))
app.set("images", path.join(__dirname, "/Public/images"))
app.set("file", path.join(__dirname, "/Public/file/FTP_DOWNLOAD_FILES"))

app.use("*", (req, res, next) => {
	if (
		req.headers.host.split(":")[0] == process.env.DEV_IP ||
		req.headers.host.split(":")[0] == process.env.DOMAIN_IP ||
		req.headers.host.split(":")[0] == process.env.STATIC_IP ||
		req.headers.host.split(":")[0] == process.env.DOMAIN_NAME
	) {
		next()
	}
})

app.use(express.static(path.join(__dirname, "/Public")))

app.use("/api", routes)

app.get("/", (req, res) => {
	res.send("Welcome To predacon api.")
})


//For Local
app.listen(port, () => {
	console.log(`listening at http://localhost:${port}`)
})


//For Live
// httpsServer.listen(port, 'api.tender.peacocktech.in');

app.post('/send-notification', (req, res) => {
	console.log(req.body)
	const notify = { data: req.body };
	socket.emit('notification', notify);
	res.send(notify);
});

const server = app.listen(port, process.env.DOMAIN_IP, () => {
	console.log(`Server connection on ${process.env.DOMAIN_IP}:${port}`);
});
const socket = require('socket.io')(server, {
	cors: { origin: "*" },
	path: '/socket.io',
	transports: ['websocket', 'polling'],
	secure: true,
});
socket.on('connection', socket => {
});

var packageJson = require('./package.json');
console.log(packageJson.version);
