
const express = require("express")
const bodyparser = require("body-parser")
const cors = require("cors")
const sql = require("mssql")
const path = require("path")
var { Export } = require('./Config/Export')

require('dotenv').config({
	path: path.resolve(__dirname, `.env`)
});

const port = process.env.MAIN_PORT

var conn = require("./Config/db/sql")
var app = express()

var routes = require("./routes")

app.use(cors())
app.use(bodyparser.json({ limit: "100mb" }))
app.use(bodyparser.urlencoded({ limit: "100mb", extended: false, parameterLimit: 100000 }))
// app.use(bodyparser.json({ limit: "100mb" }))
// app.use(bodyparser.urlencoded({ limit: "200mb", extended: false }))

// app.use(bodyparser.json({limit:1024*1024*20}));
// app.use(bodyparser.urlencoded({extended:true,limit:1024*1024*20 }));

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

app.listen(port, () => {
	console.log(`listening at http://localhost:${port}`)
})

// app.listen(port, '192.168.1.200', () => {
//   console.log(`listening at http://localhost:${port}`)
// })

//* notification
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
	// console.log('Socket: client connected');
});

// const socket = require('socket.io')(server, { cors: { origin: "*" } });

// socket.on('connection', socket => {
// 	// console.log('Socket: client connected');
// });

//* get version information from package.json.
var packageJson = require('./package.json');
console.log(packageJson.version);

//* file export function
// Export();
