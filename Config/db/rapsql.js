const sql = require('mssql')


async function RapServerConn() {

    var RapConn = {
        user: process.env.DB_USER,
        password: process.env.DB_PASSWORD,
        server: process.env.DB_SERVER,
        database: process.env.DB_DATABASERAP,
        connectionTimeout: 3000000,
        requestTimeout: 3000000,

        options: {
            encrypt: false,
            enableArithAbort: true,
            useUTC: true,
        },
        port: parseInt(process.env.DB_PORT),
    }

    const RapServer = await new sql.ConnectionPool(RapConn).connect().then().catch(err =>
        console.log('Database Connection Failed! Bad Config: ', err))

    return RapServer
}

async function msgServerconn(){
    var MsgConn = {
        user: process.env.DB_USER,
        password: process.env.DB_PASSWORD,
        server: process.env.DB_SERVER,
        database: process.env.DB_DATABASEMSG,
        connectionTimeout: 3000000,
        requestTimeout: 3000000,

        options: {
            encrypt: false,
            enableArithAbort: true,
            useUTC: true,
        },
        port: parseInt(process.env.DB_PORT),
    }

    const msgServer = await new sql.ConnectionPool(MsgConn).connect().then().catch(err =>
        console.log('Database Connection Failed! Bad Config: ', err))

    return msgServer
}

module.exports = { RapServerConn, msgServerconn }
