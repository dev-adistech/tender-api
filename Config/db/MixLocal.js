var MixSalLocalSql = require('mssql')

async function MixSalLocalServerConn() {

    var MixSalLocalConn = {
        user: process.env.MIX_LOCAL_DB_USER,
        password: process.env.MIX_LOCAL_DB_PASSWORD,
        server: process.env.MIX_LOCAL_DB_SERVER,
        database: process.env.MIX_LOCAL_DB_DATABASE,
        options: {
            encrypt: true,
            enableArithAbort: true,
            useUTC: true
        }
        // connectionTimeout: 1000000,
        // requestTimeout: 1000000,
    }


    const MixSalLocalServer = await new MixSalLocalSql.ConnectionPool(MixSalLocalConn).connect().then().catch(err =>
        console.log('Database Connection Failed! Bad Config: ', err))

    return MixSalLocalServer
}

module.exports = { MixSalLocalServerConn }
