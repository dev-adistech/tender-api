const sql = require('mssql')

const sql_config = {
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    server: process.env.DB_SERVER,
    database: process.env.DB_DATABASE,
    connectionTimeout: 10000000,
    requestTimeout: 10000000,
    // pool: {
    //     idleTimeoutMillis: 1000000000,
    //     max: 100
    // },
    // dialectOptions: {
    //     options: { requestTimeout: 1000000000 }
    // },
    options: {
        encrypt: false,
        enableArithAbort: true,
        useUTC: true,
    },
    port: parseInt(process.env.DB_PORT),
    // parseJSON: true
}

sql.on('error', err => {
    console.log('SQL ERROR:');
    console.log(err);
})

var conx;
sql.connect(sql_config).then(pool => {
    conx = pool;
});
