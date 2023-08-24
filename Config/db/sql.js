const sql = require('mssql')

const sql_config = {
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    server: process.env.DB_SERVER,
    database: process.env.DB_DATABASE,
    connectionTimeout: 300000,
    requestTimeout: 300000,
    pool: {
        idleTimeoutMillis: 300000,
        max: 100
    },
    dialectOptions: {
        options: { requestTimeout: 300000 }
    },
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
