module.exports = {
  apps: [
    {
      script: 'index.js',
      autorestart: true,
      watch: true,
      name: 'Tendar-api',
      instances: 50,
      exec_mode: "cluster",
    }
    
  ],
};
