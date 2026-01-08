const https = require('https');
const fs = require('fs');
const path = require('path');
const devCerts = require('office-addin-dev-certs');

const port = 3000;

const mimeTypes = {
    '.html': 'text/html',
    '.js': 'text/javascript',
    '.css': 'text/css',
    '.json': 'application/json',
    '.png': 'image/png',
    '.jpg': 'image/jpg',
    '.gif': 'image/gif',
    '.svg': 'image/svg+xml',
    '.xml': 'application/xml',
    '.ico': 'image/x-icon',
};

async function startServer() {
    try {
        // 获取开发证书选项
        const options = await devCerts.getHttpsServerOptions();
        
        const server = https.createServer(options, (req, res) => {
            // 处理 CORS
            res.setHeader('Access-Control-Allow-Origin', '*');
            res.setHeader('Access-Control-Request-Method', '*');
            res.setHeader('Access-Control-Allow-Methods', 'OPTIONS, GET');
            res.setHeader('Access-Control-Allow-Headers', '*');
            // 绕过 ngrok 免费版的警告页面
            res.setHeader('ngrok-skip-browser-warning', 'true');
            
            if ( req.method === 'OPTIONS' ) {
                res.writeHead(200);
                res.end();
                return;
            }

            let filePath = '.' + req.url.split('?')[0]; // 去掉 query string
            if (filePath === './') {
                filePath = './taskpane.html';
            }

            const extname = path.extname(filePath);
            const contentType = mimeTypes[extname] || 'application/octet-stream';

            fs.readFile(filePath, (error, content) => {
                if (error) {
                    if(error.code == 'ENOENT'){
                        console.log(`404: ${filePath}`);
                        res.writeHead(404);
                        res.end('File not found');
                    } else {
                        console.error(`500: ${filePath}`, error);
                        res.writeHead(500);
                        res.end('Internal Server Error: '+error.code);
                    }
                } else {
                    console.log(`200: ${filePath}`);
                    res.writeHead(200, { 'Content-Type': contentType });
                    res.end(content, 'utf-8');
                }
            });
        });

        server.listen(port, () => {
            console.log(`Server running at https://localhost:${port}/`);
            console.log('Press Ctrl+C to stop');
        });

    } catch (err) {
        console.error("Error starting server:", err);
        console.error("Please try running: npx office-addin-dev-certs install");
    }
}

startServer();

