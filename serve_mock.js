const fs = require('fs');
const http = require('http');

let indexHtml = fs.readFileSync('index.html', 'utf8');

const includeRegex = /<\?!=\s*include\('([^']+)'\);\s*\?>/g;
indexHtml = indexHtml.replace(includeRegex, (match, p1) => {
    try {
        return fs.readFileSync(p1 + '.html', 'utf8');
    } catch (e) {
        console.error("Missing included file:", p1);
        return "";
    }
});

fs.writeFileSync('index_compiled.html', indexHtml);

const server = http.createServer((req, res) => {
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(indexHtml);
});

server.listen(3000, () => {
    console.log('Server running at http://localhost:3000/');
});
