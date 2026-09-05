const fs = require('node:fs');
const path = require('node:path');

const root = __dirname;
const outputPath = path.resolve(root, '../羽茜考勤小助手.html');
const htmlPath = path.join(root, 'index.html');

const escapeScript = source => source.replace(/<\/script/gi, '<\\/script');

let html = fs.readFileSync(htmlPath, 'utf8');
const css = fs.readFileSync(path.join(root, 'styles.css'), 'utf8');
const xlsx = escapeScript(fs.readFileSync(path.join(root, 'vendor/xlsx.full.min.js'), 'utf8'));
const parser = escapeScript(fs.readFileSync(path.join(root, 'parser.js'), 'utf8'));
const app = escapeScript(fs.readFileSync(path.join(root, 'app.js'), 'utf8'));

html = html
  .replace('<link rel="stylesheet" href="./styles.css" />', () => `<style>\n${css}\n</style>`)
  .replace('<script src="./vendor/xlsx.full.min.js"></script>', () => `<script>\n${xlsx}\n</script>`)
  .replace('<script src="./parser.js"></script>', () => `<script>\n${parser}\n</script>`)
  .replace('<script src="./app.js"></script>', () => `<script>\n${app}\n</script>`);

const unresolvedAsset = html.match(/<(?:link|script)\s+[^>]*(?:href|src)=["']\.\/(?:styles\.css|parser\.js|app\.js|vendor\/xlsx\.full\.min\.js)["'][^>]*>/i);
if (unresolvedAsset) {
  console.error(unresolvedAsset[0]);
  throw new Error('Standalone build still contains local runtime dependencies');
}

fs.writeFileSync(outputPath, html, 'utf8');
console.log(outputPath);
