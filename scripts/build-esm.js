const fs = require('fs');
const path = require('path');

// Create ESM version of the built files
const distDir = path.join(__dirname, '..', 'dist');
const esmDir = path.join(distDir, 'esm');

if (fs.existsSync(esmDir)) {
    // Copy built files to ESM directory and rename main file
    const files = fs.readdirSync(esmDir);

    for (const file of files) {
        if (file.endsWith('.js')) {
            const source = path.join(esmDir, file);
            const target = path.join(distDir, file.replace('.js', '.esm.js'));

            let content = fs.readFileSync(source, 'utf8');

            // Fix relative imports to include .js extension for ESM
            content = content.replace(/from ['"](\.[^'"]+)['"]/g, (match, importPath) => {
                if (!importPath.endsWith('.js') && !importPath.endsWith('.json')) {
                    return match.replace(importPath, importPath + '.js');
                }
                return match;
            });

            // Also fix require statements that might exist
            content = content.replace(/require\(['"](\.[^'"]+)['"]\)/g, (match, importPath) => {
                if (!importPath.endsWith('.js') && !importPath.endsWith('.json')) {
                    return match.replace(importPath, importPath + '.js');
                }
                return match;
            });

            fs.writeFileSync(target, content, 'utf8');
            console.log(`Created ESM version: ${target}`);
        }
    }

    // Clean up temporary ESM directory
    fs.rmSync(esmDir, { recursive: true, force: true });
}

console.log('ESM build completed');