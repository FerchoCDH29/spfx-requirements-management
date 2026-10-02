const fs = require('fs');
const path = require('path');

const packagePath = path.resolve(__dirname, '../package.json');
const solutionPath = path.resolve(__dirname, '../config/package-solution.json');

const packageJson = JSON.parse(
  fs.readFileSync(packagePath, 'utf8')
);

const solutionJson = JSON.parse(
  fs.readFileSync(solutionPath, 'utf8')
);

const semver = packageJson.version.split('-')[0];
const parts = semver.split('.');

if (parts.length !== 3) {
  throw new Error(
    `Versión inválida en package.json: ${packageJson.version}`
  );
}

solutionJson.solution.version =
  `${parts[0]}.${parts[1]}.${parts[2]}.0`;

fs.writeFileSync(
  solutionPath,
  JSON.stringify(solutionJson, null, 2) + '\n'
);

console.log(
  `SPFx sincronizado: ${solutionJson.solution.version}`
);