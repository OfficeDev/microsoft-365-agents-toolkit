import fs from "fs-extra";
import path from "path";
const pvtFile = path.join(__dirname, "./pvt-vscuse.json");
const casesFile = path.join(__dirname, "./randomCasesVSCUSE.json");
const pvtCases: TestMatrix = fs.readJSONSync(pvtFile);
const casesIndex = fs.readJSONSync(casesFile);

type TestCase = string;

interface TestMatrix {
  [os: string]: {
    [nodeVersion: string]: TestCase[];
  };
}

function addRandomTestCase(testMatrix: TestMatrix, cases: TestCase[]): void {
  for (let i = 0; i < cases.length; i++) {
    const osVersions = Object.keys(testMatrix);
    const randomOS = osVersions[Math.floor(Math.random() * osVersions.length)];
    const nodeVersions = Object.keys(testMatrix[randomOS]);
    const randomNode =
      nodeVersions[Math.floor(Math.random() * nodeVersions.length)];
    const randomCase = cases[i];
    testMatrix[randomOS][randomNode].push(randomCase);
  }
}

function main() {
  for (let i = 0; i < casesIndex.length; i++) {
    const testMatrix = casesIndex[i]["os"];
    const cases = casesIndex[i]["cases"];
    addRandomTestCase(testMatrix, cases);
    // save testMatrix to pvtCases
    Object.keys(testMatrix).forEach((os) => {
      Object.keys(testMatrix[os]).forEach((node) => {
        pvtCases[os][node] = [...pvtCases[os][node], ...testMatrix[os][node]];
      });
    });
  }

  fs.writeJsonSync(path.join(__dirname, "./pvt-vscuse.json"), pvtCases, { spaces: 2 });
}

main();
