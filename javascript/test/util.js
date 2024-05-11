// @ts-check
const path = require('path');
const fs = require('fs');

/**
 *
 * @param {string} currentPath
 * @returns {string | null}
 *
 */
function findRootDir(currentPath) {
  if (fs.existsSync(path.join(currentPath, 'package.json'))) {
    return currentPath;
  } else {
    const parentPath = path.resolve(currentPath, '..');
    if (parentPath !== currentPath) {
      return findRootDir(parentPath);
    }
  }
  return null;
}

module.exports = findRootDir;
