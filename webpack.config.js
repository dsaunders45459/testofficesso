const path = require('path');

module.exports = {
  entry: './src/index.js',
  output: {
    filename: 'javascripts/main.js',
    path: path.resolve(__dirname, 'dist')
  },
  mode: "development"
};