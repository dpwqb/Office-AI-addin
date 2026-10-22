const path = require('path');

module.exports = {
  mode: 'production',
  target: ['web', 'es5'],
  entry: path.resolve(__dirname, 'compat/entry.js'),
  output: {
    path: path.resolve(__dirname, 'dist/assets/js'),
    filename: 'office-compat.bundle.js',
    clean: true,
    environment: {
      arrowFunction: false,
      const: false,
      destructuring: false,
      dynamicImport: false,
      forOf: false,
      module: false,
      optionalChaining: false,
      templateLiteral: false
    }
  },
  devtool: 'source-map',
  module: {
    rules: [{
      test: /\.js$/,
      exclude: /node_modules/,
      use: {
        loader: 'babel-loader',
        options: {
          babelrc: false,
          configFile: false,
          presets: [[
            '@babel/preset-env',
            {
              targets: { ie: '11' },
              bugfixes: true,
              modules: false,
              useBuiltIns: false
            }
          ]]
        }
      }
    }]
  },
  performance: { hints: false }
};
