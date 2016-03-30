var webpack = require('webpack');

module.exports = {
  entry: './src/kurve.ts',
  output: {
    library: "Kurve",
    libraryTarget: "umd",
    filename: 'dist/kurve.js'
  },
  devtool: 'source-map',
  plugins: [],
  resolve: {
    extensions: ['', '.webpack.js', '.web.js', '.ts']
  },
  module: {
    loaders: [
      { loader: 'ts-loader' }
    ]
  }
}
