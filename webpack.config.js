var webpack = require('webpack');

module.exports = {
  entry: './kurve.ts',
  output: {
    library: "kurve",
    libraryTarget: "amd",
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
