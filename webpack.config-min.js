var webpack = require('webpack');

module.exports = {
  entry: './kurve.ts',
  output: {
    library: "kurve",
    libraryTarget: "amd",
    filename: 'dist/kurve.min.js'
  },
  devtool: 'source-map',
  plugins: [ new webpack.optimize.UglifyJsPlugin({minimize: true}) ],
  resolve: {
    extensions: ['', '.webpack.js', '.web.js', '.ts']
  },
  module: {
    loaders: [
      { loader: 'ts-loader' }
    ]
  }
}
