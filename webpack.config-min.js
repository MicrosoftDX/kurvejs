var webpack = require('webpack');

module.exports = {
  entry: './src/kurve.ts',
  output: {
    library: "Kurve",
    libraryTarget: "umd",
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
